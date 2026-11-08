import {readFile,writeFile,mkdir,mkdtemp,readdir,lstat,cp,copyFile,rename,rm,stat,access} from 'node:fs/promises';
import {createReadStream,createWriteStream} from 'node:fs';
import {pipeline} from 'node:stream/promises';
import {Readable} from 'node:stream';
import {createHash} from 'node:crypto';
import {execFile} from 'node:child_process';
import {promisify} from 'node:util';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
const root=fileURLToPath(new URL('../',import.meta.url));
const exec=promisify(execFile);
export async function fileHash(file) {const hash=createHash('sha256');for await(const data of createReadStream(file))hash.update(data);return hash.digest('hex');}
export async function runtimeInventory(folder) {
    const files=[];
    async function visit(relative='') {
        for(const entry of (await readdir(path.join(folder,relative),{withFileTypes:true})).sort((a,b)=>a.name<b.name?-1:a.name>b.name?1:0)) {
            const name=relative?relative+'/'+entry.name:entry.name;
            if(entry.isSymbolicLink())throw Error('Link simbólico no motor: '+name);
            if(entry.isDirectory())await visit(name);
            else if(entry.isFile())files.push({path:name,size:(await stat(path.join(folder,name))).size,sha256:await fileHash(path.join(folder,name))});
            else throw Error('Arquivo especial no motor: '+name);
        }
    }
    await visit();return files;
}
async function download(url,destination,sha256,allowDownload) {
    try{if(await fileHash(destination)===sha256)return;throw Error('Arquivo em cache difere do SHA256: '+destination);}catch(error){if(error.code!=='ENOENT')throw error;}
    if(!allowDownload)throw Error('Falta '+destination+'. Execute o preparador com --download para obter a distribuição oficial e suas fontes.');
    const temp=destination+'.part';await mkdir(path.dirname(destination),{recursive:true});
    try {
        const response=await fetch(url);if(!response.ok||!response.body)throw Error('Download falhou: '+response.status+' '+url);
        await pipeline(Readable.fromWeb(response.body),createWriteStream(temp,{flags:'wx'}));
        if(await fileHash(temp)!==sha256)throw Error('SHA256 incorreto: '+url);
        await rename(temp,destination);
    }finally{await rm(temp,{force:true});}
}

export async function verifyPreparedLibreOffice(projectRoot=root) {
    const base=path.join(projectRoot,'build/libreoffice'),lock=JSON.parse(await readFile(path.join(projectRoot,'build/libreoffice-lock.json'),'utf8'));
    const trust=JSON.parse(await readFile(path.join(base,'trust.json'),'utf8'));
    const manifestFile=path.join(base,'runtime/runtime-manifest.json');
    if(await fileHash(manifestFile)!==trust.manifestSha256)throw Error('Manifesto do motor privado divergente. Execute o preparador.');
    const manifest=JSON.parse(await readFile(manifestFile,'utf8'));
    if(manifest.version!==lock.version||manifest.msiSha256!==lock.msi.sha256||trust.version!==lock.version)throw Error('Motor privado fora da versão fixada.');
    const actual=await runtimeInventory(path.join(base,'runtime'));
    const expected=manifest.files;
    const content=actual.filter(file=>file.path!=='runtime-manifest.json');
    if(JSON.stringify(content)!==JSON.stringify(expected))throw Error('Arquivos do motor privado ausentes, extras ou alterados. Execute o preparador.');
    for(const source of lock.sourceArchives)if(await fileHash(path.join(projectRoot,'build/libreoffice-cache',source.name))!==source.sha256)throw Error('Fonte correspondente ausente/divergente: '+source.name);
    return {version:lock.version,files:content.length,bytes:content.reduce((n,file)=>n+file.size,0),manifestSha256:trust.manifestSha256};
}

export async function prepareLibreOffice({projectRoot=root,msiFile=process.env.PRINTSHOPS_LIBREOFFICE_MSI,allowDownload=false}={}) {
    if(process.platform!=='win32'||process.arch!=='x64')throw Error('Preparador requer Windows x64.');
    const lock=JSON.parse(await readFile(path.join(projectRoot,'build/libreoffice-lock.json'),'utf8'));
    const base=path.join(projectRoot,'build/libreoffice'),cache=path.join(projectRoot,'build/libreoffice-cache');
    await mkdir(base,{recursive:true});await mkdir(cache,{recursive:true});
    const msi=msiFile||path.join(cache,lock.msi.name);
    if(!path.isAbsolute(msi))throw Error('PRINTSHOPS_LIBREOFFICE_MSI deve ser absoluto.');
    if(msiFile){if(await fileHash(msi)!==lock.msi.sha256)throw Error('MSI fornecido difere da distribuição oficial fixada.');}
    else await download(lock.msi.url,msi,lock.msi.sha256,allowDownload);
    for(const source of lock.sourceArchives)await download(lock.sourceBaseURL+source.name,path.join(cache,source.name),source.sha256,allowDownload);
    try {return await verifyPreparedLibreOffice(projectRoot);}catch(error){if(error.code!=='ENOENT')throw error;}
    const stage=await mkdtemp(path.join(base,'stage-')),image=path.join(stage,'image'),runtime=path.join(stage,'runtime');
    try {
        await mkdir(image);await mkdir(runtime);
        const quote=value=>{if(/["\x00\r\n]/.test(value))throw Error('Caminho MSI inválido.');return '"'+value+'"';};
        await exec(path.join(process.env.SystemRoot,'System32/msiexec.exe'),['/a',quote(msi),'/qn','TARGETDIR='+quote(image),'/L*v',quote(path.join(base,'extraction.log'))],{windowsHide:true,windowsVerbatimArguments:true,timeout:180000});
        const version=await readFile(path.join(image,'program/version.ini'),'utf8');
        if(!version.includes('buildid='+lock.buildId))throw Error('Build do LibreOffice inesperado.');
        for(const directory of lock.directories)await cp(path.join(image,directory),path.join(runtime,directory),{recursive:true,errorOnExist:true,force:false});
        for(const file of lock.rootFiles)await copyFile(path.join(image,file),path.join(runtime,file));
        await cp(path.join(image,'Fonts'),path.join(runtime,'share/fonts/truetype'),{recursive:true,errorOnExist:true,force:false});
        for(const entry of await readdir(path.join(image,'System64'),{withFileTypes:true})){
            if(!entry.isFile()||!entry.name.endsWith('.dll'))throw Error('Runtime x64 inesperado.');
            await copyFile(path.join(image,'System64',entry.name),path.join(runtime,'program',entry.name));
        }
        await copyFile(path.join(projectRoot,'third-party/libreoffice/NOTICE.txt'),path.join(runtime,'4PRINTSHOPS-NOTICE.txt'));
        await writeFile(path.join(runtime,'SOURCE-CODE.json'),JSON.stringify({version:lock.version,buildId:lock.buildId,archives:lock.sourceArchives.map(source=>({...source,url:lock.sourceBaseURL+source.name})),git:lock.sourceGitURL},null,2)+'\n');
        const files=await runtimeInventory(runtime);
        const manifest={schema:1,version:lock.version,buildId:lock.buildId,msiSha256:lock.msi.sha256,files};
        await writeFile(path.join(runtime,'runtime-manifest.json'),JSON.stringify(manifest)+'\n');
        const trust={schema:1,version:lock.version,manifestSha256:await fileHash(path.join(runtime,'runtime-manifest.json'))};
        // Nunca substituir uma árvore existente sem verificar o motivo.
        try{await access(path.join(base,'runtime'));throw Error('Motor anterior existente; revisar antes de substituir.');}catch(error){if(error.code!=='ENOENT')throw error;}
        await rename(runtime,path.join(base,'runtime'));
        await writeFile(path.join(base,'trust.json'),JSON.stringify(trust)+'\n');
        return await verifyPreparedLibreOffice(projectRoot);
    }finally {
        // Somente a pasta exclusiva criada acima, dentro do diretório de build.
        if(path.dirname(stage)!==base||!path.basename(stage).startsWith('stage-'))throw Error('Diretório temporário inesperado.');
        await rm(stage,{recursive:true,force:true});
    }
}
if(process.argv[1]&&path.resolve(process.argv[1])===fileURLToPath(import.meta.url))console.log(JSON.stringify(await prepareLibreOffice({allowDownload:process.argv.includes('--download')})));
