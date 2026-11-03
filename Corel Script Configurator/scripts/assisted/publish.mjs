import {readFile,writeFile,mkdir,rename,unlink,lstat,readdir,realpath,copyFile} from 'node:fs/promises';
import {createHash,createPublicKey,sign,randomUUID} from 'node:crypto';
import {createReadStream,constants} from 'node:fs';
import {createRequire} from 'node:module';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {authenticateEnvelope,validatePolicy,validateRelease,keyId,LAB_PRODUCT,MAIN_PRODUCT,MAX_INSTALLER} from '../../desktop/assisted-manifest.mjs';
import {loadPublishingKey} from './keys.mjs';
const require=createRequire(import.meta.url);

export async function auditRepository(root){
    const info=await lstat(root);if(!info.isDirectory()||info.isSymbolicLink())throw new Error('Repositório não pode ser link/reparse point.');
    for(const entry of await readdir(root,{withFileTypes:true})){
        const file=path.join(root,entry.name),stat=await lstat(file);
        if(stat.isSymbolicLink()||!stat.isDirectory()&&!stat.isFile())throw new Error('Link ou arquivo especial no repositório: '+entry.name);
        if(stat.isDirectory())await auditRepository(file);
        else if(/(?:private|secret|password|senha)|\.(?:dpapi|key|pem|pfx|p12|pdf|log|tmp)$/i.test(entry.name))throw new Error('Arquivo não publicável no repositório: '+entry.name);
    }
}
async function digest(file){const hash=createHash('sha512');for await(const chunk of createReadStream(file))hash.update(chunk);return hash.digest('hex');}
export async function publishRelease(options){
    await mkdir(options.root,{recursive:true});await auditRepository(options.root);
    const lock=path.join(options.root,'.publish.lock');let handle;
    try{handle=await import('node:fs/promises').then(fs=>fs.open(lock,'wx'));}
    catch(error){if(error.code==='EEXIST')throw new Error('Outra publicação está em andamento (ou deixou trava após interrupção).');throw error;}
    try{return await publishLocked(options);}finally{await handle.close();await unlink(lock);}
}
async function publishLocked({root,installer,policy,privateKey,version,notes,sequence,days=14,now=Date.now()}){
    validatePolicy(policy);
    const publicPem=createPublicKey(privateKey).export({type:'spki',format:'pem'}),id=keyId(publicPem);
    if(!policy.keys.some(k=>k.id===id))throw new Error('Chave privada não corresponde à raiz pública.');
    await mkdir(root,{recursive:true});await auditRepository(root);
    const channel=path.join(root,'channels',policy.channel+'.json');let previous;
    try{previous=authenticateEnvelope(await readFile(channel),policy);}catch(error){if(error.code!=='ENOENT')throw error;}
    const info=await lstat(installer);if(!info.isFile()||info.isSymbolicLink()||info.size<8||info.size>MAX_INSTALLER)throw new Error('Instalador inválido ou maior que 512 MiB.');
    const filename=path.basename(installer),relative=`releases/${version}/${filename}`;
    const data={schema:1,product:policy.product,channel:policy.channel,platform:'win32',arch:'x64',version,sequence:sequence??(previous?.data.sequence||0)+1,
        issuedAt:new Date(now).toISOString(),expiresAt:new Date(now+days*86400000).toISOString(),notes,
        installer:{path:relative,size:info.size,sha512:await digest(installer)}};
    const payload=Buffer.from(JSON.stringify(data)),envelope=Buffer.from(JSON.stringify({keyId:id,payload:payload.toString('base64url'),signature:sign(null,payload,privateKey).toString('base64url')})+'\n');
    const authenticated=authenticateEnvelope(envelope,policy);
    validateRelease(authenticated,{currentVersion:previous?.data.version||'0.0.0',floor:previous?{sequence:previous.data.sequence,digest:previous.digest}:undefined,now});
    if(previous&&data.sequence<=previous.data.sequence)throw new Error('Cada publicação exige uma sequência nova.');
    const target=path.resolve(root,relative);await mkdir(path.dirname(target),{recursive:true});
    // Não permite que um diretório de release preexistente desvie a cópia por links.
    await auditRepository(root);
    try{await copyFile(installer,target,constants.COPYFILE_EXCL);}catch(error){if(error.code!=='EEXIST')throw error;if(await digest(target)!==data.installer.sha512)throw new Error('Versão já publicada com bytes diferentes. Não sobrescrever.');}
    if(await digest(target)!==data.installer.sha512)throw new Error('Instalador mudou durante a publicação. Manifesto não publicado.');
    await mkdir(path.dirname(channel),{recursive:true});
    const temp=path.join(path.dirname(channel),randomUUID()+'.json.part');
    try{await writeFile(temp,envelope,{flag:'wx'});await rename(temp,channel);}finally{await unlink(temp).catch(()=>{});}
    return {version,sequence:data.sequence,manifest:channel,installer:target,sha512:data.installer.sha512};
}
export async function packagedIdentity(build,policy){
    const asar=require('@electron/asar'),file=path.join(build,'win-unpacked','resources','app.asar');
    const pkg=JSON.parse(asar.extractFile(file,'package.json'));
    if(JSON.stringify(pkg.assistedUpdates)!==JSON.stringify(policy))throw new Error('O pacote não incorpora exatamente a raiz pública aprovada.');
    const lab=policy.product===LAB_PRODUCT;
    if(lab?!(pkg.name==='4printshops-assisted-lab'&&pkg.desktopAssistedLab===true):policy.product!==MAIN_PRODUCT||pkg.name!=='corel-script-configurator'||pkg.desktopAssistedLab||pkg.desktopUpdateLab||pkg.desktopOpenTest)throw new Error('Identidade de laboratório/produção incompatível.');
    const name=lab?`4PrintShops-Assisted-Lab-${pkg.version}-Windows-x64-Setup.exe`:`4PrintShops-${pkg.version}-Windows-x64-Setup.exe`;
    return {version:pkg.version,installer:path.join(build,name)};
}
if(process.argv[1]&&path.resolve(process.argv[1])===fileURLToPath(import.meta.url)){
    const [root,build,trust,key,notesFile]=process.argv.slice(2);
    if(!notesFile)throw new Error('Uso: node scripts/assisted/publish.mjs <repositorio> <pasta-build> <public-trust.json> <publication.dpapi> <notas.txt>');
    const policy=validatePolicy(JSON.parse(await readFile(trust,'utf8'))),identity=await packagedIdentity(path.resolve(build),policy);
    const keyFolder=await realpath(path.dirname(key)),repo=path.resolve(root);
    if(!path.relative(repo,keyFolder).startsWith('..')&&!path.isAbsolute(path.relative(repo,keyFolder)))throw new Error('Chave privada deve ficar fora da distribuição.');
    const notes=(await readFile(notesFile,'utf8')).split(/\r?\n/).map(s=>s.trim()).filter(Boolean);
    console.log(JSON.stringify(await publishRelease({root:repo,...identity,policy,privateKey:await loadPublishingKey(key),notes}),null,2));
}
