import {readFile,lstat,realpath,readdir} from 'node:fs/promises';
import {createReadStream} from 'node:fs';
import {createHash} from 'node:crypto';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {findLibreOffice} from './word-import.mjs';
const defaultRoot=fileURLToPath(new URL('../',import.meta.url));
async function hash(file){const value=createHash('sha256');for await(const chunk of createReadStream(file))value.update(chunk);return value.digest('hex');}
export async function verifyOfficeRuntime(root,trustFile) {
    const rootInfo=await lstat(root);if(!rootInfo.isDirectory()||rootInfo.isSymbolicLink())throw Error('Pasta do conversor interno inválida.');
    const trust=JSON.parse(await readFile(trustFile,'utf8'));
    const manifestFile=path.join(root,'runtime-manifest.json');
    if(!/^[a-f0-9]{64}$/.test(trust.manifestSha256)||await hash(manifestFile)!==trust.manifestSha256)throw Error('O manifesto do conversor interno está ausente ou foi alterado. Reinstale o 4PrintShops.');
    const manifest=JSON.parse(await readFile(manifestFile,'utf8'));
    if(manifest.schema!==1||manifest.version!==trust.version||!Array.isArray(manifest.files)||!manifest.files.length||manifest.files.length>100000)throw Error('Manifesto do conversor interno inválido.');
    const absolute=await realpath(root),seen=new Set(),targets=[];
    for(const file of manifest.files) {
        if(typeof file.path!=='string'||file.path.includes('\\')||file.path.includes(':')||file.path.startsWith('/')||file.path.split('/').some(p=>!p||p==='.'||p==='..')||seen.has(file.path)||!/^[a-f0-9]{64}$/.test(file.sha256))throw Error('Caminho inválido no conversor interno.');
        seen.add(file.path);
        targets.push(file);
    }
    for(const needed of ['program/soffice.com','program/soffice.bin','program/mergedlo.dll','program/vcruntime140.dll','program/vcruntime140_1.dll','program/msvcp140.dll','LICENSE.html','license.txt','NOTICE','4PRINTSHOPS-NOTICE.txt','SOURCE-CODE.json'])if(!seen.has(needed))throw Error('Componente obrigatório ausente no conversor interno: '+needed);
    let cursor=0,failed=false;
    async function worker() {
        while(!failed&&cursor<targets.length) {
            const file=targets[cursor++];
            try {
                const target=path.join(root,file.path),info=await lstat(target),resolved=await realpath(target),relative=path.relative(absolute,resolved);
                if(!relative||relative.startsWith('..')||path.isAbsolute(relative)||!info.isFile()||info.isSymbolicLink()||info.size!==file.size||await hash(target)!==file.sha256)throw Error('O conversor interno está incompleto ou foi alterado: '+file.path+'. Reinstale o 4PrintShops.');
            }catch(error){failed=true;throw error;}
        }
    }
    // Limitar leitura concorrente evita milhares de tarefas/arquivos abertos.
    const checked=await Promise.allSettled(Array.from({length:8},()=>worker()));
    const error=checked.find(result=>result.status==='rejected');if(error)throw error.reason;
    async function checkExtras(relative='') {
        for(const entry of await readdir(path.join(root,relative),{withFileTypes:true})) {
            const name=relative?relative+'/'+entry.name:entry.name;
            if(entry.isSymbolicLink())throw Error('Link inesperado no conversor interno.');
            if(entry.isDirectory())await checkExtras(name);
            else if(!entry.isFile()||!seen.has(name)&&name!=='runtime-manifest.json')throw Error('Arquivo não autorizado no conversor interno: '+name);
        }
    }
    await checkExtras();
    return {executable:path.join(root,'program/soffice.com'),version:manifest.version,files:manifest.files.length};
}

export function privateOfficeEnvironment(root,environment=process.env) {
    const env={...environment};
    for(const key of Object.keys(env))if(/^(PATH|URE_BOOTSTRAP|UNO_PATH|UNO_TYPES|UNO_SERVICES|PYTHONHOME|PYTHONPATH|PYTHONPYCACHEPREFIX|SAL_FONTPATH_PRIVATE)$/i.test(key))delete env[key];
    env.PYTHONDONTWRITEBYTECODE='1';
    env.SAL_FONTPATH_PRIVATE=path.join(root,'share/fonts/truetype');
    // Somente o processo de conversão: sem procurar DLLs em outros Offices/PATH.
    env.PATH=path.join(environment.SystemRoot||'C:\\Windows','System32');
    return env;
}

export function createOfficeResolver({packaged=false,resourcesPath,projectRoot=defaultRoot,environment=process.env}={}) {
    const root=packaged?path.join(resourcesPath,'libreoffice'):path.join(projectRoot,'build/libreoffice/runtime');
    const trustFile=packaged?fileURLToPath(new URL('../build/libreoffice/trust.json',import.meta.url)):path.join(projectRoot,'build/libreoffice/trust.json');
    let verification;
    return async()=>{
        let exists;
        try{await lstat(root);exists=true;}catch(error){if(error.code!=='ENOENT')throw error;exists=false;}
        if(!exists) {
            if(packaged)throw Error('O conversor interno não foi encontrado. Reinstale o 4PrintShops para importar Word.');
            return {executable:await findLibreOffice(environment),mode:'system-development'};
        }
        // Uma verificação integral por processo; não há fallback após erro.
        verification??=verifyOfficeRuntime(root,trustFile);
        try{return {...await verification,mode:'bundled',env:privateOfficeEnvironment(root,environment)};}
        catch(error){verification=null;throw error;}
    };
}
