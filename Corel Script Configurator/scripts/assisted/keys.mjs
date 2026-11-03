import {generateKeyPairSync,createPrivateKey} from 'node:crypto';
import {mkdir,writeFile,readFile,lstat,realpath} from 'node:fs/promises';
import {spawn} from 'node:child_process';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {existsSync} from 'node:fs';
import {keyId,validatePolicy,LAB_PRODUCT,MAIN_PRODUCT} from '../../desktop/assisted-manifest.mjs';

const workspace=fileURLToPath(new URL('../../',import.meta.url));
let repository=path.resolve(workspace);
while(!existsSync(path.join(repository,'.git'))&&path.dirname(repository)!==repository)repository=path.dirname(repository);
const powershell=path.join(process.env.SystemRoot||'C:\\Windows','System32','WindowsPowerShell','v1.0','powershell.exe');
function processText(exe,args,input=''){
    return new Promise((resolve,reject)=>{
        const child=spawn(exe,args,{windowsHide:true,stdio:['pipe','pipe','pipe']});let output='';
        const timer=setTimeout(()=>child.kill(),15000);
        child.stdout.on('data',b=>{output+=b;if(output.length>16384)child.kill();});child.stderr.resume();
        child.on('error',reject);child.on('close',code=>{clearTimeout(timer);code===0?resolve(output.trim()):reject(new Error('Proteção local de chave falhou; nenhum segredo deve ser copiado para o servidor.'));});
        child.stdin.on('error',()=>{});child.stdin.end(input);
    });
}
async function dpapi(bytes,protect){
    if(process.platform!=='win32')throw new Error('Chaves de publicação exigem Windows DPAPI neste procedimento.');
    const script=`$ErrorActionPreference='Stop'; Add-Type -AssemblyName System.Security; $b=[Convert]::FromBase64String([Console]::In.ReadToEnd()); $e=[Text.Encoding]::UTF8.GetBytes('4PrintShops-Assisted-Manifest-v1'); $r=[Security.Cryptography.ProtectedData]::${protect?'Protect':'Unprotect'}($b,$e,[Security.Cryptography.DataProtectionScope]::CurrentUser); [Console]::Out.Write([Convert]::ToBase64String($r))`;
    return Buffer.from(await processText(powershell,['-NoProfile','-NonInteractive','-Command',script],bytes.toString('base64')),'base64');
}
export async function outsideWorkspace(folder){
    const resolved=await realpath(folder),relative=path.relative(await realpath(repository),resolved);
    if(!relative||!relative.startsWith('..')&&!path.isAbsolute(relative)||resolved.startsWith('\\\\')||!path.isAbsolute(resolved))throw new Error('A chave deve ficar em pasta local fora do repositório e do servidor.');
    return resolved;
}
export async function generatePublishingKey(folder,{laboratory=false}={}){
    const absolute=path.resolve(folder),drive=path.parse(absolute).root;
    if(process.platform!=='win32'||!/^\w:\\$/.test(drive)||await processText(powershell,['-NoProfile','-NonInteractive','-Command',`([IO.DriveInfo]::new('${drive}')).DriveType.ToString()`])!=='Fixed')throw new Error('A chave exige um disco local fixo; não use compartilhamento ou unidade de rede.');
    await mkdir(absolute,{recursive:false});
    await outsideWorkspace(absolute);
    // ACL somente da NOVA pasta no computador publicador. Não altera o Windows/servidor.
    const sid=await processText(powershell,['-NoProfile','-NonInteractive','-Command','[Security.Principal.WindowsIdentity]::GetCurrent().User.Value']);
    if(!/^S-1-5-\d+(?:-\d+)+$/.test(sid))throw new Error('Não foi possível identificar o publicador.');
    await processText(path.join(process.env.SystemRoot||'C:\\Windows','System32','icacls.exe'),[absolute,'/inheritance:r','/grant:r',`*${sid}:(OI)(CI)F`,'*S-1-5-18:(OI)(CI)F']);
    const {privateKey,publicKey}=generateKeyPairSync('ed25519'),der=privateKey.export({format:'der',type:'pkcs8'});
    try{
        const protectedKey=await dpapi(der,true),pem=publicKey.export({format:'pem',type:'spki'});
        const policy=validatePolicy({product:laboratory?LAB_PRODUCT:MAIN_PRODUCT,channel:laboratory?'internal-lab':'internal',keys:[{id:keyId(pem),pem}]});
        await writeFile(path.join(absolute,'publication.dpapi'),protectedKey,{flag:'wx',mode:0o600});
        await writeFile(path.join(absolute,'public-trust.json'),JSON.stringify(policy,null,2)+'\n',{flag:'wx'});
        return {folder:absolute,keyId:policy.keys[0].id,publicTrust:path.join(absolute,'public-trust.json')};
    }finally{der.fill(0);}
}
export async function loadPublishingKey(file){
    await outsideWorkspace(path.dirname(path.resolve(file)));const info=await lstat(file);
    if(!info.isFile()||info.isSymbolicLink()||info.size>8192)throw new Error('Chave protegida inválida.');
    const der=await dpapi(await readFile(file),false);
    try{return createPrivateKey({key:der,type:'pkcs8',format:'der'});}finally{der.fill(0);}
}
if(process.argv[1]&&path.resolve(process.argv[1])===fileURLToPath(import.meta.url)){
    if(!process.argv[2])throw new Error('Uso: node scripts/assisted/keys.mjs <NOVA-pasta-local-fora-do-repositorio> [--lab]');
    console.log(JSON.stringify(await generatePublishingKey(process.argv[2],{laboratory:process.argv.includes('--lab')}),null,2));
}
