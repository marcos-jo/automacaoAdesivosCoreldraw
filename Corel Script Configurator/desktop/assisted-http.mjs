import http from 'node:http';
import https from 'node:https';
import {createHash,randomUUID} from 'node:crypto';
import {createReadStream,createWriteStream} from 'node:fs';
import {unlink,rename,stat,mkdir} from 'node:fs/promises';
import {pipeline} from 'node:stream/promises';
import {Transform} from 'node:stream';
import path from 'node:path';

function response(url,{timeout=15000}={}){
    return new Promise((resolve,reject)=>{
        const request=(url.protocol==='https:'?https:http).get(url,{headers:{'Accept-Encoding':'identity'}},res=>{
            if(res.statusCode!==200||res.headers['content-encoding']&&res.headers['content-encoding']!=='identity'){res.destroy();reject(new Error('Servidor recusou o arquivo ou respondeu com redirecionamento/codificação inesperada.'));return;}
            resolve(res);
        });
        const timer=setTimeout(()=>request.destroy(new Error('Servidor indisponível ou prazo de transferência excedido.')),timeout);
        request.on('error',reject);request.on('close',()=>clearTimeout(timer));
    });
}
export async function getManifest(url,options){
    const res=await response(url,options);const chunks=[];let size=0;
    try{for await(const chunk of res){size+=chunk.length;if(size>65536)throw new Error('Manifesto excede o limite permitido.');chunks.push(chunk);}return Buffer.concat(chunks);}
    finally{res.destroy();}
}
export async function verifyCached(file,artifact){
    try{const info=await stat(file);if(!info.isFile()||info.size!==artifact.size)return false;const hash=createHash('sha512');for await(const chunk of createReadStream(file))hash.update(chunk);return hash.digest('hex')===artifact.sha512;}catch{return false;}
}
export async function downloadInstaller(url,artifact,folder,onProgress=()=>{},{timeout=300000}={}){
    await mkdir(folder,{recursive:true});
    const target=path.join(folder,artifact.sha512+'.exe');
    if(await verifyCached(target,artifact)){onProgress(100);return target;}
    const temp=path.join(folder,randomUUID()+'.part');let size=0;const hash=createHash('sha512');
    try{
        const res=await response(url,{timeout});
        if(res.headers['content-length']&&Number(res.headers['content-length'])!==artifact.size){res.destroy();throw new Error('Tamanho do download diverge da publicação assinada.');}
        const meter=new Transform({transform(chunk,_encoding,done){size+=chunk.length;if(size>artifact.size){done(new Error('Instalador excedeu o tamanho autorizado.'));return;}hash.update(chunk);onProgress(Math.floor(size*100/artifact.size));done(null,chunk);}});
        await pipeline(res,meter,createWriteStream(temp,{flags:'wx',mode:0o600}));
        if(size!==artifact.size||hash.digest('hex')!==artifact.sha512)throw new Error('Integridade do instalador inválida. Download descartado.');
        await rename(temp,target);return target;
    }finally{await unlink(temp).catch(()=>{});}
}
