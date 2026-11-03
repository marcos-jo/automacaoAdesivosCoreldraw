import path from 'node:path';
import {open} from 'node:fs/promises';
import {randomUUID} from 'node:crypto';

export const IMPORT_LIMIT=512*1024*1024;
export function isPDFPath(value){
    if(typeof value!=='string'||value.length>32760||/[\x00-\x1f]/.test(value))return false;
    let name=value;
    if(name.startsWith('\\\\?\\UNC\\'))name='\\\\'+name.slice(8);
    else if(name.startsWith('\\\\?\\'))name=name.slice(4);
    if(!/^[a-z]:[\\/]/i.test(name)&&!/^\\\\[^\\/.?]+\\[^\\/]+\\/.test(name))return false;
    const rest=/^[a-z]:/i.test(name)?name.slice(2):name;
    return !/[<>:"|?*]/.test(rest)&&path.win32.extname(name).toLowerCase()==='.pdf';
}

export function pdfArguments(argv,{packaged=true}={}){
    const files=[];let skipValue=false;
    for(const value of argv.slice(packaged?1:2)){
        if(value==='--open-pdf'||value==='--'){skipValue=false;continue;}
        if(value.startsWith('-')){skipValue=!value.includes('=');continue;}
        if(skipValue){skipValue=false;continue;}
        if(isPDFPath(value))files.push(value);
    }
    return [...new Map(files.map(f=>[path.win32.normalize(f).toLowerCase(),f])).values()];
}

export async function readExternalPDFs(paths){
    if(!Array.isArray(paths)||!paths.length||paths.some(f=>!isPDFPath(f)))throw new Error('Selecione caminhos de arquivos PDF locais ou de rede válidos.');
    const result=[];let total=0;
    for(const file of paths){
        let handle;
        try{
            handle=await open(file,'r');const info=await handle.stat();
            if(!info.isFile())throw new Error('O caminho não é um arquivo.');
            if(info.size<8)throw new Error('Arquivo PDF vazio ou inválido.');
            if(total+info.size>IMPORT_LIMIT)throw new Error('A seleção ultrapassa o limite de 512 MiB.');
            const bytes=Buffer.alloc(info.size);let offset=0;
            while(offset<bytes.length){const part=await handle.read(bytes,offset,bytes.length-offset,offset);if(!part.bytesRead)break;offset+=part.bytesRead;}
            const tail=await handle.read(Buffer.alloc(1),0,1,offset);
            if(offset!==bytes.length||tail.bytesRead||(await handle.stat()).mtimeMs!==info.mtimeMs)throw new Error('O arquivo mudou durante a leitura. Tente novamente.');
            if(!bytes.subarray(0,1024).includes(Buffer.from('%PDF-')))throw new Error('O arquivo não contém um cabeçalho PDF válido.');
            total+=bytes.length;result.push({name:path.win32.basename(file),bytes});
        }catch(error){throw new Error(`Não foi possível abrir "${path.win32.basename(file)}": ${error.code==='ENOENT'?'arquivo não encontrado':error.code==='EACCES'?'acesso negado':error.message}`);}
        finally{await handle?.close().catch(()=>{});}
    }
    return result;
}

export class ExternalPDFQueue{
    constructor(notify=()=>{},reader=readExternalPDFs){this.items=new Map();this.notify=notify;this.reader=reader;this.active=false;this.reading=false;}
    add(paths){
        if(!Array.isArray(paths)||paths.length>128||paths.some(p=>!isPDFPath(p)))throw new Error('Solicitação de abertura inválida (máximo de 128 PDFs pendentes).');
        const existing=new Set([...this.items.values()].map(p=>path.win32.normalize(p).toLowerCase()));
        const fresh=[];for(const p of paths){const key=path.win32.normalize(p).toLowerCase();if(!existing.has(key)){fresh.push(p);existing.add(key);}}
        if(this.items.size+fresh.length>128)throw new Error('Há mais de 128 PDFs pendentes. Conclua ou dispense as aberturas atuais.');
        for(const p of fresh)this.items.set(randomUUID(),p);this.notify();
    }
    list(){return [...this.items].map(([id,p])=>({id,name:path.win32.basename(p)}));}
    select(ids){if(!Array.isArray(ids)||!ids.length||ids.length>128||new Set(ids).size!==ids.length||ids.some(id=>typeof id!=='string'||!this.items.has(id)))throw new Error('Solicitação de arquivo expirada ou inválida.');return ids.map(id=>this.items.get(id));}
    async read(ids){
        const paths=this.select(ids);if(this.active||this.reading)throw new Error('Aguarde a leitura atual terminar.');
        this.active=this.reading=true;
        let timer;const reading=this.reader(paths).finally(()=>{this.reading=false;});
        try{return await Promise.race([reading,new Promise((_resolve,reject)=>{timer=setTimeout(()=>reject(new Error('A leitura excedeu 30 segundos. Verifique a conexão e as permissões do arquivo.')),30000);})]);}
        catch(error){this.active=false;throw error;}finally{clearTimeout(timer);}
    }
    finish(ids){this.select(ids);for(const id of ids)this.items.delete(id);this.active=false;this.notify();}
    release(ids){this.select(ids);this.active=false;}
}
