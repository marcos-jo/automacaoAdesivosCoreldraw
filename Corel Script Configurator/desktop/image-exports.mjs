import path from 'node:path';
import {mkdir,open,unlink,rmdir} from 'node:fs/promises';
import {createReadStream} from 'node:fs';
import {randomUUID} from 'node:crypto';
import {pdfName} from './security.mjs';

export function validateImagePlan(plan) {
    if (!plan || !['png','jpeg'].includes(plan.format) || !Array.isArray(plan.names) || !plan.names.length || plan.names.length>1000 || new Set(plan.names).size!==plan.names.length) throw Error('Plano de imagens inválido ou acima de 1000 páginas.');
    pdfName(plan.base+'.pdf');
    if (plan.base.length>90) throw Error('Use um nome de até 90 caracteres.');
    for(const name of plan.names){
        if(typeof name!=='string'||!name.endsWith(plan.format==='png'?'.png':'.jpg'))throw Error('Nome de imagem inválido.');
        pdfName(name.replace(/\.(png|jpg)$/,'')+'.pdf');
    }
    return {base:plan.base,format:plan.format,names:[...plan.names]};
}
export function validateImageBytes(bytes,format){
    if(!(bytes instanceof Uint8Array)||bytes.length<12||bytes.length>192*1024*1024)throw Error('Imagem inválida ou acima de 192 MiB.');
    const data=Buffer.from(bytes.buffer,bytes.byteOffset,bytes.byteLength);
    if(format==='png'&&!data.subarray(0,8).equals(Buffer.from([137,80,78,71,13,10,26,10])))throw Error('Cabeçalho PNG inválido.');
    if(format==='jpeg'&&(data[0]!==255||data[1]!==216||data.at(-2)!==255||data.at(-1)!==217))throw Error('Cabeçalho JPEG inválido.');
    return data;
}
// Cada lote recebe uma pasta nova exclusiva. Nunca substitui arquivos existentes.
export class ImageExportStore {
    constructor(window,dialog){this.window=window;this.dialog=dialog;this.job=null;this.queue=Promise.resolve();this.requests=0;}
    get active(){return !!this.job;}
    request(request,bytes){
        if(request?.action==='cancel'&&this.job?.token===request.token)this.job.abort.abort();
        if(this.requests>=2)return Promise.reject(Error('Aguarde a gravação atual.'));
        this.requests++;
        const result=this.queue.then(()=>this.execute(request,bytes));
        this.queue=result.catch(()=>{});return result.finally(()=>{this.requests--;});
    }
    async execute(request,bytes){
        if(request?.action==='begin'){
            if(this.job)throw Error('Uma exportação de imagens já está em andamento.');
            const plan=validateImagePlan(request.plan);
            const choice=await this.dialog.showOpenDialog(this.window,{title:'Exportar imagens — escolher pasta',properties:['openDirectory','createDirectory']});
            if(choice.canceled||!choice.filePaths?.length)return {canceled:true};
            const token=randomUUID();const directory=path.join(choice.filePaths[0],plan.base+'_'+plan.format.toUpperCase()+'_'+token.slice(0,8));
            await mkdir(directory,{recursive:false});
            this.job={...plan,token,directory,index:0,totalBytes:0,files:[],abort:new AbortController()};
            return {token,location:directory};
        }
        const job=this.job;
        if(!job||request?.token!==job.token)throw Error('Exportação vencida ou inválida.');
        if(request.action==='cancel'){await this.clean(job);return {canceled:true};}
        try{
            if(request.action==='write'){
                if(request.index!==job.index||job.index>=job.names.length)throw Error('Sequência de imagens inválida.');
                const data=validateImageBytes(bytes,job.format);
                if(job.totalBytes+data.length>2*1024*1024*1024)throw Error('O lote excede o limite de saída de 2 GiB.');
                const file=path.join(job.directory,job.names[job.index]+'.part');
                const handle=await open(file,'wx');job.files.push(file);
                try{await handle.writeFile(data,{signal:job.abort.signal});await handle.sync();}finally{await handle.close();}
                job.totalBytes+=data.length;job.index++;return {written:job.index};
            }
            if(request.action==='finish'){
                if(job.index!==job.names.length)throw Error('O lote ainda está incompleto.');
                for(let i=0;i<job.names.length;i++){
                    const target=path.join(job.directory,job.names[i]);
                    const handle=await open(target,'wx');job.files.push(target);
                    try{await handle.writeFile(createReadStream(job.files[i]),{signal:job.abort.signal});await handle.sync();}
                    finally{await handle.close();}
                    await unlink(job.files[i]);
                }
                this.job=null;return {saved:true,count:job.index,location:job.directory};
            }
            throw Error('Ação de imagens inválida.');
        }catch(error){await this.clean(job);throw error;}
    }
    async clean(job){
        // Apenas nomes registrados nesta pasta criada com mkdir exclusivo.
        try{for(const file of job.files)await unlink(file).catch(error=>{if(error.code!=='ENOENT')throw error;});await rmdir(job.directory).catch(error=>{if(!['ENOENT','ENOTEMPTY'].includes(error.code))throw error;});}
        finally{if(this.job===job)this.job=null;}
    }
    cancelAll(){
        if(!this.canceling){
            this.job?.abort.abort();
            this.canceling=this.queue.then(()=>this.job?this.clean(this.job):undefined).finally(()=>{this.canceling=null;});
        }
        return this.canceling;
    }
}
