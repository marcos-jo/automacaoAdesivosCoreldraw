import {spawn} from 'node:child_process';
import {readFile,lstat} from 'node:fs/promises';
import {createHash} from 'node:crypto';
import path from 'node:path';

const actions=new Set(['list','select','properties','summary','release','verify-hold']);
export function validatePrintRequest(request){
    if(!request||typeof request!=='object'||Array.isArray(request)||!actions.has(request.action))throw new Error('Solicitação de configuração inválida.');
    const allowed={list:['action'],select:['action','printer'],properties:['action','session','revision'],summary:['action','session','revision'],release:['action'],'verify-hold':['action','session','revision']}[request.action];
    if(Object.keys(request).some(key=>!allowed.includes(key)))throw new Error('Parâmetro de configuração não permitido.');
    if(request.action==='select'&&(typeof request.printer!=='string'||!request.printer.trim()||request.printer.length>512||/[\x00-\x1f]/.test(request.printer)))throw new Error('Nome de impressora inválido.');
    if(['properties','summary','verify-hold'].includes(request.action)&&(!/^[a-f0-9]{32}$/.test(request.session)||!Number.isSafeInteger(request.revision)||request.revision<1))throw new Error('Sessão de configuração inválida.');
    return {...request};
}

export function validateHoldTestRequest(request){
    const keys=request?.action==='prepare-hold'?['action','session','revision','pdf']:['action','session','revision','token'];
    if(!request||!['prepare-hold','send-hold'].includes(request.action)||Object.keys(request).some(key=>!keys.includes(key)))throw new Error('Ensaio nativo inválido.');
    validatePrintRequest({action:'summary',session:request.session,revision:request.revision});
    if(request.action==='prepare-hold'&&(typeof request.pdf!=='string'||request.pdf.length>210000||!/^[A-Za-z0-9+/]+={0,2}$/.test(request.pdf)))throw new Error('PDF sintético inválido.');
    if(request.action==='send-hold'&&!/^[a-f0-9]{32}$/.test(request.token))throw new Error('Captura aprovada inválida.');
    return {...request};
}

// O renderer recebe somente resumos públicos; DEVMODE e handles ficam no auxiliar.
export class PrintConfigClient {
    constructor({helper,manifest,launch=spawn,platform=process.platform,timeout=60000,propertiesTimeout=600000,experimental=false}){
        Object.assign(this,{helper,manifest,launch,platform,timeout,propertiesTimeout,experimental});this.serial=0;this.generation=0;this.pending=null;this.child=null;this.closed=false;
    }
    get busy(){return !!this.pending;}
    async start(){
        if(this.child)return;
        if(this.platform!=='win32')throw new Error('Configuração nativa disponível somente no Windows.');
        const metadata=JSON.parse(await readFile(this.manifest,'utf8'));
        const stat=await lstat(this.helper);
        if(!stat.isFile()||stat.isSymbolicLink()||stat.size>4*1024*1024||metadata.schema!==1||metadata.file!=='4PrintShops.PrintConfig.exe'||! /^[a-f0-9]{64}$/.test(metadata.sha256)||createHash('sha256').update(await readFile(this.helper)).digest('hex')!==metadata.sha256)throw new Error('Auxiliar de configuração ausente ou incompatível. Reinstale a versão de teste.');
        if(this.experimental)for(const file of ['pdfium.dll','synthetic.pdf']){
            if(createHash('sha256').update(await readFile(path.join(path.dirname(this.helper),file))).digest('hex')!==metadata.files?.[file])throw new Error('Recurso nativo experimental incompatível.');
        }
        if(this.closed||!this.pending)return;
        const child=this.launch(this.helper,[String(process.pid)],{windowsHide:true,stdio:['pipe','pipe','ignore'],shell:false});
        this.child=child;const generation=++this.generation;let buffer='';
        child.stdout.setEncoding('utf8');
        child.stdout.on('data',chunk=>{
            if(generation!==this.generation)return;
            buffer+=chunk;
            if(buffer.length>262144){this.reset('Resposta inválida do auxiliar.');return;}
            let newline;
            while((newline=buffer.indexOf('\n'))>=0){const line=buffer.slice(0,newline).trim();buffer=buffer.slice(newline+1);if(!line)continue;
                let message;try{message=JSON.parse(line);}catch{this.reset('Resposta inválida do auxiliar.');return;}
                const pending=this.pending;
                if(!pending||message.id!==pending.id)continue;
                clearTimeout(pending.timer);this.pending=null;
                if(message.ok===true)pending.resolve(message.result);
                else pending.reject(new Error(typeof message.error==='string'&&message.error.length<300?message.error:'Falha no auxiliar de configuração.'));
            }
        });
        child.stdin.on('error',()=>{if(generation===this.generation)this.reset('A comunicação com o driver foi encerrada.');});
        child.once('error',()=>{if(generation===this.generation)this.reset('Não foi possível iniciar o auxiliar de configuração.');});
        child.once('exit',()=>{if(generation===this.generation)this.reset('O auxiliar de configuração foi encerrado. Selecione novamente a impressora.');});
    }
    async request(input,{owner='0'}={}){
        const request=this.experimental&&['prepare-hold','send-hold'].includes(input?.action)?validateHoldTestRequest(input):validatePrintRequest(input);
        if(this.closed)throw new Error('Configuração encerrada.');
        if(this.busy)throw new Error('Aguarde a configuração em andamento.');
        const id=++this.serial;
        // Reservar antes de await evita duas inicializações ou comandos concorrentes.
        const promise=new Promise((resolve,reject)=>{this.pending={id,resolve,reject};});
        void (async()=>{try{
            await this.start();
            if(this.pending?.id!==id||this.closed)return;
            if(request.action==='properties'){
                if(typeof owner!=='string'||!/^\d{1,20}$/.test(owner))throw new Error('Janela de configuração inválida.');
                request.owner=owner;
            }
            this.pending.timer=setTimeout(()=>this.reset('O driver não respondeu a tempo. A configuração foi cancelada.'),request.action==='properties'?this.propertiesTimeout:this.timeout);
            this.child.stdin.write(JSON.stringify({...request,id})+'\n');
        }catch(error){if(this.pending?.id===id){const pending=this.pending;this.pending=null;clearTimeout(pending.timer);pending.reject(error);}}})();
        return promise;
    }
    reset(message='Configuração cancelada.'){
        ++this.generation;const child=this.child;this.child=null;
        const pending=this.pending;this.pending=null;if(pending){clearTimeout(pending.timer);pending.reject(new Error(message));}
        if(child){child.stdin.destroy();child.kill();}
    }
    destroy(){this.closed=true;this.reset();}
}
