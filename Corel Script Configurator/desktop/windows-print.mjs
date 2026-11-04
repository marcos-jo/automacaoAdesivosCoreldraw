import {mkdir,writeFile,unlink} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import {randomUUID} from 'node:crypto';
import path from 'node:path';
import {pdfData} from './security.mjs';
import {validatePrintRequest} from './print-config.mjs';

export function validateDocumentPrint(request){
    const allowed=['session','revision','scale','name','operation'];
    if(!request||typeof request!=='object'||Object.keys(request).some(key=>!allowed.includes(key)))throw new Error('Solicitação de impressão inválida.');
    validatePrintRequest({action:'summary',session:request.session,revision:request.revision});
    if(!Number.isFinite(request.scale)||request.scale<1||request.scale>400)throw new Error('Use escala entre 1% e 400%.');
    if(typeof request.name!=='string'||!request.name.trim()||request.name.length>180||/[\x00-\x1f]/.test(request.name))throw new Error('Nome de trabalho inválido.');
    if(!/^[a-f0-9]{32}$/.test(request.operation))throw new Error('Identificação de operação inválida.');
    return {...request};
}

// Nenhum retry. O clique cria um snapshot e um token que só pode ser consumido uma vez.
export class WindowsPrintCoordinator {
    constructor(client,chooseOutput){this.client=client;this.chooseOutput=chooseOutput;this.busy=false;this.operations=new Set();}
    async run(input,bytes){
        const request=validateDocumentPrint(input);pdfData(bytes);
        if(this.busy)throw new Error('Aguarde o trabalho atual.');
        if(this.operations.has(request.operation))throw new Error('Esta operação já foi utilizada. Um novo clique deve criar outro trabalho.');
        this.busy=true;this.operations.add(request.operation);
        const folder=path.join(tmpdir(),'4PrintShops-print-jobs'),file=path.join(folder,randomUUID().replaceAll('-','')+'.pdf');
        let sending=false;
        const session={session:request.session,revision:request.revision};
        try {
            await mkdir(folder,{recursive:true});await writeFile(file,bytes,{flag:'wx',mode:0o600});
            const state=await this.client.request({action:'summary',...session});
            const proof=await this.client.request({action:'prepare-document',...session,file,scale:request.scale});
            if(proof.session!==request.session||proof.revision!==request.revision||proof.printer!==state.printer||!Number.isSafeInteger(proof.pages)||proof.pages<1||!/^[a-f0-9]{32}$/.test(proof.token))throw new Error('O documento e a sessão de impressão não foram comprovados.');
            let output;
            if(state.printer==='Microsoft Print to PDF'&&state.queue?.driver==='Microsoft Print To PDF'){
                output=await this.chooseOutput(request.name);
                if(!output)return {canceled:true};
            }
            sending=true;
            return await this.client.request({action:'print-document',...session,token:proof.token,name:request.name,...(output?{output}:{})});
        }catch(error){
            if(sending)return {uncertain:true,message:'Resultado de impressão incerto. Confira a fila e a impressora antes de iniciar outro trabalho. '+error.message};
            throw error;
        }finally{
            if(!this.client.busy)await this.client.request({action:'cancel-document',...session}).catch(()=>{});
            await unlink(file).catch(()=>{});this.busy=false;
        }
    }
}
