import {pdfData} from './security.mjs';

// A confirmação permanece no processo principal. O renderer nunca recebe o token de envio.
export class HoldTestCoordinator {
    constructor(client,dialog,getWindow){Object.assign(this,{client,dialog,getWindow});this.busy=false;this.attempted=false;}
    async run(state,bytes){
        if(this.busy||this.attempted)throw new Error('Ensaio já iniciado. Confira a Konica antes de outro trabalho.');
        pdfData(bytes);
        if(bytes.length>150000)throw new Error('Use somente o PDF sintético deste laboratório.');
        this.busy=true;
        try {
            const proof=await this.client.request({action:'prepare-hold',session:state.session,revision:state.revision,pdf:Buffer.from(bytes).toString('base64')});
            if(proof.printer!=='KONICA MINOLTA C6100PS'||proof.revision!==state.revision||proof.paper!=='A3'||proof.scale!==120||proof.copies!==1||proof.destination!=='Hold'||proof.outputValidated!==true||!/^[a-f0-9]{32}$/.test(proof.token)||!/^4PS HOLD EXP [a-f0-9]{8}$/.test(proof.jobName))throw new Error('A saída não foi comprovada. Envio bloqueado.');
            const window=this.getWindow();
            if(!window||window.isDestroyed())throw new Error('Janela do ensaio encerrada.');
            const answer=await this.dialog.showMessageBox(window,{
                type:'warning',title:'Enviar trabalho retido — Experimental',
                message:'Autorizar um único trabalho sintético retido na Konica?',
                detail:'Trabalho '+proof.jobName+'\nImpressora KONICA MINOLTA C6100PS\nPapel A3 retrato\nConteúdo A4 a 120%\nZoom do driver 100%\nCópias 1\nDestino Hold\n\nA saída PostScript foi interpretada e verificada. Confira pessoalmente o armazenamento na máquina antes de liberar qualquer impressão física. O documento continua pendente de exportação quando houver edições.',
                buttons:['Cancelar','Enviar trabalho retido'],defaultId:0,cancelId:0,noLink:true
            });
            if(answer.response!==1)return {canceled:true,outputValidated:true};
            this.attempted=true;
            try{return await this.client.request({action:'send-hold',session:state.session,revision:state.revision,token:proof.token});}
            catch{return {uncertain:true,message:'Resultado de envio incerto. Confira a fila e a Konica antes de outro ensaio. Reenvio automático bloqueado.'};}
        } finally {this.busy=false;}
    }
}
