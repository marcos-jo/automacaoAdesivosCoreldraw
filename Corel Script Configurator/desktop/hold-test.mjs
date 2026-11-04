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
            const scale=state.appScale===undefined?100:state.appScale;
            if(!Number.isFinite(scale)||scale<1||scale>400)throw new Error('Use escala entre 1% e 400%.');
            const proof=await this.client.request({action:'prepare-hold',session:state.session,revision:state.revision,scale,pdf:Buffer.from(bytes).toString('base64')});
            if(proof.printer!=='KONICA MINOLTA C6100PS'||proof.revision!==state.revision||!['A4','A3'].includes(proof.paper)||proof.scale!==scale||proof.driverZoom!==100||proof.copies!==1||proof.destination!=='Hold'||proof.outputValidated!==true||typeof proof.cutRisk!=='boolean'||!/^[a-f0-9]{32}$/.test(proof.token)||!/^4PS HOLD EXP [a-f0-9]{8}$/.test(proof.jobName))throw new Error('A saída não foi comprovada. Envio bloqueado.');
            const window=this.getWindow();
            if(!window||window.isDestroyed())throw new Error('Janela do ensaio encerrada.');
            this.attempted=true;this.lastSession=state.session;this.lastJob=proof.jobName;
            const dimensions='PDF original A4 — 210 × 297 mm\nPapel do driver '+proof.paper+'\nEscala do aplicativo '+proof.scale+'%\nConteúdo '+Number(proof.contentWidthMM).toFixed(2)+' × '+Number(proof.contentHeightMM).toFixed(2)+' mm\nZoom do driver 100%\nCópias 1\nDestino Hold';
            if(proof.cutRisk){
                const crop=await this.dialog.showMessageBox(window,{
                    type:'warning',title:'Possível corte na impressão',message:'A arte ultrapassa o papel ou a área imprimível informada pelo driver.',
                    detail:dimensions+'\n\nAs partes externas poderão ser cortadas. A verificação Hold permanece obrigatória.',
                    buttons:['Cancelar','Continuar com possível corte'],defaultId:0,cancelId:0,noLink:true
                });
                if(crop.response!==1)return {canceled:true,diagnostic:proof.psSettings};
            }
            const answer=await this.dialog.showMessageBox(window,{
                type:'warning',title:'Enviar trabalho retido — Experimental',
                message:'Autorizar um único trabalho sintético retido na Konica?',
                detail:'Trabalho '+proof.jobName+'\nImpressora KONICA MINOLTA C6100PS\n'+dimensions+'\n\nCampos recuperados do PS\n'+(proof.psSettings||'Não informados')+'\n\nO trabalho permanece sujeito à conferência na IC-604. A edição do PDF continua pendente de exportação quando houver alterações.',
                buttons:['Cancelar','Enviar trabalho retido'],defaultId:0,cancelId:0,noLink:true
            });
            if(answer.response!==1)return {canceled:true,outputValidated:true,diagnostic:proof.psSettings};
            try{const result=await this.client.request({action:'send-hold',session:state.session,revision:state.revision,token:proof.token});return {...result,diagnostic:proof.psSettings};}
            catch{return {uncertain:true,diagnostic:proof.psSettings,message:'Resultado de envio incerto. Confira a fila e a Konica antes de criar deliberadamente outra operação. Reenvio automático bloqueado.'};}
        } finally {this.busy=false;}
    }
    async newOperation(){
        if(this.busy||!this.attempted)throw new Error('Conclua ou cancele a operação atual.');
        this.busy=true;
        try {
            const window=this.getWindow();
            if(!window||window.isDestroyed())throw new Error('Janela encerrada.');
            const answer=await this.dialog.showMessageBox(window,{
                type:'question',title:'Nova operação de impressão',message:'Criar uma nova operação independente?',
                detail:'Confira o resultado de '+this.lastJob+' na Konica, especialmente se houve erro ou timeout. A nova operação terá outra sessão e exigirá propriedades, verificação Hold, captura e confirmação individuais.',
                buttons:['Cancelar','Conferi o resultado, criar nova operação'],defaultId:0,cancelId:0,noLink:true
            });
            if(answer.response!==1)return {canceled:true};
            await this.client.request({action:'release'});
            const state=await this.client.request({action:'select',printer:'KONICA MINOLTA C6100PS'});
            if(!/^[a-f0-9]{32}$/.test(state.session)||state.session===this.lastSession)throw new Error('A nova sessão não foi comprovada.');
            this.attempted=false;
            return state;
        } finally {this.busy=false;}
    }
}
