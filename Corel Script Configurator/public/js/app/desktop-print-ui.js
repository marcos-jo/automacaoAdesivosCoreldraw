import {calcularPreviaImpressao,PrintConfigPresentation} from './desktop-print-layout.js';
import {interfaceImpressao} from './desktop-print-interface.js';
import {PreviaDocumentoImpressao} from './desktop-print-preview.js';
import {FORMATOS_PADRAO} from '../tools/pdf-formats.js';

export class JanelaConfiguracaoImpressao {
    constructor(container,bridge,{enviarTeste,abrirSintetico,imprimir,fonte,prepararPrevia}={}){
        this.bridge=bridge;this.model=new PrintConfigPresentation();this.busy=false;
        this.enviarTeste=enviarTeste;this.abrirSintetico=abrirSintetico;this.imprimir=imprimir;this.fonte=fonte;this.attempted=false;
        this.prepararPrevia=prepararPrevia;
        this.dialog=document.createElement('dialog');this.dialog.id='print-config-dialog';this.dialog.className='print-config-dialog';this.dialog.setAttribute('aria-labelledby','print-config-title');
        this.dialog.innerHTML=interfaceImpressao;
        container.append(this.dialog);this.get=id=>this.dialog.querySelector('#print-config-'+id);
        this.get('refresh').addEventListener('click',()=>this.list());this.get('printer').addEventListener('change',()=>this.select());
        this.get('properties').addEventListener('click',()=>this.properties());
        this.get('verify-hold').addEventListener('click',()=>this.verifyHold());
        this.get('fixture').addEventListener('click',()=>this.runTest(true));
        this.get('send-hold').addEventListener('click',()=>this.runTest(false));
        this.get('new-operation').addEventListener('click',()=>this.newOperation());
        this.get('print').addEventListener('click',()=>this.print());
        this.get('previous').addEventListener('click',()=>this.page(-1));this.get('next').addEventListener('click',()=>this.page(1));
        for(const id of ['paper','scale'])this.get(id).addEventListener('change',()=>{this.preview();this.controls();});
        this.get('custom').addEventListener('input',()=>{this.preview();this.controls();});
        this.get('close').addEventListener('click',()=>this.close());this.dialog.addEventListener('cancel',e=>{e.preventDefault();this.close();});
        this.preview();
    }
    async open(){
        if(this.closing)await this.closing;if(this.dialog.open)return;
        this.dialog.showModal();await this.list();
        if(!this.prepararPrevia||!this.fonte?.())return;
        this.busy=true;this.controls();
        this.documentPreview=new PreviaDocumentoImpressao(this.get('canvas'));
        try{const snapshot=await this.prepararPrevia();if(snapshot)await this.documentPreview.open(snapshot);this.render();}
        catch(error){this.get('status').textContent='Prévia indisponível. '+error.message;}
        finally{this.busy=false;this.controls();}
    }
    async page(step){
        if(this.busy||!this.documentPreview?.source)return;
        this.busy=true;this.controls();
        try{await this.documentPreview.show(this.documentPreview.page+step);this.preview();}
        catch(error){this.get('status').textContent='Não foi possível atualizar a prévia. '+error.message;}
        finally{this.busy=false;this.controls();}
    }
    controls(){for(const id of ['refresh','printer','paper','scale','close'])this.get(id).disabled=this.busy;this.get('properties').disabled=this.busy||!this.model.state;this.get('verify-hold').disabled=this.busy||!this.model.state;
        if(this.model.state?.experimentalHold)this.get('paper').disabled=true;
        const normal=this.model.state?.windowsPrint;
        this.get('print').hidden=!normal;
        this.get('print').disabled=this.busy||!normal||!this.imprimir||!this.fonte?.()||!Number.isFinite(this.selectedScale());
        this.get('diagnostics').hidden=!!normal&&this.model.state.printer!=='KONICA MINOLTA C6100PS';
        this.get('custom').disabled=this.busy;
        const page=this.documentPreview?.source;
        this.get('previous').disabled=this.busy||!page||page.page<=1;
        this.get('next').disabled=this.busy||!page||page.page>=page.pages;
        if(normal&&this.driverScaleWins()){this.get('scale').disabled=true;this.get('custom').disabled=true;}
        this.get('new-operation').hidden=!this.model.state?.experimentalHold||!this.attempted;this.get('new-operation').disabled=this.busy;
        this.get('fixture').hidden=!this.model.state?.experimentalHold;this.get('fixture').disabled=this.busy;
        this.get('send-hold').disabled=this.busy||this.attempted||!this.model.state?.experimentalHold||!this.model.state?.hold?.holdConfirmed||!this.enviarTeste||!Number.isFinite(this.selectedScale());
    }
    selectedScale(){
        if(this.get('scale').value!=='custom')return Number(this.get('scale').value);
        const text=this.get('custom').value.trim();
        const value=/^\d{1,3}([.,]\d{1,2})?$/.test(text)?Number(text.replace(',','.')):NaN;
        return value>=1&&value<=400?value:NaN;
    }
    driverScaleWins(){const scale=this.model.state?.driverScale;return typeof scale==='number'&&scale>0&&scale!==100;}
    async print(){
        if(this.busy||!this.model.state?.windowsPrint||!this.fonte?.())return;
        this.busy=true;this.controls();this.get('status').textContent='Preparando o documento de trabalho e enviando ao driver…';
        try {
            const result=await this.imprimir(this.model.state,this.driverScaleWins()?100:this.selectedScale());
            if(!result)throw new Error('Documento de trabalho indisponível.');
            this.get('status').textContent=result.canceled?'Impressão cancelada. Documento preservado.':result.uncertain?result.message:`Trabalho ${result.jobId} enviado para ${result.printer}. ${result.pages} página(s). Confira a fila e o resultado na impressora.`;
        }catch(error){this.get('status').textContent=error.message;}
        finally{this.busy=false;this.controls();}
    }
    async newOperation(){
        if(this.busy||!this.attempted)return;
        this.busy=true;this.controls();
        const ticket=this.model.begin();
        try {
            const state=await this.bridge.ensaioRetido({kind:'new-operation'});
            if(!state.canceled&&this.model.accept(ticket,state)){
                this.attempted=false;this.get('ps-details').textContent='Nova operação. Aguardando nova captura.';this.render();
                this.get('status').textContent='Nova sessão criada. Configure as propriedades e verifique Hold novamente.';
            }
        }catch(error){this.get('status').textContent=error.message;}
        finally{this.busy=false;this.controls();}
    }
    async runTest(fixture){
        if(this.busy||!this.model.state?.experimentalHold)return;
        this.busy=true;this.controls();
        this.get('status').textContent=fixture?'Abrindo sintético…':'Capturando e verificando o PS antes da confirmação…';
        try {
            if(fixture){await this.abrirSintetico();this.get('status').textContent='Sintético solicitado. O workspace e suas proteções continuam ativos.';}
            else {
                const result=await this.enviarTeste(this.model.state,this.selectedScale());
                this.attempted=!!(result?.submitted||result?.uncertain||result?.canceled);
                if(result?.diagnostic)this.get('ps-details').textContent=result.diagnostic;
                this.get('status').textContent=result?.canceled?'Envio cancelado. Nenhum trabalho físico iniciado.':result?.message||'Confira a retenção na Konica.';
            }
        } catch(error){this.get('status').textContent=error.message;}
        finally{this.busy=false;this.controls();}
    }
    async run(action,apply){
        if(this.busy)return;this.busy=true;this.controls();this.get('status').textContent=action.action==='properties'?'Aguardando as propriedades nativas…':'Consultando configuração…';
        const ticket=this.model.begin();
        try{const result=await this.bridge.configuracaoImpressao(action);if(this.dialog.open&&ticket===this.model.generation){apply(result,ticket);}}
        catch(error){if(ticket===this.model.generation){this.model.clear();this.get('status').textContent=error.message;this.render();}}
        finally{this.busy=false;this.controls();}
    }
    async list(){await this.run({action:'list'},printers=>{
        const previous=this.model.state?.printer||'';this.get('printer').replaceChildren(new Option('Selecione uma impressora',''));
        for(const printer of printers)this.get('printer').add(new Option(printer.name,printer.name));
        this.get('printer').value=previous;if(this.get('printer').value!==previous){this.model.clear();this.render();}
        this.render();this.get('status').textContent=printers.length+' impressora(s) instalada(s). Selecione uma fila para consultar seus ajustes.';
    });}
    async select(){const name=this.get('printer').value;if(!name){this.model.clear();this.render();return;}
        await this.run({action:'select',printer:name},(state,ticket)=>{if(this.model.accept(ticket,state)){this.render();this.get('status').textContent='Configuração consultada. Nenhuma preferência foi gravada.';}});
    }
    async properties(){const state=this.model.state;if(!state)return;
        await this.run({action:'properties',session:state.session,revision:state.revision},(next,ticket)=>{if(this.model.accept(ticket,next)){this.render();this.get('status').textContent=next.canceled?'Propriedades canceladas. Configuração anterior preservada.':'Configuração confirmada somente nesta sessão.';}});
    }
    async verifyHold(){
        const state=this.model.state;
        if(!state)return;
        await this.run({action:'verify-hold',session:state.session,revision:state.revision},(next,ticket)=>{
            if(this.model.accept(ticket,next)){
                this.render();
                this.get('status').textContent=next.hold?.holdConfirmed?'Hold confirmado nesta configuração. O diagnóstico pode capturar e verificar a saída.':'Hold não comprovado para o diagnóstico especializado. A impressão normal usa as propriedades do driver.';
            }
        });
    }
    render(){
        const state=this.model.state,summary=this.get('summary');summary.replaceChildren();
        this.get('send-hold').textContent=state?.experimentalHold?'Enviar trabalho retido — Experimental':'Enviar trabalho retido';
        const source=this.fonte?.();
        this.get('scope').textContent=source?`${source.name||'Documento de trabalho'} · ${source.pages} página(s)`:'Selecione a impressora e configure suas propriedades.';
        this.get('copies').textContent=state?.copies>0?state.copies:'Controlado pelo driver';
        this.get('orientation').textContent=({1:'Retrato',2:'Paisagem'})[state?.orientation]||'Controlada pelo driver';
        this.get('pages').textContent=source?`Todas as ${source.pages} páginas`:'Todas as páginas do documento';
        this.get('paper-label').hidden=!!state?.windowsPrint;
        if(state?.experimentalHold)this.get('paper').value='driver';
        if(state){const unknown='Controlado pelo driver / não informado',paper=state.paper;
            if(state.queue)for(const [label,value]of [['Driver',state.queue.driver],['Linguagem',state.queue.language],['Caminho de impressão',state.windowsPrint?'PDFium → contexto Windows → driver selecionado':'Configuração nativa']]){
                const term=document.createElement('dt'),node=document.createElement('dd');term.textContent=label;node.textContent=value||unknown;summary.append(term,node);
            }
            for(const [label,value]of [['Papel lógico do driver',paper?`${paper.name||'Formato'} — ${this.mm(paper.width)} × ${this.mm(paper.height)} mm`:unknown],['Folha física informada',state.geometry?`${this.mm(state.geometry.width)} × ${this.mm(state.geometry.height)} mm`:unknown],['Orientação',({1:'Retrato',2:'Paisagem'})[state.orientation]||unknown],['Cópias',state.copies>0?state.copies:unknown],['Bandeja',state.tray||unknown],['Duplex',({1:'Uma face',2:'Borda longa',3:'Borda curta'})[state.duplex]||unknown],['Zoom público do driver',state.driverScale>0?state.driverScale+'%':unknown],['Área imprimível',state.geometry?`${this.mm(state.geometry.printableWidth)} × ${this.mm(state.geometry.printableHeight)} mm`:unknown]]){
                const term=document.createElement('dt'),valueNode=document.createElement('dd');term.textContent=label;valueNode.textContent=String(value);summary.append(term,valueNode);
            }
        }
        this.get('hold').textContent=state?.hold?.reason||'O envio à Konica aguarda comprovação da retenção na saída gerada. Use o leitor externo para o fluxo de impressão atual.';
        if(state?.experimentalHold&&state.hold?.holdConfirmed)this.get('hold').textContent='Hold confirmado nesta sessão. A saída será capturada e interpretada antes da confirmação final de envio.';
        this.preview();this.controls();
    }
    mm(value){return Number(value).toLocaleString('pt-BR',{maximumFractionDigits:2});}
    preview(){
        const state=this.model.state,paper=this.get('paper').value==='driver'?(state?.geometry||state?.paper):{width:297,height:420};
        const svg=this.get('preview');svg.replaceChildren();
        if(!paper){this.get('measures').textContent='Selecione uma impressora que informe as dimensões do papel.';this.get('warning').textContent='';return;}
        this.get('custom-label').hidden=this.get('scale').value!=='custom';
        const scale=this.selectedScale();
        if(!Number.isFinite(scale)||scale<1||scale>400){this.get('warning').textContent='Informe uma escala entre 1% e 400%, com até duas casas decimais.';this.get('measures').textContent='';return;}
        const source=this.documentPreview?.source||(state?.windowsPrint?this.fonte?.():{width:210,height:297});
        if(!source){this.get('measures').textContent='Abra um PDF para visualizar suas dimensões.';this.get('warning').textContent='';return;}
        const appliedScale=state?.windowsPrint&&this.driverScaleWins()?100:scale;
        const layout=calcularPreviaImpressao(paper,appliedScale,state?.geometry,source),content=layout.content;
        svg.setAttribute('aria-label','Prévia técnica da página atual centralizada na folha');
        const padding=Math.max(12,-content.left+8,-content.top+8);
        svg.setAttribute('viewBox',`${-padding} ${-padding} ${paper.width+padding*2} ${paper.height+padding*2}`);
        const rect=(x,y,width,height,kind)=>{const element=document.createElementNS('http://www.w3.org/2000/svg','rect');for(const [key,value]of Object.entries({x,y,width,height,class:kind}))element.setAttribute(key,String(value));svg.append(element);return element;};
        rect(0,0,paper.width,paper.height,'print-sheet');rect(content.left,content.top,content.width,content.height,'print-overflow');
        const defs=document.createElementNS(svg.namespaceURI,'defs'),clip=document.createElementNS(svg.namespaceURI,'clipPath');clip.id='print-config-paper-clip';const clipRect=document.createElementNS(svg.namespaceURI,'rect');clipRect.setAttribute('width',String(paper.width));clipRect.setAttribute('height',String(paper.height));clip.append(clipRect);defs.append(clip);svg.append(defs);
        rect(content.left,content.top,content.width,content.height,'print-content').setAttribute('clip-path','url(#print-config-paper-clip)');
        if(this.documentPreview?.image){
            const image=document.createElementNS(svg.namespaceURI,'image');
            for(const [key,value]of Object.entries({href:this.documentPreview.image,x:content.left,y:content.top,width:content.width,height:content.height,'preserveAspectRatio':'none','clip-path':'url(#print-config-paper-clip)'}))image.setAttribute(key,String(value));
            svg.append(image);
        }
        if(layout.printable){const p=layout.printable;rect(p.left,p.top,p.width,p.height,'print-printable');}
        this.get('page').textContent=`${source.page||1} de ${source.pages||1}`;
        const format=dimensions=>FORMATOS_PADRAO.find(f=>Math.abs(f.largura-Math.min(dimensions.width,dimensions.height))<.5&&Math.abs(f.altura-Math.max(dimensions.width,dimensions.height))<.5)?.nome||'';
        this.get('measures').textContent=`PDF ${format(source)} ${this.mm(source.width)} × ${this.mm(source.height)} mm · Papel ${format(paper)} ${this.mm(paper.width)} × ${this.mm(paper.height)} mm · Escala ${this.mm(layout.scale)}%`;
        const warnings=[];if(layout.outsidePaper)warnings.push('O conteúdo ultrapassa o papel.');if(layout.outsidePrintable)warnings.push('O conteúdo ultrapassa a área imprimível informada.');if(layout.geometryMismatch)warnings.push('As dimensões nominais e as do contexto diferem; área imprimível não sobreposta.');if(state?.driverScale&&state.driverScale!==100)warnings.push('O driver informa escala diferente de 100%; esta simulação não incorpora esse ajuste.');if(!layout.printable)warnings.push('Área imprimível não disponível para este papel.');
        if(state?.windowsPrint&&this.driverScaleWins())warnings.push('O zoom do driver prevalece. Escala do aplicativo suspensa; a prévia não incorpora a transformação do driver.');
        this.get('warning').textContent=warnings.join(' ');
        this.get('warning').hidden=!warnings.length;
    }
    async close(){
        if(this.busy)return;this.busy=true;this.model.clear();this.dialog.close();this.render();
        const preview=this.documentPreview;this.documentPreview=null;
        this.closing=Promise.all([this.bridge.configuracaoImpressao({action:'release'}).catch(()=>{}),preview?.close()]);
        try{await this.closing;}finally{this.closing=null;this.busy=false;this.controls();}
    }
}
