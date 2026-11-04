import {calcularPreviaImpressao,PrintConfigPresentation} from './desktop-print-layout.js';

export class JanelaConfiguracaoImpressao {
    constructor(container,bridge,{enviarTeste,abrirSintetico,imprimir,fonte}={}){
        this.bridge=bridge;this.model=new PrintConfigPresentation();this.busy=false;
        this.enviarTeste=enviarTeste;this.abrirSintetico=abrirSintetico;this.imprimir=imprimir;this.fonte=fonte;this.attempted=false;
        this.dialog=document.createElement('dialog');this.dialog.id='print-config-dialog';this.dialog.className='print-config-dialog';this.dialog.setAttribute('aria-labelledby','print-config-title');
        this.dialog.innerHTML=`<h2 id="print-config-title">Impressão</h2>
            <p id="print-config-scope">Somente configuração e prévia técnica. Nenhuma folha será impressa.</p>
            <div class="print-config-fields"><label>Impressora<select id="print-config-printer"><option value="">Selecione uma impressora</option></select></label>
            <button type="button" id="print-config-refresh">Atualizar lista</button><button type="button" id="print-config-properties" disabled>Propriedades do driver</button></div>
            <dl id="print-config-summary"></dl>
            <div class="print-config-fields"><label>Papel da simulação<select id="print-config-paper"><option value="a3">A3 — 297 × 420 mm (simulação)</option><option value="driver">Papel informado pelo driver</option></select></label>
            <label>Escala do aplicativo<select id="print-config-scale"><option value="100">100%</option><option value="120">120%</option><option value="custom">Personalizada</option></select></label>
            <label id="print-config-custom-label" hidden>Percentual (1% a 400%)<input id="print-config-custom" type="text" inputmode="decimal" value="100" aria-label="Escala personalizada em porcentagem"></label></div>
            <svg id="print-config-preview" role="img" aria-label="Prévia técnica de uma página A4 centralizada"></svg>
            <p id="print-config-measures"></p><p id="print-config-warning"></p>
            <details><summary>Limites desta prévia</summary><p>O contorno azul representa a página atual do PDF. A área tracejada indica os limites imprimíveis informados pelo driver. A zona vermelha indica conteúdo fora do papel. A impressão inclui todas as páginas, na ordem do documento de trabalho.</p>
            <p>O aplicativo centraliza cada página e aplica a escala uma vez. Quando o zoom público do driver difere de 100%, ele prevalece e a escala do aplicativo fica suspensa. Ajustes privados, encaixe, imposição, cor, cópias e acabamentos são executados pelo driver e podem alterar o resultado da prévia técnica.</p></details>
            <button type="button" id="print-config-print" hidden>Imprimir</button>
            <details id="print-config-diagnostics"><summary>Diagnóstico opcional Konica / Hold</summary>
            <p id="print-config-hold" role="status"></p>
            <div class="print-config-fields"><button type="button" id="print-config-verify-hold" disabled>Verificar Hold da sessão</button>
            <button type="button" id="print-config-fixture" hidden>Abrir PDF sintético de teste</button>
            <button type="button" id="print-config-send-hold" disabled>Enviar trabalho retido</button></div>
            <button type="button" id="print-config-new-operation" hidden>Nova operação</button>
            <details><summary>Configurações recuperadas do PostScript</summary><pre id="print-config-ps-details">Aguardando captura verificada.</pre><p>São os valores codificados no PS do diagnóstico especializado. O botão Imprimir utiliza o contexto Windows da impressora selecionada.</p></details></details>
            <p id="print-config-status" role="status" aria-live="polite"></p><div class="print-config-actions"><button type="button" id="print-config-close">Fechar</button></div>`;
        container.append(this.dialog);this.get=id=>this.dialog.querySelector('#print-config-'+id);
        this.get('refresh').addEventListener('click',()=>this.list());this.get('printer').addEventListener('change',()=>this.select());
        this.get('properties').addEventListener('click',()=>this.properties());
        this.get('verify-hold').addEventListener('click',()=>this.verifyHold());
        this.get('fixture').addEventListener('click',()=>this.runTest(true));
        this.get('send-hold').addEventListener('click',()=>this.runTest(false));
        this.get('new-operation').addEventListener('click',()=>this.newOperation());
        this.get('print').addEventListener('click',()=>this.print());
        for(const id of ['paper','scale'])this.get(id).addEventListener('change',()=>{this.preview();this.controls();});
        this.get('custom').addEventListener('input',()=>{this.preview();this.controls();});
        this.get('close').addEventListener('click',()=>this.close());this.dialog.addEventListener('cancel',e=>{e.preventDefault();this.close();});
        this.preview();
    }
    async open(){if(this.closing)await this.closing;if(this.dialog.open)return;this.dialog.showModal();await this.list();}
    controls(){for(const id of ['refresh','printer','paper','scale','close'])this.get(id).disabled=this.busy;this.get('properties').disabled=this.busy||!this.model.state;this.get('verify-hold').disabled=this.busy||!this.model.state;
        if(this.model.state?.experimentalHold)this.get('paper').disabled=true;
        const normal=this.model.state?.windowsPrint;
        this.get('print').hidden=!normal;
        this.get('print').disabled=this.busy||!normal||!this.imprimir||!this.fonte?.()||!Number.isFinite(this.selectedScale());
        this.get('diagnostics').hidden=!!normal&&this.model.state.printer!=='KONICA MINOLTA C6100PS';
        this.get('custom').disabled=this.busy;
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
        this.get('scope').textContent=state?.windowsPrint?source?`Documento com ${source.pages} página(s). Prévia técnica da página ${source.page}, ${this.mm(source.width)} × ${this.mm(source.height)} mm. Imprimir envia o documento completo com as propriedades confirmadas.`:'Abra um PDF para imprimir. As propriedades permanecem disponíveis.':state?.experimentalHold?'Original PDF A4 — 210 × 297 mm. O driver define a folha. A escala do aplicativo controla a arte; o zoom do driver deve permanecer 100%.':'Somente configuração e prévia técnica. Nenhuma folha será impressa.';
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
        const source=state?.windowsPrint?this.fonte?.():{width:210,height:297};
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
        if(layout.printable){const p=layout.printable;rect(p.left,p.top,p.width,p.height,'print-printable');}
        this.get('measures').textContent=`Página a ${layout.scale}% no aplicativo: ${this.mm(content.width)} × ${this.mm(content.height)} mm. Margens geométricas: ${this.mm(layout.margins.left)} mm laterais; ${this.mm(layout.margins.top)} mm superior/inferior.`;
        const warnings=[];if(layout.outsidePaper)warnings.push('O conteúdo ultrapassa o papel.');if(layout.outsidePrintable)warnings.push('O conteúdo ultrapassa a área imprimível informada.');if(layout.geometryMismatch)warnings.push('As dimensões nominais e as do contexto diferem; área imprimível não sobreposta.');if(state?.driverScale&&state.driverScale!==100)warnings.push('O driver informa escala diferente de 100%; esta simulação não incorpora esse ajuste.');if(!layout.printable)warnings.push('Área imprimível não disponível para este papel.');
        if(state?.windowsPrint&&this.driverScaleWins())warnings.push('O zoom do driver prevalece. Escala do aplicativo suspensa; a prévia não incorpora a transformação do driver.');
        if(state?.windowsPrint)warnings.push('Possíveis cortes são avisos. O clique em Imprimir envia com os ajustes escolhidos.');
        this.get('warning').textContent=warnings.join(' ');
    }
    async close(){
        if(this.busy)return;this.busy=true;this.model.clear();this.dialog.close();this.render();
        this.closing=this.bridge.configuracaoImpressao({action:'release'}).catch(()=>{});
        try{await this.closing;}finally{this.closing=null;this.busy=false;this.controls();}
    }
}
