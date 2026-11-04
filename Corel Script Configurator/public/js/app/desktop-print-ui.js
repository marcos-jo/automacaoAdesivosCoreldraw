import {calcularPreviaImpressao,PrintConfigPresentation} from './desktop-print-layout.js';

export class JanelaConfiguracaoImpressao {
    constructor(container,bridge){
        this.bridge=bridge;this.model=new PrintConfigPresentation();this.busy=false;
        this.dialog=document.createElement('dialog');this.dialog.id='print-config-dialog';this.dialog.className='print-config-dialog';this.dialog.setAttribute('aria-labelledby','print-config-title');
        this.dialog.innerHTML=`<h2 id="print-config-title">Configuração de impressão — protótipo</h2>
            <p>Somente configuração e prévia técnica. Nenhuma folha será impressa.</p>
            <div class="print-config-fields"><label>Impressora<select id="print-config-printer"><option value="">Selecione uma impressora</option></select></label>
            <button type="button" id="print-config-refresh">Atualizar lista</button><button type="button" id="print-config-properties" disabled>Propriedades do driver</button></div>
            <dl id="print-config-summary"></dl>
            <div class="print-config-fields"><label>Papel da simulação<select id="print-config-paper"><option value="a3">A3 — 297 × 420 mm (simulação)</option><option value="driver">Papel informado pelo driver</option></select></label>
            <label>Escala do A4<select id="print-config-scale"><option value="100">100%</option><option value="120">120%</option></select></label></div>
            <svg id="print-config-preview" role="img" aria-label="Prévia técnica de uma página A4 centralizada"></svg>
            <p id="print-config-measures"></p><p id="print-config-warning"></p>
            <details><summary>Limites desta prévia</summary><p>O contorno azul representa uma página A4 sintética. A área tracejada indica os limites imprimíveis quando informados pelo driver. A zona vermelha indica conteúdo fora do papel. As margens em branco permanecem visíveis.</p>
            <p>A escala desta simulação não altera o driver. Zoom, ajuste ao papel, imposição e acabamentos privados continuam controlados pelas propriedades nativas e ainda exigem homologação. A prévia não comprova a saída física.</p></details>
            <p id="print-config-status" role="status" aria-live="polite"></p><div class="print-config-actions"><button type="button" id="print-config-close">Fechar</button></div>`;
        container.append(this.dialog);this.get=id=>this.dialog.querySelector('#print-config-'+id);
        this.get('refresh').addEventListener('click',()=>this.list());this.get('printer').addEventListener('change',()=>this.select());
        this.get('properties').addEventListener('click',()=>this.properties());
        for(const id of ['paper','scale'])this.get(id).addEventListener('change',()=>this.preview());
        this.get('close').addEventListener('click',()=>this.close());this.dialog.addEventListener('cancel',e=>{e.preventDefault();this.close();});
        this.preview();
    }
    async open(){if(this.closing)await this.closing;if(this.dialog.open)return;this.dialog.showModal();await this.list();}
    controls(){for(const id of ['refresh','printer','paper','scale','close'])this.get(id).disabled=this.busy;this.get('properties').disabled=this.busy||!this.model.state;}
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
        this.get('status').textContent=printers.length+' impressora(s) instalada(s). Selecione uma fila para consultar seus ajustes.';
    });}
    async select(){const name=this.get('printer').value;if(!name){this.model.clear();this.render();return;}
        await this.run({action:'select',printer:name},(state,ticket)=>{if(this.model.accept(ticket,state)){this.render();this.get('status').textContent='Configuração consultada. Nenhuma preferência foi gravada.';}});
    }
    async properties(){const state=this.model.state;if(!state)return;
        await this.run({action:'properties',session:state.session,revision:state.revision},(next,ticket)=>{if(this.model.accept(ticket,next)){this.render();this.get('status').textContent=next.canceled?'Propriedades canceladas. Configuração anterior preservada.':'Configuração confirmada somente nesta sessão.';}});
    }
    render(){
        const state=this.model.state,summary=this.get('summary');summary.replaceChildren();
        if(state){const unknown='Controlado pelo driver / não informado',paper=state.paper;
            for(const [label,value]of [['Papel',paper?`${paper.name||'Formato'} — ${this.mm(paper.width)} × ${this.mm(paper.height)} mm`:unknown],['Orientação',({1:'Retrato',2:'Paisagem'})[state.orientation]||unknown],['Cópias',state.copies>0?state.copies:unknown],['Bandeja',state.tray||unknown],['Duplex',({1:'Uma face',2:'Borda longa',3:'Borda curta'})[state.duplex]||unknown],['Escala pública do driver',state.driverScale>0?state.driverScale+'%':unknown],['Área imprimível',state.geometry?`${this.mm(state.geometry.printableWidth)} × ${this.mm(state.geometry.printableHeight)} mm`:unknown]]){
                const term=document.createElement('dt'),valueNode=document.createElement('dd');term.textContent=label;valueNode.textContent=String(value);summary.append(term,valueNode);
            }
        }
        this.preview();this.controls();
    }
    mm(value){return Number(value).toLocaleString('pt-BR',{maximumFractionDigits:2});}
    preview(){
        const state=this.model.state,paper=this.get('paper').value==='driver'?(state?.paper||state?.geometry):{width:297,height:420};
        const svg=this.get('preview');svg.replaceChildren();
        if(!paper){this.get('measures').textContent='Selecione uma impressora que informe as dimensões do papel.';this.get('warning').textContent='';return;}
        const layout=calcularPreviaImpressao(paper,Number(this.get('scale').value),state?.geometry),content=layout.content;
        const padding=Math.max(12,-content.left+8,-content.top+8);
        svg.setAttribute('viewBox',`${-padding} ${-padding} ${paper.width+padding*2} ${paper.height+padding*2}`);
        const rect=(x,y,width,height,kind)=>{const element=document.createElementNS('http://www.w3.org/2000/svg','rect');for(const [key,value]of Object.entries({x,y,width,height,class:kind}))element.setAttribute(key,String(value));svg.append(element);return element;};
        rect(0,0,paper.width,paper.height,'print-sheet');rect(content.left,content.top,content.width,content.height,'print-overflow');
        const defs=document.createElementNS(svg.namespaceURI,'defs'),clip=document.createElementNS(svg.namespaceURI,'clipPath');clip.id='print-config-paper-clip';const clipRect=document.createElementNS(svg.namespaceURI,'rect');clipRect.setAttribute('width',String(paper.width));clipRect.setAttribute('height',String(paper.height));clip.append(clipRect);defs.append(clip);svg.append(defs);
        rect(content.left,content.top,content.width,content.height,'print-content').setAttribute('clip-path','url(#print-config-paper-clip)');
        if(layout.printable){const p=layout.printable;rect(p.left,p.top,p.width,p.height,'print-printable');}
        this.get('measures').textContent=`A4 a ${layout.scale}%: ${this.mm(content.width)} × ${this.mm(content.height)} mm. Margens geométricas: ${this.mm(layout.margins.left)} mm laterais; ${this.mm(layout.margins.top)} mm superior/inferior.`;
        const warnings=[];if(layout.outsidePaper)warnings.push('O conteúdo ultrapassa o papel.');if(layout.outsidePrintable)warnings.push('O conteúdo ultrapassa a área imprimível informada.');if(layout.geometryMismatch)warnings.push('As dimensões nominais e as do contexto diferem; área imprimível não sobreposta.');if(state?.driverScale&&state.driverScale!==100)warnings.push('O driver informa escala diferente de 100%; esta simulação não incorpora esse ajuste.');if(!layout.printable)warnings.push('Área imprimível não disponível para este papel.');
        this.get('warning').textContent=warnings.join(' ');
    }
    async close(){
        if(this.busy)return;this.busy=true;this.model.clear();this.dialog.close();this.render();
        this.closing=this.bridge.configuracaoImpressao({action:'release'}).catch(()=>{});
        try{await this.closing;}finally{this.closing=null;this.busy=false;this.controls();}
    }
}
