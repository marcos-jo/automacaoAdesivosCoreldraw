import {renderizarFolhaMontagem,aguardarMontagem} from './pdf-imposition-preview.js';
import {desenharInspecaoMontagem,desenharFolhaVazia} from './pdf-imposition-view.js';
import {salvarArquivoPDF} from '../app/save-file.js';
const fmt=n=>n.toLocaleString('pt-BR',{maximumFractionDigits:4});
const nome=s=>!s?'Fora do livreto':s.origem?`Página ${s.origem}`:`Branca ${s.logica}`;

// Janela independente: recebe somente uma cópia do workspace, nunca o modifica.
export class JanelaLivreto {
    constructor(root,{snapshot}){
        this.snapshot=snapshot;this.version=0;this.sequence=Promise.resolve();this.page=1;this.zoom=1;
        const input=(id,label,value)=>`<label>${label}<input data-id="${id}" value="${value}" inputmode="decimal"></label>`;
        const select=(id,label,options)=>`<label>${label}<select data-id="${id}">${options.map(([v,t])=>`<option value="${v}">${t}</option>`).join('')}</select></label>`;
        const check=(id,label)=>`<label><input type="checkbox" data-id="${id}"> ${label}</label>`;
        this.dialog=document.createElement('dialog');this.dialog.className='pdf-montagem pdf-livreto';this.dialog.setAttribute('aria-label','Livreto grampeado');
        this.dialog.innerHTML=`<h2>Livreto grampeado — dobra central</h2><div class="montagem-corpo"><form>
        <p>Tipo: grampo canoa · um caderno · leitura da esquerda para a direita.</p><div class="montagem-grade">
        ${select('final','Formato final nominal',[['A5','A5 · 148 × 210 mm'],['A4','A4 · 210 × 297 mm']])}
        ${select('papel','Papel de impressão',[['A4','A4'],['A3','A3']])}
        ${select('orientacao','Orientação da folha',[['paisagem','Paisagem'],['retrato','Retrato (gira toda a montagem)']])}</div>
        <p data-id="contagem" role="status"></p><p data-id="brancas"></p><div data-id="aceite-brancas" hidden>${check('autorizarBrancas','Autorizo acrescentar as páginas em branco indicadas, somente no livreto.')}</div>
        <p data-id="medidas-resumo"></p><button type="button" data-id="ajustar">Ajustar conteúdo ao formato final</button>
        <details class="montagem-avancadas"><summary>Configurações avançadas</summary>
        <fieldset><legend>Escala e encaixe do conteúdo</legend><div class="montagem-grade">
        ${select('escalaModo','Dimensionar conteúdo',[['real','100% / tamanho real'],['ajustar','Ajustar proporcionalmente ao formato final'],['personalizada','Escala personalizada']])}
        ${input('escala','Escala (%)',100)}${select('rotacao','Giro adicional do conteúdo',[['0','0°'],['90','90°'],['180','180°'],['270','270°']])}</div>
        <p>O formato final não diminui. Ajustar pode ampliar ou reduzir o conteúdo; confira a escala efetiva abaixo.</p></fieldset>
        <fieldset><legend>Margens internas do conteúdo (mm)</legend><div class="montagem-grade">
        ${input('esquerda','Esquerda',0)}${input('direita','Direita',0)}${input('superior','Superior',0)}${input('inferior','Inferior',0)}</div><p>Relativas a cada página na ordem de leitura; não são margens espelhadas.</p></fieldset>
        <fieldset><legend>Marcas e sangria</legend>${check('marcasCorte','Marcas de corte externas')}${check('marcasDobra','Marcas de dobra externas')}
        <div class="montagem-grade">${input('sangria','Sangria existente (mm)',0)}${input('afastamento','Afastamento das marcas (mm)',3)}${input('comprimento','Comprimento das marcas (mm)',3)}${input('espessura','Traço (mm)',.1)}</div>
        <p>Não cria sangria. No encontro central, a sangria é limitada à própria metade para não cobrir a página vizinha. Marcas exigem papel maior ou serão recortadas.</p></fieldset>
        <fieldset><legend>Frente, verso e registro</legend>
        ${select('verso','Transformação do verso no PDF',[['normal','Verso na mesma orientação'],['girar180','Girar o verso 180°']])}
        <div class="montagem-grade">${input('registroX','Deslocar conteúdo do verso X (mm)',0)}${input('registroY','Deslocar conteúdo do verso Y (mm)',0)}</div>
        <p>Deslocamentos nos eixos físicos da folha. Sem alterar duplex do driver. Faça prova de 4 páginas; escolha a virada que mantenha ambos os lados em pé. Escala de impressão 100%, sem livreto/N-up no driver.</p></fieldset></details>
        <details><summary>Sequência de impressão e limites técnicos</summary><p data-id="sequencia"></p><p>Frente e verso consecutivos de cada folha, da externa para a interna. Sem compensação de espessura. Fidelidade de ICC, cores especiais e sobreimpressão requer conferência no RIP. Prévia de leitura não simula espessura, refile ou transparência do papel.</p></details></form>
        <section>${select('visualizacao','Conferir',[['impressao','Folhas de impressão'],['leitura','Ordem de leitura após dobrar']])}
        <nav><button type="button" data-id="anterior">Anterior</button><span data-id="folha"></span><button type="button" data-id="proxima">Próxima</button></nav>
        <nav class="montagem-zoom"><button type="button" data-id="menos">−</button><button type="button" data-id="enquadrar">Enquadrar</button><button type="button" data-id="mais">+</button><span data-id="zoom"></span></nav>
        <div class="montagem-previa"><canvas aria-label="Prévia real do livreto"></canvas></div>
        <p data-id="legenda"></p><div data-id="excesso" hidden><p data-id="excesso-texto"></p><button type="button" data-id="aceitar-corte">Autorizar corte somente das marcas/sangria externas</button></div>
        <div data-id="compatibilidade" hidden><p data-id="compat-detalhes"></p><button type="button" data-id="preparar">Preparar cópia para livreto</button></div>
        </section></div><p data-id="origem">Origem: cópia do PDFWorkspace atual, incluindo alterações pendentes.</p><button type="button" data-id="original" hidden>Voltar à cópia original</button>
        <p data-id="status" role="status"></p><footer><button type="button" data-id="cancelar">Cancelar</button><button type="button" data-id="tentar" hidden>Tentar novamente</button><button type="button" data-id="salvar">Gerar PDF de livreto</button></footer>`;
        root.append(this.dialog);this.get=id=>this.dialog.querySelector(`[data-id="${id}"]`);
        this.dialog.querySelector('form').onsubmit=e=>e.preventDefault();
        this.dialog.querySelector('form').oninput=e=>{if(e.target.dataset.id==='autorizarBrancas')this.buttons();else{this.controls();this.schedule();}};
        this.get('visualizacao').onchange=()=>{this.page=1;this.schedule(false);};
        this.get('anterior').onclick=()=>{if(this.page>1){this.page--;this.schedule(false);}};
        this.get('proxima').onclick=()=>{if(this.page<this.totalPrevia()){this.page++;this.schedule(false);}};
        for(const [id,mult]of [['menos',.8],['mais',1.25],['enquadrar',1]])this.get(id).onclick=()=>{this.zoom=id==='enquadrar'?1:Math.max(.25,Math.min(4,this.zoom*mult));this.schedule(false);};
        this.get('ajustar').onclick=()=>{this.get('escalaModo').value='ajustar';this.controls();this.schedule();};
        this.get('aceitar-corte').onclick=()=>{const r=this.result;if(this.busy||!r||r.plan.bloqueios.length)return;if(confirm('Recortar somente marcas ou sangria que ultrapassam a folha?\n\nO formato final e o conteúdo devem caber integralmente. Marcas fora do papel não serão impressas. Não altera papel nem escala.')){if(r===this.result){this.approved=r;this.buttons();}}};
        this.get('cancelar').onclick=()=>this.close();this.dialog.addEventListener('cancel',e=>{e.preventDefault();this.close();});
        this.get('tentar').onclick=()=>this.bytes?this.schedule():this.capture();
        this.get('salvar').onclick=()=>this.save();this.get('preparar').onclick=()=>this.prepare();
        this.get('original').onclick=()=>{this.bytes=this.original.slice();this.prepared=false;this.get('original').hidden=true;this.get('origem').textContent='Origem: cópia original do workspace; nenhuma alteração nele.';this.schedule();};
        this.controls();this.buttons();
    }
    controls(){this.get('escala').closest('label').hidden=this.get('escalaModo').value!=='personalizada';}
    config(){const n=id=>Number(this.get(id).value.trim().replace(',','.'));return {modo:'livreto',...Object.fromEntries(['final','papel','orientacao','verso','escalaModo'].map(k=>[k,this.get(k).value])),...Object.fromEntries(['escala','rotacao','sangria','afastamento','comprimento','espessura','registroX','registroY'].map(k=>[k,n(k)])),marcasCorte:this.get('marcasCorte').checked,marcasDobra:this.get('marcasDobra').checked,margensConteudo:Object.fromEntries(['esquerda','direita','superior','inferior'].map(k=>[k,n(k)]))};}
    async open(){if(this.closing)await this.closing;if(this.dialog.open)return;this.dialog.showModal();await this.capture();}
    async capture(){const v=++this.version;this.busy=true;this.buttons();this.get('status').textContent='Copiando documento atual…';try{const bytes=await this.snapshot();if(v!==this.version||!this.dialog.open)return;this.bytes=bytes.slice();this.original=bytes.slice();this.prepared=false;this.get('original').hidden=true;this.get('origem').textContent='Origem: cópia do PDFWorkspace atual, incluindo alterações pendentes.';this.get('autorizarBrancas').checked=false;this.page=1;this.schedule();}catch(e){if(v===this.version){this.get('status').textContent=e.message;this.get('tentar').hidden=false;}}}
    totalPrevia(){return this.result?(this.get('visualizacao').value==='leitura'?this.result.plan.leitura.totalFolhas:this.result.plan.totalFolhas):0;}
    canOutput(){const r=this.result;return !!r&&!this.busy&&!this.saving&&!r.plan.bloqueios.length&&(!r.plan.paginacao.adicionadas||this.get('autorizarBrancas').checked)&&(!r.geometria.requerConfirmacao||this.approved===r);}
    buttons(){this.get('salvar').disabled=!this.canOutput();for(const id of ['anterior','proxima'])this.get(id).disabled=this.busy||!this.result||(id==='anterior'?this.page<=1:this.page>=this.totalPrevia());this.get('aceitar-corte').disabled=this.busy||!this.result||!!this.result.plan.bloqueios.length;}
    schedule(generate=true){
        if(!this.bytes||!this.dialog.open||this.saving||this.preparing)return;
        clearTimeout(this.timer);this.abort?.abort();const v=++this.version;this.busy=true;if(generate){this.result=null;this.approved=null;}
        this.buttons();this.get('tentar').hidden=true;this.get('compatibilidade').hidden=true;this.get('status').textContent='Atualizando livreto…';
        const c=this.dialog.querySelector('canvas');c.width=c.height=0;const w=this.get('papel').value==='A3'?420:297,h=this.get('papel').value==='A3'?297:210;desenharFolhaVazia(c,{largura:this.get('orientacao').value==='paisagem'?w:h,altura:this.get('orientacao').value==='paisagem'?h:w});
        this.timer=setTimeout(()=>{this.sequence=this.sequence.catch(()=>{}).then(async()=>{
            if(v!==this.version||!this.dialog.open)return;const a=this.abort=new AbortController();
            try{
                if(generate){const {gerarMontagem}=await import('./pdf-imposition-engine.js');const r=await aguardarMontagem(gerarMontagem(this.bytes,this.config(),{signal:a.signal,previsualizarExcesso:true}),{signal:a.signal});if(v!==this.version)return;this.result=r;this.page=1;}
                this.summary();await this.render(a.signal,this.result,this.page);if(v!==this.version)return;this.busy=false;this.buttons();
            }catch(e){a.abort();if(v!==this.version)return;this.result=null;this.busy=false;this.buttons();this.get('status').textContent=e.name==='AbortError'?'Prévia cancelada.':e.message;this.get('tentar').hidden=false;if(e.diagnostico){this.diagnostico=e.diagnostico;this.get('compatibilidade').hidden=false;this.get('compat-detalhes').textContent=[...e.diagnostico.removiveis,...e.diagnostico.bloqueios].map(r=>r.descricao).join('\n');this.get('preparar').disabled=!e.diagnostico.podePreparar;}}
        });},180);
    }
    summary(){
        const r=this.result,p=r.plan,g=p.paginacao,scales=[...new Set(p.folhas.flat().map(v=>fmt(v.escala*100)))];
        this.get('contagem').textContent=`${g.originais} páginas originais + ${g.adicionadas} brancas = ${g.total} páginas · ${p.folhasFisicas} folhas físicas · ${p.totalFolhas} faces de impressão.`;
        this.get('brancas').textContent=g.adicionadas?`Proposta: brancas nas posições ${g.brancas.join(', ')}, ao final.`:'Não precisa acrescentar páginas em branco.';this.get('aceite-brancas').hidden=!g.adicionadas;
        this.get('medidas-resumo').textContent=`Final nominal ${p.final.largura} × ${p.final.altura} mm · folha ${p.folha.largura} × ${p.folha.altura} mm · após dobrar, antes do refile ${p.aposDobra.largura} × ${p.aposDobra.altura} mm. Escala efetiva do conteúdo: ${scales.join('%; ')}%. Giro adicional: ${this.get('rotacao').value}°. Verso: ${p.verso==='normal'?'normal':'giro 180°'}.`;
        this.get('sequencia').textContent=p.faces.map(f=>`Folha ${f.folha}, ${f.lado}: ${f.slots.map(nome).join(' | ')}`).join('\n');
        const edges=[...new Set(r.geometria.sheets.flatMap(s=>s.foraFolha))];this.get('excesso').hidden=!r.geometria.requerConfirmacao;
        this.get('excesso-texto').textContent=`Marcas/sangria podem ser recortadas nas bordas ${edges.join(', ')}. Escolha papel maior, reduza as marcas ou autorize somente esse corte. Conteúdo fora do formato final continua bloqueado.`;
        this.get('status').textContent=p.bloqueios.length?p.bloqueios.join(' '):'Prévia atualizada. O original e suas alterações pendentes permanecem intactos. Geração independente; nenhuma configuração de driver alterada.';
    }
    async render(signal,r,number){
        const reading=this.get('visualizacao').value==='leitura',value=reading?{bytes:r.bytesLeitura,plan:r.plan.leitura}:r,box=this.dialog.querySelector('.montagem-previa');
        const rendered=await renderizarFolhaMontagem(value,number,{signal,largura:Math.max(100,box.clientWidth-24),altura:Math.max(100,box.clientHeight-24),zoom:this.zoom});
        const {canvas,viewport}=rendered;if(signal.aborted||!this.dialog.open){canvas.width=canvas.height=0;return;}
        desenharInspecaoMontagem(canvas,viewport,value,number,{guias:true});
        // Molduras mostram também as páginas brancas; dobra é guia de tela, nunca tinta sobre a arte.
        const ctx=canvas.getContext('2d'),frames=value.plan.molduras[number-1],point=(x,y)=>viewport.convertToViewportPoint(x*72/25.4,y*72/25.4);
        ctx.save();ctx.strokeStyle='#326ba4';ctx.setLineDash([4,4]);ctx.lineWidth=1;
        for(const f of frames){const a=point(f.x,f.y+f.height),b=point(f.x+f.width,f.y);ctx.strokeRect(a[0],a[1],b[0]-a[0],b[1]-a[1]);}ctx.restore();
        const old=this.dialog.querySelector('canvas');old.replaceWith(canvas);old.width=old.height=0;
        const p=r.plan;this.get('folha').textContent=reading?`Abertura ${number}/${p.leitura.totalFolhas}`:`Folha ${p.faces[number-1].folha} · ${p.faces[number-1].lado} · face ${number}/${p.totalFolhas}`;
        this.get('legenda').textContent=(reading?p.paginacao.aberturas[number-1]:p.faces[number-1].slots).map(nome).join(' | ')+(reading?' · Simulação da ordem de leitura; brancos não recebem tinta.':' · Azul: conteúdo; hachura: fora da folha. Dobra ao meio; área imprimível não verificada.');this.get('zoom').textContent=`Zoom ${Math.round(this.zoom*100)}%`;
    }
    async save(){
        if(!this.canOutput())return;const r=this.result;if(!confirm('Gerar PDF de livreto?\n\n'+this.get('contagem').textContent+'\n'+this.get('medidas-resumo').textContent+'\n\nFrente e verso consecutivos, uma imposição só. Ao imprimir: escala 100%, livreto/N-up desligados no driver. A virada duplex exige prova física. Documento original preservado.')||r!==this.result||!this.canOutput())return;
        this.saving=true;this.lock(true);this.buttons();try{await salvarArquivoPDF(r.bytes,'Livreto_4PrintShops.pdf');this.get('status').textContent='PDF de livreto salvo. Origem e alterações pendentes preservadas; duplex físico ainda exige prova.';}catch(e){this.get('status').textContent=e.message;}finally{this.saving=false;this.lock(false);this.buttons();}
    }
    lock(value){for(const e of this.dialog.querySelectorAll('input,select,button'))e.disabled=value;this.get('escala').disabled=value||this.get('escalaModo').value!=='personalizada';}
    async prepare(){
        const d=this.diagnostico;if(this.preparing||!d?.podePreparar||!this.bytes)return;
        if(!confirm('Preparar cópia independente para livreto?\n\n'+d.removiveis.map(r=>r.descricao).join('\n')+'\n\nPode perder acessibilidade, navegação e links não imprimíveis apenas na cópia. Texto e aparência serão comparados; ActualText essencial e aparências não comprovadas continuam bloqueados. O workspace não será alterado.'))return;
        clearTimeout(this.timer);this.abort?.abort();const v=++this.version,a=this.abort=new AbortController(),bytes=this.bytes.slice();this.preparing=true;this.busy=true;this.lock(true);this.get('tentar').hidden=true;this.get('status').textContent='Comparando texto e aparência da cópia…';
        this.sequence=this.sequence.catch(()=>{}).then(async()=>{try{const {prepararCopiaMontagem}=await import('./pdf-imposition-compatibility.js'),{compararCopiaMontagem}=await import('./pdf-imposition-compare.js');const r=await prepararCopiaMontagem(bytes,{signal:a.signal,confirmado:true,validar:compararCopiaMontagem});if(v===this.version){this.bytes=r.bytes;this.prepared=true;this.get('original').hidden=false;this.get('origem').textContent='Cópia preparada e comparada; workspace original intacto.';}}catch(e){if(v===this.version){this.get('status').textContent=e.message;this.get('tentar').hidden=false;}}finally{this.preparing=false;this.lock(false);this.busy=false;this.buttons();if(v===this.version&&this.prepared)this.schedule();}});
    }
    async close(){if(this.saving)return;if(this.closing)return this.closing;clearTimeout(this.timer);++this.version;this.abort?.abort();this.dialog.close();this.closing=this.sequence.catch(()=>{}).then(()=>{this.bytes=this.original=this.result=this.approved=null;const c=this.dialog.querySelector('canvas');c.width=c.height=0;}).finally(()=>{this.closing=null;});return this.closing;}
    async destroy(){await this.close();this.dialog.remove();}
}
