import {salvarArquivoPDF} from '../app/save-file.js';
import {renderizarFolhaMontagem, aguardarMontagem} from './pdf-imposition-preview.js';
const presets={A4:[210,297],A3:[297,420],SRA3:[320,450]};
export class JanelaMontagem {
    constructor(root,{snapshot,imprimir}){
        this.snapshot=snapshot;this.imprimir=imprimir;this.sequence=Promise.resolve();this.version=0;
        this.dialog=document.createElement('dialog');this.dialog.className='pdf-montagem';this.dialog.setAttribute('aria-label','Montagem / Imposição');
        const input=(id,label,value,type='text')=>`<label>${label}<input data-id="${id}" type="${type}" value="${value}"></label>`;
        const select=(id,label,items)=>`<label>${label}<select data-id="${id}">${items.map(([v,t])=>`<option value="${v}">${t}</option>`).join('')}</select></label>`;
        this.dialog.innerHTML=`<h2>Montagem / Imposição</h2><div class="montagem-corpo"><form><div class="montagem-grade">
        ${select('modo','Modo',[['repeat','Step & Repeat'],['nup','N-up']])}${input('paginas','Origem: páginas/intervalos; 0 = vaga','1')}
        ${select('formato','Folha',[['A3','A3'],['A4','A4'],['SRA3','SRA3'],['custom','Personalizada']])}${select('orientacao','Orientação',[['retrato','Retrato'],['paisagem','Paisagem']])}
        ${input('largura','Folha: largura (mm)',297)}${input('altura','Folha: altura (mm)',420)}
        ${select('grade','Grade',[['auto','Automática (repetição)'],['2','2-up'],['4','4-up'],['6','6-up'],['8','8-up'],['custom','Personalizada']])}
        ${select('ordem','Preenchimento',[['horizontal','Horizontal'],['vertical','Vertical']])}
        ${input('colunas','Colunas',2)}${input('linhas','Linhas',2)}
        ${input('quantidade','Peças desejadas (repetição)',24)}${select('rotacao','Giro adicional',[['0','0°'],['90','90°'],['180','180°'],['270','270°']])}
        ${select('escalaModo','Escala',[['ajustar','Ajuste proporcional'],['real','100%'],['personalizada','Personalizada']])}${input('escala','Escala (%)',120)}
        ${input('pecaLargura','Repetição: largura final (mm; vazio = origem)','')}${input('pecaAltura','Altura final (mm; vazio = proporcional)','')}
        ${input('esquerda','Margem esquerda (mm)',10)}${input('direita','Margem direita (mm)',10)}${input('superior','Margem superior (mm)',10)}${input('inferior','Margem inferior (mm)',10)}
        ${input('espacoX','Intervalo horizontal entre reservas (mm)',0)}${input('espacoY','Intervalo vertical entre reservas (mm)',0)}
        ${input('sangria','Sangria existente a usar (mm)',0)}${input('comprimento','Comprimento das marcas (mm)',3)}${input('afastamento','Afastamento do corte (mm)',3)}${input('espessura','Traço das marcas (mm)',.1)}
        ${select('modoMarcas','Marcas de corte',[['individuais','Individuais'],['grupo','Apenas no grupo / contorno externo'],['bordas','Contorno externo + divisões'],['sem','Sem marcas']])}
        </div><p>Contorno + divisões: indica todos os cortes nas bordas externas, sem traços sobre as artes. Apenas no grupo mantém somente os cantos. Com sangria zero, intervalo zero permite peças encostadas.</p><label><input data-id="centralizar" type="checkbox" checked> Centralizar montagem</label><label><input data-id="rotacaoAutomatica" type="checkbox"> Giro automático para melhor ajuste N-up</label><label><input data-id="guias" type="checkbox" checked> Guias azuis de corte (somente prévia)</label>
        <p>A folha do PDF é independente do papel do driver. Sangria não é criada. Medidas finais devem manter a proporção. Em N-up, 0 reserva uma posição vazia.</p></form>
        <section><nav><button type="button" data-id="anterior">Anterior</button><span data-id="folha"></span><button type="button" data-id="proxima">Próxima</button></nav><div class="montagem-previa"><div data-id="compatibilidade" hidden role="alert"><strong>Este PDF contém estruturas especiais que precisam ser preparadas antes da montagem.</strong><p data-id="compat-resumo"></p><details><summary>Motivos técnicos</summary><p data-id="compat-detalhes"></p></details><button type="button" data-id="preparar">Preparar cópia para montagem</button></div><canvas></canvas></div></section></div>
        <div><span data-id="origem"></span> <button type="button" data-id="original" hidden>Voltar ao original</button></div>
        <p data-id="status" role="status"></p><footer><button type="button" data-id="cancelar">Cancelar</button><button type="button" data-id="tentar" hidden>Tentar novamente</button><button type="button" data-id="salvar">Gerar PDF de montagem</button><button type="button" data-id="imprimir">Imprimir montagem</button></footer>`;
        root.append(this.dialog);this.get=id=>this.dialog.querySelector(`[data-id="${id}"]`);
        this.dialog.querySelector('form').addEventListener('submit',e=>e.preventDefault());
        this.dialog.querySelector('form').addEventListener('input',e=>{
            if(e.target.dataset.id==='modo'&&this.get('modo').value==='nup'&&this.get('grade').value==='auto')this.get('grade').value='4';
            if(['largura','altura'].includes(e.target.dataset.id))this.get('formato').value='custom';
            if(['formato','orientacao'].includes(e.target.dataset.id)&&presets[this.get('formato').value]){let [w,h]=presets[this.get('formato').value];if(this.get('orientacao').value==='paisagem')[w,h]=[h,w];this.get('largura').value=w;this.get('altura').value=h;}
            this.controls();this.schedule(e.target.dataset.id!=='guias');
        });
        this.get('cancelar').onclick=()=>this.close();this.dialog.addEventListener('cancel',e=>{e.preventDefault();this.close();});
        this.get('anterior').onclick=()=>{if(this.result&&this.page>1){this.page--;this.schedule(false);}};
        this.get('proxima').onclick=()=>{if(this.result&&this.page<this.result.plan.totalFolhas){this.page++;this.schedule(false);}};
        this.get('salvar').onclick=()=>this.save();
        this.get('tentar').onclick=()=>this.bytes?this.schedule():this.capture();
        this.get('preparar').onclick=()=>this.prepare();
        this.get('original').onclick=()=>{if(this.preparing||this.saving)return;this.bytes=this.originalBytes.slice();this.prepared=null;this.sourceState();this.schedule();};
        this.get('imprimir').onclick=async()=>{if(!this.result||this.busy)return;try{if(!confirm('Imprimir a montagem já imposta? Confira papel, escala 100% e desative N-up/livreto no driver para evitar transformação dupla. As opções privadas do driver não são interpretadas integralmente.'))return;const result=this.result;await this.close();await this.imprimir?.(result);}catch(e){this.get('status').textContent=e.message;}};
    }
    async open(){
        if(this.closing)await this.closing;
        if(this.dialog.open)return;
        this.dialog.showModal();await this.capture();
    }
    async capture(){
        this.get('status').textContent='Copiando o documento atual…';this.buttons(true);this.get('tentar').hidden=true;
        const version=++this.version;
        try{const bytes=await this.snapshot();if(version!==this.version||!this.dialog.open)return;this.originalBytes=bytes.slice();this.bytes=bytes;this.prepared=null;this.sourceState();this.page=1;this.controls();this.schedule();}catch(e){if(version===this.version&&this.dialog.open){this.get('status').textContent=e.message;this.buttons(true);this.get('tentar').hidden=false;}}
    }
    config(){
        const n=id=>Number(this.get(id).value.trim().replace(',','.'));
        const paginas=[];
        for(const item of this.get('paginas').value.split(',')){const m=item.trim().match(/^(\d+)(?:\s*-\s*(\d+))?$/);if(!m)throw Error('Informe páginas como 1,3-5; use 0 para vaga.');const a=Number(m[1]),b=Number(m[2]??m[1]);if(b<a||b-a>10000)throw Error('Intervalo inválido.');for(let p=a;p<=b;p++)paginas.push(p||null);}
        let colunas=n('colunas'),linhas=n('linhas');const grade=this.get('grade').value;
        if(['2','4','6','8'].includes(grade)){colunas=2;linhas=Number(grade)/2;}
        return {modo:this.get('modo').value,paginas,largura:n('largura'),altura:n('altura'),automatico:grade==='auto',colunas,linhas,quantidade:n('quantidade'),ordem:this.get('ordem').value,rotacao:n('rotacao'),rotacaoAutomatica:this.get('rotacaoAutomatica').checked,escalaModo:this.get('escalaModo').value,escala:n('escala'),pecaLargura:this.get('pecaLargura').value.trim()?n('pecaLargura'):undefined,pecaAltura:this.get('pecaAltura').value.trim()?n('pecaAltura'):undefined,margens:Object.fromEntries(['esquerda','direita','superior','inferior'].map(k=>[k,n(k)])),espacoX:n('espacoX'),espacoY:n('espacoY'),sangria:n('sangria'),comprimento:n('comprimento'),afastamento:n('afastamento'),espessura:n('espessura'),modoMarcas:this.get('modoMarcas').value,centralizar:this.get('centralizar').checked};
    }
    controls(){const repeat=this.get('modo').value==='repeat';for(const id of ['pecaLargura','pecaAltura','quantidade'])this.get(id).disabled=!repeat;this.get('escala').disabled=this.get('escalaModo').value!=='personalizada';for(const id of ['colunas','linhas'])this.get(id).disabled=this.get('grade').value!=='custom';this.get('rotacaoAutomatica').disabled=repeat;}
    buttons(busy){this.busy=busy;for(const id of ['salvar','imprimir','anterior','proxima'])this.get(id).disabled=busy||!this.result;this.get('imprimir').hidden=!this.imprimir;}
    sourceState(){this.get('origem').textContent=this.prepared?'Cópia independente preparada e validada para montagem. O workspace original permanece intacto.':'Origem: cópia do workspace atual; nenhuma preparação aplicada.';this.get('original').hidden=!this.prepared;}
    compatibility(error,allow=true){
        const diag=error.diagnostico;this.diagnostico=diag;
        this.get('compatibilidade').hidden=false;
        this.get('compat-resumo').textContent=diag.podePreparar&&allow?'É possível tentar uma cópia vetorial, com confirmação e comparação de texto e aparência.':'Há recursos cuja aparência ou texto ainda não podem ser preparados com segurança. A montagem continuará bloqueada.';
        this.get('compat-detalhes').textContent=[...diag.removiveis.map(r=>r.descricao),...diag.bloqueios.map(b=>b.descricao)].join('\n');
        this.get('preparar').disabled=!allow||!diag.podePreparar;
        this.get('tentar').hidden=true;
    }
    async prepare(){
        if(this.preparing||!this.diagnostico?.podePreparar||!this.bytes)return;
        if(!confirm('Preparar uma cópia independente para montagem?\n\nPoderão ser removidos:\n'+this.diagnostico.removiveis.map(r=>'• '+r.descricao).join('\n')+'\n\nAcessibilidade, navegação e links não imprimíveis podem ser perdidos somente na cópia. ActualText exige validação específica. Texto e aparência de impressão serão comparados antes/depois. Não haverá rasterização. Formulários e aparências não comprovadas continuam bloqueados.\n\nO original e as alterações do workspace serão preservados. Continuar?'))return;
        clearTimeout(this.timer);this.abort?.abort();const version=++this.version,abort=this.abort=new AbortController();const bytes=this.bytes.slice();
        this.preparing=true;this.result=null;this.buttons(true);this.get('preparar').disabled=true;this.get('original').disabled=true;this.get('status').textContent='Preparando cópia e comparando texto/aparência de todas as páginas…';
        const fields=[...this.dialog.querySelectorAll('form input,form select')];for(const field of fields)field.disabled=true;
        this.sequence=this.sequence.catch(()=>{}).then(async()=>{
            if(version!==this.version||abort.signal.aborted)return;
            try{
                const {prepararCopiaMontagem}=await import('./pdf-imposition-compatibility.js');const {compararCopiaMontagem}=await import('./pdf-imposition-compare.js');
                const prepared=await prepararCopiaMontagem(bytes,{signal:abort.signal,confirmado:true,validar:compararCopiaMontagem});
                if(version!==this.version||abort.signal.aborted)return;
                this.bytes=prepared.bytes;this.prepared=prepared;this.sourceState();this.schedule();
            }catch(error){if(version===this.version){this.get('status').textContent=error.message;if(error.diagnostico)this.compatibility(error,false);else{this.get('compat-resumo').textContent=error.message;this.get('preparar').disabled=false;}}}
            finally{this.preparing=false;for(const field of fields)field.disabled=false;this.controls();this.get('original').disabled=false;}
        });
        return this.sequence;
    }
    schedule(generate=true){
        generate=generate||!this.result;
        this.abort?.abort();clearTimeout(this.timer);const version=++this.version;this.buttons(true);if(generate)this.result=null;
        this.get('status').textContent='Atualizando montagem…';this.get('tentar').hidden=true;this.get('compatibilidade').hidden=true;const visible=this.dialog.querySelector('canvas');visible.width=visible.height=0;
        this.timer=setTimeout(()=>{this.sequence=this.sequence.catch(()=>{}).then(async()=>{
            if(version!==this.version||!this.dialog.open)return;
            const abort=this.abort=new AbortController();
            try{
                if(generate){const {gerarMontagem}=await import('./pdf-imposition-engine.js');const result=await aguardarMontagem(gerarMontagem(this.bytes,this.config(),{signal:abort.signal}),{signal:abort.signal});if(version!==this.version)return;this.result=result;this.page=1;}
                if(!this.result)throw Error('Montagem indisponível. Tente novamente.');
                const result=this.result,number=this.page;await this.render(abort.signal,result,number);
                if(version!==this.version)return;
                const p=this.result.plan;const sizes=[...new Set(p.folhas.flat().map(piece=>`${Number(piece.largura.toFixed(2))} × ${Number(piece.altura.toFixed(2))} mm`))];this.get('status').textContent=`Folha ${p.folha.largura} × ${p.folha.altura} mm · útil ${p.util.largura} × ${p.util.altura} mm · peças: ${sizes.slice(0,3).join('; ')}${sizes.length>3?'…':''} · ${p.colunas} × ${p.linhas} = ${p.capacidade} posições · ${p.quantidade} posições solicitadas · ${p.totalFolhas} folha(s). ${this.result.avisos.join(' ')}`;
                this.buttons(false);
            }catch(e){abort.abort();if(version===this.version){this.result=null;this.buttons(true);this.get('status').textContent=e.name==='AbortError'?'Montagem cancelada.':e.message;this.get('tentar').hidden=false;if(e.diagnostico)this.compatibility(e);}}
        });},250);
    }
    async render(signal,result=this.result,number=this.page){
        const {canvas,viewport,pages}=await renderizarFolhaMontagem(result,number,{signal});
        if(signal.aborted||!this.dialog.open){canvas.width=canvas.height=0;return;}
        if(this.get('guias').checked){const ctx=canvas.getContext('2d');ctx.strokeStyle='#0077cc';ctx.lineWidth=1;ctx.setLineDash([4,4]);const p=result.plan;for(const piece of p.folhas[number-1])ctx.strokeRect(piece.x/p.folha.largura*viewport.width,(p.folha.altura-piece.y-piece.altura)/p.folha.altura*viewport.height,piece.largura/p.folha.largura*viewport.width,piece.altura/p.folha.altura*viewport.height);ctx.setLineDash([]);}
        const previous=this.dialog.querySelector('canvas');previous.replaceWith(canvas);previous.width=previous.height=0;
        this.get('folha').textContent=`${number} / ${pages}`;
    }
    async save(){if(!this.result||this.busy)return;this.buttons(true);this.saving=true;const fields=[...this.dialog.querySelectorAll('form input,form select')];for(const field of fields)field.disabled=true;this.get('cancelar').disabled=true;try{await salvarArquivoPDF(this.result.bytes,'Montagem_4PrintShops.pdf');this.get('status').textContent='PDF de montagem salvo. Documento original e alterações pendentes preservados.';}catch(e){this.get('status').textContent=e.message;}finally{this.saving=false;for(const field of fields)field.disabled=false;this.get('cancelar').disabled=false;this.controls();if(this.dialog.open)this.buttons(false);}}
    async close(){if(this.saving)return;if(this.closing)return this.closing;clearTimeout(this.timer);++this.version;this.abort?.abort();this.dialog.close();this.closing=this.sequence.catch(()=>{}).then(()=>{this.bytes=this.originalBytes=this.prepared=this.result=null;this.dialog.querySelector('canvas').width=0;}).finally(()=>{this.closing=null;});return this.closing;}
    async destroy(){await this.close();this.dialog.remove();}
}
