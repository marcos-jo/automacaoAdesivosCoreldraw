import {salvarArquivoPDF} from '../app/save-file.js';
import {renderizarFolhaMontagem, aguardarMontagem} from './pdf-imposition-preview.js';
import {analisarMontagem,sugerirAjusteMontagem} from './pdf-imposition-bounds.js';
import {desenharInspecaoMontagem,desenharFolhaVazia} from './pdf-imposition-view.js';
import {ConfiguracoesMontagem,camposPorModo} from './pdf-imposition-settings.js';
import {resumirMontagem} from './pdf-imposition-summary.js';
import {selecionarPaginasMontagem} from './pdf-imposition-selection.js';
import {agruparAvisosMontagem,textoAvisosMontagem,AvisosMontagem} from './pdf-imposition-warnings.js';
import {AlternativaImagemMontagem} from './pdf-imposition-raster-ui.js';
const presets={A4:[210,297],A3:[297,420],SRA3:[320,450]};
export class JanelaMontagem {
    constructor(root,{snapshot,imprimir,livreto,areaImprimivel=()=>null}){
        this.snapshot=snapshot;this.imprimir=imprimir;this.areaImprimivel=areaImprimivel;this.sequence=Promise.resolve();this.version=0;this.zoom=1;this.fatorMontagem=1;
        this.dialog=document.createElement('dialog');this.dialog.className='pdf-montagem';this.dialog.setAttribute('aria-label','Montagem / Imposição');
        const input=(id,label,value,type='text')=>`<label>${label}<input data-id="${id}" type="${type}" value="${value}"></label>`;
        const select=(id,label,items)=>`<label>${label}<select data-id="${id}">${items.map(([v,t])=>`<option value="${v}">${t}</option>`).join('')}</select></label>`;
        this.dialog.innerHTML=`<h2>Montagem / Imposição</h2><div class="montagem-corpo"><form><div class="montagem-grade" data-id="basico">
        ${select('modo','Tipo de montagem',[['repeat','Step & Repeat'],['nup','N-up']])}<div data-id="origem-repeat">${input('paginas','Página de origem','1')}</div>
        ${select('formato','Tamanho da folha',[['A3','A3'],['A4','A4'],['SRA3','SRA3'],['custom','Personalizada']])}${select('orientacao','Orientação da folha',[['retrato','Retrato'],['paisagem','Paisagem']])}
        <div class="montagem-grade montagem-personalizada" data-id="folha-personalizada" hidden>${input('largura','Largura da folha (mm)',297)}${input('altura','Altura da folha (mm)',420)}</div>
        ${input('quantidade','Quantidade de peças desejadas',24)}
        <div data-id="nup-opcoes" class="montagem-nup" hidden><p data-id="total-paginas" role="status"></p><div data-id="grade-nup"></div>${select('nupDistribuicao','Distribuição N-up',[['celulas','Em células (tradicional)'],['compacto','Compacto (peças encostadas)']])}<div data-id="giro-compacto"></div><p data-id="nup-explicacao"></p></div>
        </div><p data-id="distribuicao" role="status">A distribuição será exibida após atualizar a prévia.</p>
        <div data-id="avisos-detalhes" class="montagem-aviso-detalhes" role="region" tabindex="0" hidden><strong>Detalhes dos avisos</strong><div data-id="avisos-lista"></div><button type="button" data-id="avisos-recolher">Ocultar detalhes</button></div>
        <p data-id="medidas-resumo" role="status"></p><button type="button" data-id="tamanho-real">Voltar a 100% / tamanho real</button>
        <details data-id="medidas-paginas" hidden><summary>Medidas e escala por página</summary><p data-id="medidas-detalhes"></p></details>
        <details class="montagem-avancadas" data-id="avancadas"><summary>Configurações avançadas</summary>
        <fieldset data-id="selecao-nup" hidden><legend>Páginas do N-up</legend>${select('nupSelecao','Usar páginas',[['todas','Todas as páginas do documento'],['especificas','Selecionar páginas específicas']])}<div data-id="origem-nup"></div><p>As páginas seguem a ordem indicada, sem repetição automática. A última folha pode ficar incompleta. Zero não representa uma página.</p></fieldset>
        <fieldset><legend>Distribuição</legend><div class="montagem-grade">
        <div data-id="grade-avancada">${select('grade','Grade',[['auto','Automática (repetição)'],['2','2-up'],['4','4-up'],['6','6-up'],['8','8-up'],['custom','Personalizada']])}</div>
        ${select('ordem','Preenchimento',[['horizontal','Horizontal'],['vertical','Vertical']])}
        ${input('colunas','Colunas',2)}${input('linhas','Linhas',2)}
        </div></fieldset>
        <fieldset><legend>Escala e peças</legend><div class="montagem-grade">
        ${select('rotacao','Giro adicional',[['0','0°'],['90','90°'],['180','180°'],['270','270°']])}
        ${select('escalaModo','Política de escala',[['real','Manter tamanho real (100%)'],['reduzir','Reduzir somente se necessário'],['preencher','Preencher células (permite ampliação)'],['personalizada','Escala personalizada'],['medidas','Definir tamanho final da peça (repetição)']])}${input('escala','Escala manual (%)',100)}
        ${input('pecaLargura','Repetição: largura final (mm; vazio = origem)','')}${input('pecaAltura','Altura final (mm; vazio = proporcional)','')}
        </div><p data-id="politica-ajuda"></p></fieldset><fieldset><legend>Margens e espaços</legend><div class="montagem-grade">
        ${input('esquerda','Margem esquerda (mm)',0)}${input('direita','Margem direita (mm)',0)}${input('superior','Margem superior (mm)',0)}${input('inferior','Margem inferior (mm)',0)}
        ${input('espacoX','Intervalo horizontal entre reservas (mm)',0)}${input('espacoY','Intervalo vertical entre reservas (mm)',0)}
        </div></fieldset><fieldset><legend>Marcas de corte e sangria</legend><div class="montagem-grade">
        ${input('sangria','Sangria existente a usar (mm)',0)}${input('comprimento','Comprimento das marcas (mm)',3)}${input('afastamento','Afastamento do corte (mm)',3)}${input('espessura','Traço das marcas (mm)',.1)}
        ${select('modoMarcas','Marcas de corte',[['individuais','Individuais'],['grupo','Apenas no grupo / contorno externo'],['bordas','Contorno externo + divisões'],['sem','Sem marcas']])}
        </div><p>Contorno + divisões: indica todos os cortes nas bordas externas, sem traços sobre as artes. Apenas no grupo mantém somente os cantos. Com sangria zero, intervalo zero permite peças encostadas.</p></fieldset>
        <fieldset><legend>Posicionamento e guias</legend><label><input data-id="centralizar" type="checkbox" checked> Centralizar montagem</label><div data-id="giro-avancado"><label><input data-id="rotacaoAutomatica" type="checkbox"> Permitir giro automático de 0° ou 90° no N-up</label></div><label><input data-id="guias" type="checkbox" checked> Guias azuis de corte (somente prévia)</label></fieldset>
        <p>A folha do PDF é independente do papel do driver. Sangria não é criada. Medidas finais devem manter a proporção.</p></details></form>
        <section><nav><button type="button" data-id="anterior">Anterior</button><span data-id="folha"></span><button type="button" data-id="proxima">Próxima</button></nav>
        <nav class="montagem-zoom" aria-label="Enquadramento da montagem"><button type="button" data-id="enquadrar">Ver tudo</button><button type="button" data-id="enquadrar-folha">Folha inteira</button><button type="button" data-id="menos" aria-label="Reduzir prévia">−</button><span data-id="zoom" title="Ampliação visual relativa ao enquadramento; não altera o tamanho impresso">Zoom da prévia: 100%</span><button type="button" data-id="mais" aria-label="Ampliar prévia">+</button></nav>
        <div data-id="compatibilidade" hidden role="alert"><strong>Montagem bloqueada por compatibilidade.</strong><p data-id="compat-resumo"></p><details><summary>Motivos técnicos</summary><p data-id="compat-detalhes"></p></details><button type="button" data-id="preparar">Preparar cópia para montagem</button></div>
        <div class="montagem-previa"><p data-id="previa-bloqueada" role="status" hidden>Prévia bloqueada por compatibilidade. O documento original não está vazio; esta folha ainda não é uma montagem validada.</p><canvas aria-label="Folha física e montagem; hachura vermelha fora do papel"></canvas></div>
        <p data-id="aproveitamento"></p><p data-id="legenda">Contorno: papel · tracejado ocre: margens · azul: peças · hachura vermelha: fora do papel. Área imprimível: não verificada.</p>
        <div data-id="excesso" class="montagem-aviso" hidden><p data-id="excesso-texto" role="status"></p><div><button type="button" data-id="avisos-ver" aria-expanded="false">Ver detalhes</button><button type="button" data-id="aceitar-corte">Continuar com corte nas bordas</button></div></div><div class="montagem-ajustes"><button type="button" data-id="ajustar">Ajustar montagem para caber</button><button type="button" data-id="desfazer-ajuste" hidden>Remover ajuste de escala</button></div><p data-id="encaixe"></p></section></div>
        <div><span data-id="origem"></span> <button type="button" data-id="original" hidden>Voltar ao original</button></div>
        <p data-id="status" role="status"></p><footer><button type="button" data-id="cancelar">Cancelar</button><button type="button" data-id="tentar" hidden>Tentar novamente</button><button type="button" data-id="marcas-externas" hidden>Usar marcas externas + divisões</button><button type="button" data-id="salvar">Gerar PDF de montagem</button><button type="button" data-id="imprimir">Imprimir montagem</button></footer>`;
        root.append(this.dialog);this.get=id=>this.dialog.querySelector(`[data-id="${id}"]`);
        this.avisos=new AvisosMontagem(this.get);
        this.imagem=new AlternativaImagemMontagem(this.get,()=>this.prepare(true));
        if(livreto)this.get('modo').add(new Option('Livreto grampeado','livreto'));
        this.opcoes=new ConfiguracoesMontagem(this.lerOpcoes());this.ajusteBase=null;
        this.dialog.querySelector('form').addEventListener('submit',e=>e.preventDefault());
        this.dialog.querySelector('form').addEventListener('input',e=>{
            const id=e.target.dataset.id;
            if(id==='modo'&&this.get('modo').value==='livreto'){this.get('modo').value=this.opcoes.modo;Promise.resolve(livreto()).catch(e=>{this.get('status').textContent=e.message;});return;}
            if(id==='modo'){
                const saved=this.opcoes.trocar(this.get('modo').value,this.lerOpcoes(),this.fatorMontagem,this.ajusteBase);
                this.restaurarOpcoes(saved.values);this.fatorMontagem=saved.fator;this.ajusteBase=saved.base;
            }else if(id==='paginas'&&this.get('modo').value==='nup'){
                this.get('nupSelecao').value='especificas';
            }else if(['escalaModo','escala','pecaLargura','pecaAltura','grade','colunas','linhas','nupDistribuicao'].includes(id)){
                this.limparAjuste(!['grade','colunas','linhas'].includes(id));
                if(id==='escalaModo'&&this.get('escalaModo').value!=='medidas')for(const k of ['pecaLargura','pecaAltura'])this.get(k).value='';
                if(['pecaLargura','pecaAltura'].includes(id))this.get('escalaModo').value=this.get('pecaLargura').value.trim()||this.get('pecaAltura').value.trim()?'medidas':'real';
            }
            if(['largura','altura'].includes(e.target.dataset.id))this.get('formato').value='custom';
            if(['formato','orientacao'].includes(e.target.dataset.id)&&presets[this.get('formato').value]){let [w,h]=presets[this.get('formato').value];if(this.get('orientacao').value==='paisagem')[w,h]=[h,w];this.get('largura').value=w;this.get('altura').value=h;}
            this.controls();this.schedule(e.target.dataset.id!=='guias');
        });
        this.get('cancelar').onclick=()=>this.close();this.dialog.addEventListener('cancel',e=>{e.preventDefault();this.close();});
        this.get('anterior').onclick=()=>{if(this.result&&this.page>1){this.page--;this.schedule(false);}};
        this.get('proxima').onclick=()=>{if(this.result&&this.page<this.result.plan.totalFolhas){this.page++;this.schedule(false);}};
        this.get('salvar').onclick=()=>this.save();
        this.get('marcas-externas').onclick=()=>{if(this.preparing||this.saving)return;this.get('modoMarcas').value='bordas';this.controls();this.schedule();};
        this.get('aceitar-corte').onclick=()=>this.acceptClipping();
        this.get('ajustar').onclick=()=>this.fit();
        this.get('desfazer-ajuste').onclick=()=>{if(this.busy||this.saving)return;this.limparAjuste(true,true);this.controls();this.schedule();};
        this.get('tamanho-real').onclick=()=>{if(this.saving||this.preparing)return;this.limparAjuste();this.get('escalaModo').value='real';this.get('escala').value='100';for(const id of ['pecaLargura','pecaAltura'])this.get(id).value='';this.controls();this.schedule();};
        for(const [id,value] of [['mais',1.25],['menos',.8],['enquadrar',1],['enquadrar-folha',1]])this.get(id).onclick=()=>{if(this.saving||this.preparing)return;this.zoom=id==='enquadrar'?1:id==='enquadrar-folha'?this.paperZoom():Math.max(.25,Math.min(8,(this.zoom||1)*value));this.centerPaper=id==='enquadrar-folha';this.schedule(false);};
        this.get('tentar').onclick=()=>this.bytes?this.schedule():this.capture();
        this.get('preparar').onclick=()=>this.prepare();
        this.get('original').onclick=()=>{if(this.preparing||this.saving)return;this.bytes=this.originalBytes.slice();this.prepared=null;this.sourceState();this.schedule();};
        this.get('imprimir').onclick=async()=>{if(!this.canOutput())return;try{const result=this.result;if(!confirm('Imprimir a montagem já imposta?\n\n'+resumirMontagem(result).confirmacao)||result!==this.result||!this.canOutput())return;await this.close();await this.imprimir?.(result);}catch(e){this.get('status').textContent=e.message;}};
    }
    lerOpcoes(){return Object.fromEntries(camposPorModo.map(id=>{const field=this.get(id);return [id,field.type==='checkbox'?field.checked:field.value];}));}
    restaurarOpcoes(values){for(const [id,value]of Object.entries(values)){const field=this.get(id);if(field.type==='checkbox')field.checked=value;else field.value=value;}}
    limparAjuste(restaurarGrade=true,restaurarEscala=false){if(this.ajusteBase){const values=Object.fromEntries(Object.entries(this.ajusteBase).filter(([k])=>['escalaModo','escala'].includes(k)?restaurarEscala:restaurarGrade));this.restaurarOpcoes(values);}this.fatorMontagem=1;this.ajusteBase=null;}
    async open(){
        if(this.closing)await this.closing;
        if(this.dialog.open)return;
        this.dialog.showModal();await this.capture();
    }
    async capture(){
        this.abort?.abort();clearTimeout(this.timer);this.result=null;this.approved=null;this.selectionError=null;this.dirty=true;this.totalPaginas=null;
        const old=this.dialog.querySelector('canvas');old.width=old.height=0;
        this.get('status').textContent='Copiando o documento atual…';this.buttons(true);this.get('tentar').hidden=true;
        const version=++this.version;
        try{const bytes=await this.snapshot();const {PDFDocument}=await import('/vendor/pdf-lib/pdf-lib.js');const source=await PDFDocument.load(bytes.slice(),{updateMetadata:false});if(version!==this.version||!this.dialog.open)return;this.totalPaginas=source.getPageCount();this.limparAjuste(true,true);this.opcoes.limparAjustes();this.originalBytes=bytes.slice();this.bytes=bytes;this.prepared=null;this.sourceState();this.page=1;this.controls();this.schedule();}catch(e){if(version===this.version&&this.dialog.open){this.get('status').textContent=e.message;this.buttons(false);this.get('tentar').hidden=false;}}
    }
    config(){
        const n=id=>Number(this.get(id).value.trim().replace(',','.'));
        const paginas=selecionarPaginasMontagem({modo:this.get('modo').value,selecao:this.get('nupSelecao').value,texto:this.get('paginas').value,total:this.totalPaginas});
        let colunas=n('colunas'),linhas=n('linhas');const grade=this.get('grade').value;
        if(['2','4','6','8'].includes(grade)){colunas=2;linhas=Number(grade)/2;}
        return {nupDistribuicao:this.get('nupDistribuicao').value,fatorMontagem:this.fatorMontagem||1,modo:this.get('modo').value,paginas,largura:n('largura'),altura:n('altura'),automatico:grade==='auto',colunas,linhas,quantidade:n('quantidade'),ordem:this.get('ordem').value,rotacao:n('rotacao'),rotacaoAutomatica:this.get('rotacaoAutomatica').checked,escalaModo:this.get('escalaModo').value==='medidas'?'real':this.get('escalaModo').value,escala:n('escala'),pecaLargura:this.get('pecaLargura').value.trim()?n('pecaLargura'):undefined,pecaAltura:this.get('pecaAltura').value.trim()?n('pecaAltura'):undefined,margens:Object.fromEntries(['esquerda','direita','superior','inferior'].map(k=>[k,n(k)])),espacoX:n('espacoX'),espacoY:n('espacoY'),sangria:n('sangria'),comprimento:n('comprimento'),afastamento:n('afastamento'),espessura:n('espessura'),modoMarcas:this.get('modoMarcas').value,centralizar:this.get('centralizar').checked};
    }
    paperConfig(){const n=id=>Number(this.get(id).value.trim().replace(',','.'));return {largura:n('largura'),altura:n('altura'),margens:Object.fromEntries(['esquerda','direita','superior','inferior'].map(k=>[k,n(k)]))};}
    controls(){const repeat=this.get('modo').value==='repeat',mode=this.get('escalaModo').value,compact=!repeat&&this.get('nupDistribuicao').value==='compacto';for(const id of ['pecaLargura','pecaAltura','quantidade']){const show=repeat&&(id==='quantidade'||mode==='medidas');this.get(id).disabled=!show;this.get(id).closest('label').hidden=!show;}this.get('folha-personalizada').hidden=this.get('formato').value!=='custom';this.get('escala').disabled=mode!=='personalizada';this.get('escala').closest('label').hidden=mode!=='personalizada';for(const policy of ['reduzir','preencher'])this.get('escalaModo').querySelector(`option[value=${policy}]`).disabled=repeat;this.get('escalaModo').querySelector('option[value=medidas]').disabled=!repeat;this.get('grade').querySelector('option[value=auto]').disabled=!repeat;for(const id of ['colunas','linhas'])this.get(id).disabled=this.get('grade').value!=='custom';this.get('rotacaoAutomatica').disabled=repeat;
        this.get('politica-ajuda').textContent=repeat?'Repetição conserva escala e tamanho final explícitos.':mode==='preencher'?(compact?'Escolha explícita: amplia ou reduz o bloco compacto à área útil, mantendo as relações entre as peças.':'Escolha explícita: amplia ou reduz cada arte proporcionalmente à sua célula.'):mode==='reduzir'?'Nunca amplia acima de 100%. Pode reduzir proporcionalmente para caber.':'Não altera dimensões automaticamente. Se não couber, ajuste explicitamente a montagem ou escolha papel maior.';
        this.get('nup-opcoes').hidden=repeat;const label=this.get('rotacaoAutomatica').closest('label'),target=this.get(compact?'giro-compacto':'giro-avancado');if(label.parentElement!==target)target.append(label);
        const origin=this.get('paginas').closest('label'),originTarget=this.get(repeat?'origem-repeat':'origem-nup');if(origin.parentElement!==originTarget)originTarget.append(origin);
        origin.hidden=!repeat&&this.get('nupSelecao').value!=='especificas';origin.firstChild.textContent=repeat?'Página de origem':'Páginas/intervalos (ex.: 1,3-5)';
        this.get('origem-repeat').hidden=!repeat;this.get('selecao-nup').hidden=repeat;
        this.get('total-paginas').textContent=`Documento atual: ${this.totalPaginas??'…'} página(s). `+(this.get('nupSelecao').value==='todas'?'Todas serão utilizadas na ordem original.':'Seleção específica nas configurações avançadas.');
        const grade=this.get('grade').closest('label'),gradeTarget=this.get(repeat?'grade-avancada':'grade-nup');if(grade.parentElement!==gradeTarget)gradeTarget.append(grade);
        this.get('nup-explicacao').textContent=compact?'Posiciona pelas medidas finais, sem células. Intervalo zero encosta peças sem sangria na direção de preenchimento; tamanhos mistos podem deixar vazios. Giro só se permitido. Marcas externas não reduzem as artes; confira os avisos. Individuais exigem espaço.':'Centraliza cada página na célula. Pode haver folga interna mesmo com intervalo zero.';
        for(const [id,axis]of [['espacoX','horizontal'],['espacoY','vertical']])this.get(id).closest('label').firstChild.textContent=`Intervalo ${axis} ${compact?'entre cortes':'entre reservas'} (mm)`;
    }
    canOutput(){return !!this.result&&!this.busy&&!this.selectionError&&!this.dirty&&(!this.result.geometria?.requerConfirmacao||(this.result.config?.modo!=='nup'&&this.approved===this.result));}
    buttons(busy){this.busy=busy;for(const id of ['salvar','imprimir'])this.get(id).disabled=!this.canOutput();for(const id of ['anterior','proxima','aceitar-corte','ajustar','desfazer-ajuste'])this.get(id).disabled=busy||!this.result||!!this.selectionError||!!this.dirty;this.get('imprimir').hidden=!this.imprimir;}
    acceptClipping(){if(this.busy||this.result?.config?.modo==='nup'||!this.result?.geometria?.requerConfirmacao)return false;const result=this.result,details=textoAvisosMontagem(agruparAvisosMontagem(result.geometria));if(!confirm('Continuar com corte nas bordas?\n\n'+this.get('excesso-texto').textContent+'\n'+details+'\n\nO PDF será recortado nos limites físicos da folha. Artes, sangria ou marcas externas poderão ser perdidas. Ultrapassar a margem não significa necessariamente sair do papel. A impressora pode ter margens adicionais. Papel e escala não serão alterados.'))return false;if(result!==this.result)return false;this.approved=result;this.geometryState();this.buttons(false);return true;}
    fit(){
        if(this.busy||this.selectionError||this.dirty||!this.result)return;
        try{
            const fit=sugerirAjusteMontagem(this.result.plan,this.result.config||this.config()),pct=n=>n.toLocaleString('pt-BR',{maximumFractionDigits:4});
            if(!confirm(`Ajustar montagem para caber?\nEscala final das artes: ${pct(fit.escalaMin)}%${Math.abs(fit.escalaMax-fit.escalaMin)>.000001?' a '+pct(fit.escalaMax)+'%':''}.\nPapel, proporção, grade, margens, espaços e marcas serão mantidos. A sangria precisa continuar disponível. Aplicar?`))return;
            const real=this.get('modo').value==='nup'&&this.get('escalaModo').value==='real';
            this.ajusteBase??=Object.fromEntries((real?['grade','colunas','linhas','escalaModo','escala']:['grade','colunas','linhas']).map(id=>[id,this.get(id).value]));
            if(real){this.get('escalaModo').value='personalizada';this.get('escala').value=String(fit.fatorMontagem*100);this.fatorMontagem=1;}else this.fatorMontagem=fit.fatorMontagem;
            this.get('grade').value='custom';this.get('colunas').value=fit.config.colunas;this.get('linhas').value=fit.config.linhas;this.controls();this.schedule();
        }catch(e){this.get('status').textContent=e.message;}
    }
    paperZoom(){const g=this.result?.geometria;if(!g)return 1;const box=this.dialog.querySelector('.montagem-previa'),w=Math.max(100,box.clientWidth-20),h=Math.max(100,box.clientHeight-20);return Math.min(8,Math.min(w/g.paper.width,h/g.paper.height)/Math.min(w/g.palco.width,h/g.palco.height));}
    geometryState(){
        const r=this.result,g=r?.geometria;if(!g){this.get('excesso').hidden=true;return;}
        const index=(this.page||1)-1,s=g.sheets[index],fmt=n=>Number(n.toFixed(2)).toLocaleString('pt-BR');
        this.get('aproveitamento').textContent=`Folha ${fmt(g.paper.width)} × ${fmt(g.paper.height)} mm · ${r.plan.folhas[index].length} peças nesta folha · artes ${fmt(s.areaArtesNaFolha/100)} cm² dentro da folha (${fmt(s.aproveitamento)}%) · área sem arte ${fmt(s.areaVazia/100)} cm². Área total das peças ${fmt(s.areaTotalArtes/100)} cm².`;
        this.get('legenda').textContent='Contorno: papel · ocre: margens · azul: peças · hachura vermelha: fora do papel. '+(g.printable?`Roxo: área imprimível da última consulta (${g.printable.printer}); confirme o papel no driver antes de imprimir.`:'Área imprimível: não verificada.');
        const warnings=agruparAvisosMontagem(g);
        this.get('excesso').hidden=!warnings.afetadas;this.get('excesso').dataset.warning=String(g.requerConfirmacao);
        this.get('excesso-texto').textContent=warnings.resumo;
        this.avisos?.atualizar(warnings);
        this.get('aceitar-corte').hidden=r.config?.modo==='nup'||!g.requerConfirmacao||this.approved===r;
        this.get('desfazer-ajuste').hidden=this.fatorMontagem===1&&!this.ajusteBase?.escalaModo;
        if(r.config?.modo==='nup'&&g.requerConfirmacao)this.get('excesso-texto').textContent+=' Saída bloqueada até resolver o conflito.';
        this.get('encaixe').textContent=`Escala final das artes: ${resumirMontagem(r).efetiva}. `+(this.fatorMontagem!==1?'Ajuste proporcional adicional ativo; remover para voltar às medidas configuradas. ':'')+(this.approved===r?'Corte/excesso autorizado somente para esta configuração.':'');
    }
    sourceState(){this.get('origem').textContent=this.prepared?.imagem?`Cópia como imagem RGB · ${this.prepared.imagem.dpi} DPI · aparência comparada. Texto/vetores convertidos em pixels; fidelidade ICC/RIP não garantida. Workspace original intacto.`:this.prepared?'Cópia independente preparada e validada para montagem. O workspace original permanece intacto.':'Origem: cópia do workspace atual; nenhuma preparação aplicada.';this.get('original').hidden=!this.prepared;}
    compatibility(error,allow=true){
        const diag=error.diagnostico;this.diagnostico=diag;
        this.get('compatibilidade').hidden=false;
        this.get('compat-resumo').textContent=diag.podePreparar&&allow?'É possível tentar uma cópia vetorial, com confirmação e comparação de texto e aparência.':'Preparação vetorial indisponível. Conteúdo bloqueado até validar uma cópia independente.';
        this.get('compat-detalhes').textContent=[...diag.removiveis.map(r=>r.descricao),...diag.bloqueios.map(b=>b.descricao)].join('\n');
        this.get('preparar').disabled=!allow||!diag.podePreparar;
        this.get('preparar').hidden=!allow||!diag.podePreparar;
        this.get('tentar').hidden=true;
        this.get('previa-bloqueada').hidden=false;
        if((!allow||!diag.podePreparar)&&!this.prepared?.imagem)this.imagem.mostrar(this.bytes);
    }
    async prepare(imagem=false){
        if(this.preparing||this.saving||!this.bytes)return;
        if(imagem){if(!this.imagem.confirmar())return;}
        else if(!this.diagnostico?.podePreparar||!confirm('Preparar uma cópia independente para montagem?\n\nPoderão ser removidos:\n'+this.diagnostico.removiveis.map(r=>'• '+r.descricao).join('\n')+'\n\nAcessibilidade, navegação e links não imprimíveis podem ser perdidos somente na cópia. ActualText exige validação específica. Texto e aparência de impressão serão comparados antes/depois. Não haverá rasterização. Formulários e aparências não comprovadas continuam bloqueados.\n\nO original e as alterações do workspace serão preservados. Continuar?'))return;
        const dpi=this.imagem.dpi;
        clearTimeout(this.timer);this.abort?.abort();const version=++this.version,abort=this.abort=new AbortController();const bytes=this.bytes.slice();
        this.preparing=true;this.result=null;this.buttons(true);this.imagem.ocupar(true);this.get('preparar').disabled=true;this.get('original').disabled=true;this.get('status').textContent='Preparando cópia e comparando aparência de todas as páginas…';
        const fields=[...this.dialog.querySelectorAll('form input,form select')];for(const field of fields)field.disabled=true;
        this.sequence=this.sequence.catch(()=>{}).then(async()=>{
            let aplicada=false;
            try{
                if(version!==this.version||abort.signal.aborted)return;
                let prepared;
                if(imagem){const {prepararImagemMontagem}=await import('./pdf-imposition-raster.js');prepared=await prepararImagemMontagem(bytes,{dpi,confirmado:true,signal:abort.signal,progresso:text=>{if(version===this.version)this.get('status').textContent=text;}});}
                else{const {prepararCopiaMontagem}=await import('./pdf-imposition-compatibility.js'),{compararCopiaMontagem}=await import('./pdf-imposition-compare.js');prepared=await prepararCopiaMontagem(bytes,{signal:abort.signal,confirmado:true,validar:compararCopiaMontagem});}
                if(version!==this.version||abort.signal.aborted)return;
                this.bytes=prepared.bytes;this.prepared=prepared;this.sourceState();aplicada=true;
            }catch(error){if(version===this.version){this.get('status').textContent='Nenhuma cópia aplicada. '+error.message;this.get('compat-resumo').textContent=error.message;this.get('preparar').disabled=!this.diagnostico?.podePreparar;if(error.diagnostico)this.compatibility(error,false);if(!imagem){this.get('preparar').disabled=true;this.get('preparar').hidden=true;this.imagem.mostrar(bytes);}}}
            finally{this.preparing=false;for(const field of fields)field.disabled=false;this.controls();this.get('original').disabled=false;this.imagem.ocupar(false);if(version===this.version){this.buttons(false);if(aplicada)this.schedule();}}
        });
        return this.sequence;
    }
    schedule(generate=true){
        if(this.preparing||this.saving||!this.dialog.open)return;
        this.imagem.limpar();this.get('previa-bloqueada').hidden=true;
        generate=generate||!this.result||!!this.dirty;
        if(generate)this.avisos?.limpar();
        this.abort?.abort();clearTimeout(this.timer);const version=++this.version;let config;
        if(generate){this.dirty=true;this.approved=null;try{config=this.config();this.selectionError=null;}catch(e){this.selectionError=e;this.buttons(false);this.get('status').textContent=e.message+(this.result?' Prévia anterior mantida; corrija a seleção para gerar ou imprimir.':' Corrija a seleção para visualizar a montagem.');this.get('distribuicao').textContent='Seleção inválida; aguardando correção.';this.get('medidas-resumo').textContent=this.result?resumirMontagem(this.result).basico:'Nenhuma montagem gerada.';this.get('tentar').hidden=true;this.get('compatibilidade').hidden=true;this.get('marcas-externas').hidden=true;this.get('excesso').hidden=true;return;}}
        this.buttons(true);if(generate){this.get('excesso').hidden=true;this.get('aproveitamento').textContent='';this.get('encaixe').textContent='';}
        this.get('status').textContent=this.result?'Atualizando montagem… Prévia anterior até concluir.':'Atualizando montagem…';this.get('distribuicao').textContent='Atualizando distribuição…';this.get('medidas-resumo').textContent='Atualizando medidas e escala efetiva…';this.get('medidas-paginas').hidden=true;this.get('tentar').hidden=true;this.get('marcas-externas').hidden=true;this.get('compatibilidade').hidden=true;
        if(!this.result){const visible=this.dialog.querySelector('canvas');visible.width=visible.height=0;try{desenharFolhaVazia(visible,this.paperConfig());}catch{/* Dimensões inválidas aguardam correção. */}}
        this.timer=setTimeout(()=>{this.sequence=this.sequence.catch(()=>{}).then(async()=>{
            if(version!==this.version||!this.dialog.open)return;
            const abort=this.abort=new AbortController();
            try{
                let result=this.result;const number=generate?1:this.page;
                if(generate){const {gerarMontagem}=await import('./pdf-imposition-engine.js');result=await aguardarMontagem(gerarMontagem(this.bytes,config,{signal:abort.signal,previsualizarExcesso:true}),{signal:abort.signal});if(version!==this.version)return;if(result.geometria)result.geometria=analisarMontagem(result.plan,this.areaImprimivel?.());result.config=config;}
                if(!result)throw Error('Montagem indisponível. Tente novamente.');
                await this.render(abort.signal,result,number);if(version!==this.version)return;
                this.result=result;this.page=number;this.dirty=false;
                const p=this.result.plan;const sizes=[...new Set(p.folhas.flat().map(piece=>`${Number(piece.largura.toFixed(2))} × ${Number(piece.altura.toFixed(2))} mm`))];this.get('status').textContent=`Folha ${p.folha.largura} × ${p.folha.altura} mm · útil ${p.util.largura} × ${p.util.altura} mm · peças: ${sizes.slice(0,3).join('; ')}${sizes.length>3?'…':''} · ${p.colunas} × ${p.linhas} = ${p.capacidade} posições · ${p.quantidade} posições solicitadas · ${p.totalFolhas} folha(s). ${this.result.avisos.join(' ')}`;
                this.get('distribuicao').textContent=`${p.distribuicao==='compacto'?'Distribuição compacta em faixas':this.result.config?.modo==='repeat'&&this.result.config.automatico?'Distribuição automática':'Distribuição em células/grade'}: ${p.colunas} coluna(s) × ${p.linhas} linha(s) · até ${p.capacidade} ${this.result.config?.modo==='nup'?'página(s)':'peça(s)'} por folha · ${p.totalFolhas} folha(s).`;
                const resumo=resumirMontagem(result);this.get('medidas-resumo').textContent=resumo.basico;this.get('medidas-detalhes').textContent=resumo.detalhes;this.get('medidas-paginas').hidden=!resumo.rows?.length;
                this.geometryState();
                if(version!==this.version)return;
                this.buttons(false);
            }catch(e){abort.abort();if(version===this.version){this.result=null;this.buttons(false);this.avisos?.limpar();this.get('excesso').hidden=true;const visible=this.dialog.querySelector('canvas');visible.width=visible.height=0;try{desenharFolhaVazia(visible,this.paperConfig());}catch{}this.get('distribuicao').textContent='Distribuição indisponível. Confira o aviso abaixo.';this.get('medidas-resumo').textContent='Sem resultado validado; geração e impressão bloqueadas.';this.get('medidas-paginas').hidden=true;this.get('status').textContent=e.name==='AbortError'?'Montagem cancelada.':e.message;this.get('tentar').hidden=false;this.get('marcas-externas').hidden=e.code!=='COMPACT_MARKS';if(e.diagnostico)this.compatibility(e);}}
        });},250);
    }
    async render(signal,result=this.result,number=this.page){
        const box=this.dialog.querySelector('.montagem-previa');
        const {canvas,viewport,pages}=await renderizarFolhaMontagem(result,number,{signal,largura:Math.max(100,box.clientWidth-20),altura:Math.max(100,box.clientHeight-20),zoom:this.zoom||1});
        if(signal.aborted||!this.dialog.open){canvas.width=canvas.height=0;return;}
        desenharInspecaoMontagem(canvas,viewport,result,number,{guias:this.get('guias').checked});
        // O aviso/resumo pode mudar a altura disponível após publicar a prévia.
        // Só o enquadramento visual acompanha o espaço; bytes e render não mudam.
        canvas.classList.toggle('montagem-enquadrada',(this.zoom||1)===1);
        const previous=this.dialog.querySelector('canvas');previous.replaceWith(canvas);previous.width=previous.height=0;
        this.get('folha').textContent=`${number} / ${pages}`;
        this.get('zoom').textContent=`Zoom da prévia: ${Math.round((this.zoom||1)*100)}%`;
        if(this.centerPaper&&result.geometria){const {paper,palco}=result.geometria;box.scrollLeft=(paper.width/2-palco.x)/palco.width*canvas.width-box.clientWidth/2;box.scrollTop=(palco.y+palco.height-paper.height/2)/palco.height*canvas.height-box.clientHeight/2;this.centerPaper=false;}else if(this.zoom===1){box.scrollTop=0;box.scrollLeft=0;}
    }
    async save(){if(!this.canOutput())return;const result=this.result;if(!confirm('Conferir montagem antes de gerar o PDF?\n\n'+resumirMontagem(result).confirmacao)||result!==this.result||!this.canOutput())return;this.buttons(true);this.saving=true;const fields=[...this.dialog.querySelectorAll('form input,form select')];for(const field of fields)field.disabled=true;this.get('cancelar').disabled=true;try{await salvarArquivoPDF(result.bytes,'Montagem_4PrintShops.pdf');this.get('status').textContent='PDF de montagem salvo nos limites físicos da folha. Documento original e alterações pendentes preservados.';}catch(e){this.get('status').textContent=e.message;}finally{this.saving=false;for(const field of fields)field.disabled=false;this.get('cancelar').disabled=false;this.controls();if(this.dialog.open)this.buttons(false);}}
    async close(){if(this.saving)return;if(this.closing)return this.closing;clearTimeout(this.timer);++this.version;this.abort?.abort();this.imagem.limpar();this.avisos?.limpar();this.get('excesso').hidden=true;this.dialog.close();this.closing=this.sequence.catch(()=>{}).then(()=>{this.bytes=this.originalBytes=this.prepared=this.result=null;this.dialog.querySelector('canvas').width=0;}).finally(()=>{this.closing=null;});return this.closing;}
    async destroy(){await this.close();this.dialog.remove();}
}
