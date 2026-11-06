// Controles pertencem a cada modal. Respostas antigas não reabrem uma oferta.
export class AlternativaImagemMontagem {
    constructor(get,preparar){
        this.get=get;this.revisao=0;
        const box=document.createElement('div');box.dataset.id='imagem-opcoes';box.hidden=true;
        box.innerHTML='<div class="montagem-imagem-acoes"><label>Resolução<select data-id="imagem-dpi"><option value="300">300 DPI</option><option value="600">600 DPI</option></select></label><button type="button" data-id="preparar-imagem">Preparar cópia como imagem</button></div><details><summary>Perdas e limites da cópia como imagem</summary><p data-id="imagem-aviso"></p></details><p data-id="imagem-status" role="status"></p>';
        get('compatibilidade').insertBefore(box,get('compat-detalhes').closest('details'));this.box=box;
        get('imagem-dpi').onchange=()=>this.atualizar();get('preparar-imagem').onclick=()=>{if(this.disponivel)preparar();};
    }
    get disponivel(){return !this.box.hidden&&!this.get('preparar-imagem').disabled&&this.opcoes?.some(o=>o.dpi===this.dpi&&o.disponivel);}
    get dpi(){return Number(this.get('imagem-dpi').value);}
    limpar(){++this.revisao;this.box.hidden=true;this.opcoes=null;this.get('preparar-imagem').disabled=true;}
    async mostrar(bytes){
        this.limpar();const revisao=this.revisao;this.box.hidden=false;this.get('imagem-status').textContent='Verificando recursos e limites de memória…';
        try{const {avaliarImagemMontagem,avisoImagemMontagem}=await import('./pdf-imposition-raster.js');const opcoes=await avaliarImagemMontagem(bytes);if(revisao!==this.revisao)return;this.aviso=avisoImagemMontagem;this.get('imagem-aviso').textContent=this.aviso;this.opcoes=opcoes;for(const o of opcoes)this.get('imagem-dpi').querySelector(`option[value="${o.dpi}"]`).disabled=!o.disponivel;if(!opcoes.some(o=>o.dpi===this.dpi&&o.disponivel))this.get('imagem-dpi').value=String(opcoes.find(o=>o.disponivel)?.dpi||300);this.atualizar();}
        catch(e){if(revisao===this.revisao){this.get('imagem-status').textContent='Cópia como imagem indisponível: '+e.message;this.get('preparar-imagem').disabled=true;}}
    }
    atualizar(){
        const o=this.opcoes?.find(o=>o.dpi===this.dpi);this.get('preparar-imagem').disabled=!o?.disponivel;
        this.get('imagem-status').textContent=o?.disponivel?`${o.paginas.length} página(s), ${o.dpi} DPI. Estimativa RGB: ${Math.ceil(o.totalRGB/1048576)} MiB; haverá verificação de aparência antes de aplicar à janela.`:'Cópia como imagem indisponível: '+(o?.motivo||'Resolução não validada.');
    }
    ocupar(value){this.get('imagem-dpi').disabled=value;if(value)this.get('preparar-imagem').disabled=true;else this.atualizar();}
    confirmar(){return this.disponivel&&confirm(`Preparar cópia como imagem em ${this.dpi} DPI?\n\n${this.aviso}\n\nApenas a cópia desta janela será usada na montagem/livreto. O original e suas alterações pendentes serão preservados. Confira visualmente o resultado antes de gerar ou imprimir. Continuar?`);}
}
