import { imagePages } from "../tools/pdf-image-export.js";
export function nomeExportacao(valor, impressao = false) {
    const nome = String(valor || "").trim().replace(/\.pdf$/i, "");
    if (!nome || /[\\/:*?"<>|\x00-\x1f]/.test(nome) || /^\.+$/.test(nome) || /[. ]$/.test(nome)) throw new Error("Informe um nome de arquivo válido, sem barras ou caracteres especiais de caminho.");
    return nome + (impressao && !/_IMPRESSAO$/i.test(nome) ? "_IMPRESSAO" : "") + ".pdf";
}

// Janela compartilhada; cada módulo fornece os dados e seu exportador real.
export class JanelaExportacao {
    constructor(container) {
        this.dialogo = document.createElement("dialog"); this.dialogo.id = "app-exportacao-dialogo";
        this.dialogo.className = "app-exportacao-dialogo"; this.dialogo.setAttribute("aria-labelledby","app-exportacao-titulo");
        this.dialogo.innerHTML = `<h2 id="app-exportacao-titulo">Exportar cópia</h2>
            <label id="app-exportacao-nome-label" for="app-exportacao-nome">Nome do arquivo</label><input id="app-exportacao-nome" type="text" autocomplete="off" spellcheck="false">
            <label for="app-exportacao-formato">Formato</label><select id="app-exportacao-formato"><option value="pdf">PDF</option><option value="png">PNG</option><option value="jpeg">JPEG</option></select>
            <fieldset id="app-exportacao-imagens" hidden><legend>Imagens por página</legend>
            <label>Páginas<select id="app-exportacao-alvo"><option value="all">Todas</option><option value="current">Página atual</option><option value="selected">Selecionadas nas miniaturas</option><option value="range">Intervalo</option></select></label>
            <label>Intervalo<input id="app-exportacao-intervalo" placeholder="1,3-5"></label>
            <label>Resolução<select id="app-exportacao-dpi"><option value="150">150 DPI</option><option value="300" selected>300 DPI</option><option value="600">600 DPI</option></select></label>
            <label>Qualidade JPEG (%)<input id="app-exportacao-qualidade" type="number" min="10" max="100" value="92"></label>
            <p>Será criada uma subpasta nova e identificável no local escolhido, com imagens numeradas. Arquivos existentes são preservados. Saída RGB sobre branco; ICC/CMYK e equivalência com o RIP não são garantidos. PNG sem perdas e JPEG com perdas. Exportar imagens mantém as alterações pendentes do PDF.</p></fieldset>
            <dl id="app-exportacao-metadados"><dt>Formato</dt><dd id="app-exportacao-formato-resumo">PDF</dd><dt>Modalidade</dt><dd id="app-exportacao-modalidade"></dd><dt>Documento</dt><dd id="app-exportacao-resumo"></dd><dt>Dimensões</dt><dd id="app-exportacao-dimensoes"></dd><dt>Alterações</dt><dd id="app-exportacao-estado"></dd></dl>
            <p id="app-exportacao-avisos" class="app-exportacao-avisos"></p>
            <p id="app-exportacao-status" role="status"></p>
            <div class="app-exportacao-acoes"><button id="app-exportacao-cancelar" type="button">Cancelar</button><button id="app-exportacao-confirmar" type="button">Exportar</button></div>
            <small>O arquivo original será preservado.</small>`;
        container.append(this.dialogo); this.listeners = [];
        this.campo = id => this.dialogo.querySelector("#app-exportacao-" + id);
        const evento = (alvo, tipo, acao) => { alvo.addEventListener(tipo,acao); this.listeners.push(() => alvo.removeEventListener(tipo,acao)); };
        evento(this.campo("cancelar"),"click",() => this.cancelar());
        evento(this.dialogo,"cancel",e => { e.preventDefault(); this.cancelar(); });
        evento(this.dialogo,"close",() => { if (!this.aberta) this.cancelar(); });
        evento(this.campo("confirmar"),"click",() => this.confirmar());
        evento(this.campo("formato"),"change",() => {
            this.campo("nome").value=this.campo("nome").value.replace(/\.(pdf|png|jpe?g)$/i, "");
            this.atualizar();
        });
        evento(this.campo("alvo"),"change",() => this.atualizar());
    }
    get aberta() { return !!this.dialogo.open; }
    abrir(dados, exportar, dono) {
        if (this.ocupada || this.aberta) return false;
        this.dados = dados; this.exportar = exportar; this.dono = dono;
        this.campo("nome").value = dados.nome;
        this.campo("formato").value = "pdf"; this.campo("alvo").value = "all"; this.campo("dpi").value = "300";
        for (const id of ["modalidade","resumo","dimensoes","estado","avisos"]) this.campo(id).textContent = dados[id] || "";
        this.campo("status").textContent = ""; this.atualizar(); this.dialogo.showModal(); this.campo("nome").focus(); return true;
    }
    async confirmar() {
        if (this.ocupada || !this.aberta) return;
        let nome;
        try { nome = nomeExportacao(this.campo("nome").value, this.dados.impressao); }
        catch (erro) { this.campo("status").textContent = erro.message; this.campo("nome").focus(); return; }
        const abort = this.abort = new AbortController(); this.ocupada = true; this.atualizar(); this.campo("status").textContent = "Exportando cópia…";
        try {
            const format=this.campo('formato').value;
            if(format==='pdf')await this.exportar(nome, {signal:abort.signal});
            else {
                if(!this.dados.imagens)throw Error('Exportação de imagens indisponível neste documento.');
                const pages=imagePages(this.campo('alvo').value,this.campo('intervalo').value,this.dados.imagens.total,this.dados.imagens.atual,this.dados.imagens.selecionadas);
                const result=await this.dados.imagens.exportar({format,pages,dpi:Number(this.campo('dpi').value),quality:Number(this.campo('qualidade').value)/100,base:nome.replace(/\.pdf$/i,'')}, {signal:abort.signal,progress:text=>{this.campo('status').textContent=text;}});
                if(!abort.signal.aborted)this.campo('status').textContent=`${result.count} imagem(ns) exportada(s) em ${result.location}`;
            }
            if (!abort.signal.aborted) this.cancelar();
        } catch (erro) {
            if (!abort.signal.aborted) this.campo("status").textContent = erro.message;
        } finally {
            this.ocupada = false; this.atualizar();
            if (!this.aberta) { this.dono = this.exportar = null; }
        }
    }
    atualizar() {
        for(const id of ['confirmar','nome','formato','alvo','intervalo','dpi','qualidade'])this.campo(id).disabled=!!this.ocupada;
        const format=this.campo('formato').value||'pdf';
        this.campo('imagens').hidden=format==='pdf';
        this.campo('metadados').hidden=format!=='pdf';
        this.campo('nome-label').textContent=format==='pdf'?'Nome do arquivo':'Nome base das imagens';
        this.campo('qualidade').disabled=!!this.ocupada||format!=='jpeg';
        this.campo('intervalo').disabled=!!this.ocupada||this.campo('alvo').value!=='range';
        this.campo('formato-resumo').textContent=format.toUpperCase();
    }
    cancelar() {
        this.abort?.abort();
        if (this.aberta) this.dialogo.close();
        if (!this.ocupada) this.dono = this.exportar = null;
    }
    fecharDono(dono) { if (this.dono === dono) this.cancelar(); }
    destruir() { this.cancelar(); for (const remover of this.listeners) remover(); this.dialogo.remove(); }
}
