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
            <label for="app-exportacao-nome">Nome do arquivo</label><input id="app-exportacao-nome" type="text" autocomplete="off" spellcheck="false">
            <dl><dt>Formato</dt><dd>PDF</dd><dt>Modalidade</dt><dd id="app-exportacao-modalidade"></dd><dt>Documento</dt><dd id="app-exportacao-resumo"></dd><dt>Dimensões</dt><dd id="app-exportacao-dimensoes"></dd><dt>Alterações</dt><dd id="app-exportacao-estado"></dd></dl>
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
    }
    get aberta() { return !!this.dialogo.open; }
    abrir(dados, exportar, dono) {
        if (this.ocupada || this.aberta) return false;
        this.dados = dados; this.exportar = exportar; this.dono = dono;
        this.campo("nome").value = dados.nome;
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
            await this.exportar(nome, { signal:abort.signal });
            if (!abort.signal.aborted) this.cancelar();
        } catch (erro) {
            if (!abort.signal.aborted) this.campo("status").textContent = erro.message;
        } finally {
            this.ocupada = false; this.atualizar();
            if (!this.aberta) { this.dono = this.exportar = null; }
        }
    }
    atualizar() { this.campo("confirmar").disabled = !!this.ocupada; this.campo("nome").disabled = !!this.ocupada; }
    cancelar() {
        this.abort?.abort();
        if (this.aberta) this.dialogo.close();
        if (!this.ocupada) this.dono = this.exportar = null;
    }
    fecharDono(dono) { if (this.dono === dono) this.cancelar(); }
    destruir() { this.cancelar(); for (const remover of this.listeners) remover(); this.dialogo.remove(); }
}
