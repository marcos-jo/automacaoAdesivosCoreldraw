export class PDFImpressaoUI {
    constructor(buscar, sessao, interpretar, confirmar = texto => globalThis.confirm(texto)) {
        Object.assign(this, { sessao, interpretar, confirmar });
        this.campo = id => buscar("#pdf-impressao-" + id); this.listeners = []; this.geracao = 0;
        const evento = (elemento, tipo, acao) => { elemento.addEventListener(tipo, acao); this.listeners.push(() => elemento.removeEventListener(tipo, acao)); };
        evento(buscar("#pdf-assinaturas-resolver"), "click", () => this.tentar(async () => {
            const resumo = sessao.abrir();
            this.campo("modo").value = "necessarias"; this.campo("dpi").value = "300"; this.campo("manual").value = "";
            this.campo("conferi").checked = false; this.campo("origem").checked = false;
            this.campo("detalhes").open = false;
            this.campo("titulo").textContent = resumo.preparada ? "Refazer preparo para impressão" : "Preparar cópia para impressão";
            const paginas = resumo.paginasVisuais.length;
            this.campo("contexto").textContent = resumo.preparada
                ? "Cópia atual sem assinatura ativa. Para nova rasterização, informe páginas ou escolha documento inteiro."
                : `Páginas visuais: ${paginas ? `${paginas} identificada(s)` : "nenhuma determinada"}.${resumo.invisiveisPendentes ? " Há assinatura invisível; informe páginas ou escolha documento inteiro." : !paginas ? " Informe páginas ou escolha documento inteiro." : ""}`;
            this.campo("diagnostico").textContent = `Páginas visuais conhecidas na posição atual: ${resumo.paginasVisuais.join(", ") || "nenhuma determinada"}.\n${resumo.preparada ? "A versão atual já é uma cópia de impressão: rasterizar novamente em maior DPI não recupera detalhes perdidos; reimporte o original se necessário.\n" : ""}${resumo.validade}`;
            this.campo("resultado-tecnico").textContent = "";
            this.mensagem("");
            try { this.campo("dialogo").showModal(); } catch (erro) { await sessao.cancelar(); throw erro; }
            this.atualizar();
        }));
        for (const id of ["modo", "dpi", "manual"]) evento(this.campo(id), id === "manual" ? "input" : "change", () => this.tentar(async () => {
            ++this.geracao; this.campo("conferi").checked = false; this.campo("resultado-tecnico").textContent = ""; await sessao.descartarPrevia(); this.mensagem(""); this.atualizar();
        }));
        evento(this.campo("prever"), "click", () => this.tentar(async () => {
            const geracao = ++this.geracao; this.campo("conferi").checked = false;
            this.campo("resultado-tecnico").textContent = "";
            const manual = this.campo("manual").value.trim(), total = sessao.acoes.workspace.getTotalPaginas();
            const resultado = await sessao.prever({ modo: this.campo("modo").value, dpi: Number(this.campo("dpi").value), paginas: manual ? interpretar(manual, total) : [] });
            if (geracao !== this.geracao || this.encerrada) return;
            this.campo("resultado-tecnico").textContent = `Proposta preparada e comparada com a renderização PDF.js. Rasterizadas: ${resultado.plano.rasterizadas.join(", ")}. Sem nova rasterização: ${resultado.plano.vetoriais.length}. ${resultado.plano.avisos.join(" ")} Não é garantia de renderização correta do original.`;
            this.mensagem(`Prévia pronta · ${resultado.plano.rasterizadas.length} página(s) rasterizada(s). Confira antes de aplicar.`);
            this.atualizar();
        }));
        evento(this.campo("ir"), "click", () => this.tentar(() => this.rever(Number(this.campo("pagina").value))));
        for (const [id, passo] of [["anterior", -1], ["proxima", 1]]) evento(this.campo(id), "click", () => this.tentar(async () => {
            const lista = sessao.preparacao?.plano.rasterizadas || [], i = lista.indexOf(sessao.pagina);
            await this.rever(lista[Math.max(0, Math.min(lista.length - 1, i + passo))]);
        }));
        evento(this.campo("origem"), "change", () => this.tentar(() => this.rever(sessao.pagina)));
        evento(this.campo("conferi"), "change", () => this.atualizar());
        evento(this.campo("aplicar"), "click", () => this.tentar(async () => {
            if (!this.campo("conferi").checked || !sessao.pronta) return;
            if (!confirmar("Será criada uma cópia para impressão. A nova versão não preservará a validade da assinatura digital. O documento original permanecerá intacto.")) return;
            await sessao.aplicar(); this.campo("dialogo").close();
        }));
        evento(this.campo("cancelar"), "click", () => this.tentar(() => this.fechar()));
        evento(this.campo("dialogo"), "cancel", e => { e.preventDefault(); this.tentar(() => this.fechar()); });
        evento(this.campo("dialogo"), "close", () => { if (!this.campo("dialogo").open && sessao.aberta) this.tentar(() => this.fechar()); });
        this.atualizar();
    }
    mensagem(texto) { if (!this.encerrada) this.campo("progresso").textContent = texto; }
    async tentar(acao) { try { await acao(); } catch (erro) { if (erro.name !== "AbortError") this.mensagem(erro.message); } finally { if (!this.encerrada) this.atualizar(); } }
    async rever(numero) { await this.sessao.rever(numero, !!this.campo("origem").checked); this.atualizar(); }
    async fechar() { if (this.sessao.aplicando) return; ++this.geracao; await this.sessao.cancelar(); this.campo("dialogo").close(); this.atualizar(); }
    atualizar() {
        const s = this.sessao, ocupada = !!s.preparando || !!s.aplicando, lista = s.preparacao?.plano.rasterizadas || [], pronta = s.pronta;
        for (const id of ["controles-revisao", "confirmacao", "aplicar"]) this.campo(id).hidden = !pronta;
        for (const id of ["modo", "dpi", "manual", "prever"]) this.campo(id).disabled = ocupada;
        for (const id of ["pagina", "ir", "origem", "conferi"]) this.campo(id).disabled = !pronta || ocupada;
        this.campo("cancelar").disabled = !!s.aplicando;
        this.campo("aplicar").disabled = !pronta || ocupada || !this.campo("conferi").checked || !lista.every(n => s.revisadas.has(n));
        this.campo("anterior").disabled = !pronta || ocupada || lista.indexOf(s.pagina) <= 0;
        this.campo("proxima").disabled = !pronta || ocupada || lista.indexOf(s.pagina) >= lista.length - 1;
        this.campo("pagina").max = s.acoes.workspace.getTotalPaginas();
        if (s.pagina) this.campo("pagina").value = String(s.pagina);
        this.campo("revisadas").textContent = pronta ? `${s.mostrandoOrigem ? "Antes do preparo" : "Cópia proposta"}: página ${s.pagina}. Necessárias exibidas: ${lista.filter(n => s.revisadas.has(n)).length} / ${lista.length}.` : "";
    }
    destruir() { this.encerrada = true; for (const remover of this.listeners) remover(); }
}
