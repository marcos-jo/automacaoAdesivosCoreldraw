import { rasterizarPaginaImpressao, conferirCopiaImpressao, documentoDeImpressao } from "./pdf-print-raster.js";

// Todas as tarefas de carregamento pertencem ao PDFPreview principal. O preview
// do modal apenas renderiza proxies e nunca adota/destrói tarefas de outro dono.
export class PDFImpressaoSessao {
    constructor(acoes) { this.acoes = acoes; this.geracao = 0; this.revisadas = new Set(); }
    get ativa() { return !!this.aberta; }
    get pronta() { return !!this.preparacao?.verificacao?.compativel && !!this.candidato && !this.preparando; }
    abrir() {
        if (!this.acoes.ativa() || !this.acoes.podeAbrir()) throw new Error("Conclua a operação ou prévia atual antes de Resolver.");
        this.aberta = true; this.acoes.atualizar();
        return this.acoes.workspace.obterResumoAssinaturas();
    }
    async prever(config) {
        await this.descartarPrevia();
        if (!this.aberta || !this.acoes.iniciar()) throw new Error("Aguarde a operação atual terminar.");
        const geracao = ++this.geracao;
        this.abort = new AbortController(); this.preparando = true; this.acoes.atualizar();
        this.tarefa = this.executar(config, geracao).finally(() => { this.preparando = false; this.tarefa = null; this.acoes.liberar(); });
        return this.tarefa;
    }
    async executar(config, geracao) {
        const signal = this.abort.signal, { previa, workspace } = this.acoes;
        const vigente = () => !signal.aborted && geracao === this.geracao && this.aberta && this.acoes.ativa();
        try {
            this.preparacao = await workspace.prepararCopiaImpressao(config, {
                signal, progresso: texto => this.acoes.progresso(texto),
                carregarFonte: async bytes => { this.fonte = await previa.preparar(bytes, { signal }); },
                rasterizar: (n, plano) => rasterizarPaginaImpressao(this.fonte.documento, n, plano, { signal, progresso: this.acoes.progresso })
            });
            if (!vigente()) throw new DOMException("Cancelada", "AbortError");
            const bytes = await this.preparacao.pdf.save({ updateFieldAppearances: false });
            this.candidato = await previa.preparar(bytes, { signal });
            if (!this.candidato || !vigente()) throw new DOMException("Cancelada", "AbortError");
            this.preparacao.verificacao = await conferirCopiaImpressao(this.fonte.documento, this.candidato.documento, this.preparacao, { signal, progresso: this.acoes.progresso });
            if (!vigente()) throw new DOMException("Cancelada", "AbortError");
            this.preparacao.imagens.clear();
            await this.rever(this.preparacao.plano.rasterizadas[0]);
            return this.preparacao;
        } catch (erro) { await this.limpar(); throw erro; }
    }
    async rever(numero, origem = false) {
        if (!this.candidato || !this.preparacao || !Number.isSafeInteger(numero) || numero < 1 || numero > this.candidato.documento.numPages) throw new Error("Página de conferência inválida.");
        const geracao = this.geracao;
        const pedido = this.pedidoVisual = (this.pedidoVisual || 0) + 1;
        const documento = origem ? this.fonte.documento : this.candidato.documento;
        await this.acoes.modalPrevia.atualizarDocumento(documentoDeImpressao(documento));
        if (geracao !== this.geracao || pedido !== this.pedidoVisual) return;
        if (await this.acoes.modalPrevia.mostrar(numero) && geracao === this.geracao && pedido === this.pedidoVisual) {
            this.pagina = numero; this.mostrandoOrigem = origem;
            if (!origem) this.revisadas.add(numero);
            this.acoes.atualizar();
        }
    }
    async aplicar() {
        if (!this.pronta || !this.preparacao.plano.rasterizadas.every(n => this.revisadas.has(n)) || !this.acoes.iniciar()) throw new Error("Gere e confira todas as páginas rasterizadas antes de aplicar.");
        this.aplicando = true; this.acoes.atualizar();
        try {
            await this.acoes.modalPrevia.atualizarDocumento(null); this.acoes.modalPrevia.limparCanvas();
            await this.acoes.aplicar(this.preparacao, this.candidato);
            this.candidato = null; this.preparacao = null; this.aberta = false;
            await this.limpar();
        } catch (erro) { await this.limpar(); throw erro; }
        finally { this.aplicando = false; this.acoes.liberar(); }
    }
    async limpar() {
        await this.acoes.modalPrevia.atualizarDocumento(null); this.acoes.modalPrevia.limparCanvas();
        const fonte = this.fonte, candidato = this.candidato;
        this.fonte = this.candidato = this.preparacao = null; this.revisadas.clear();
        await Promise.all([this.acoes.previa.descartar(fonte), this.acoes.previa.descartar(candidato)]);
    }
    async descartarPrevia() {
        ++this.geracao; this.abort?.abort(); this.acoes.modalPrevia.invalidar();
        try { await this.tarefa; } catch {}
        await this.limpar(); this.acoes.atualizar();
    }
    async cancelar() {
        if (this.aplicando) return;
        this.aberta = false;
        await this.descartarPrevia();
    }
}
