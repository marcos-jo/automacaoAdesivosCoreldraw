import { validarCompatibilidadeImpressao } from "./pdf-print-validation.js";

// Ciclo da prévia temporária. Não edita o workspace antes de Aplicar.
export class PDFRedimensionarSessao {
    constructor(acoes) { this.acoes = acoes; this.versao = 0; }
    get ativa() { return !!(this.preparando || this.vista || this.restauracao); }
    async prever(config) {
        await this.cancelar();
        if (!this.acoes.iniciar()) throw new Error("Aguarde a operação atual terminar.");
        const versao = ++this.versao;
        this.abort = new AbortController(); this.preparando = true;
        this.acoes.atualizar();
        this.tarefa = this.executar(config, versao).finally(() => {
            this.preparando = false; this.tarefa = null; this.acoes.liberar();
        });
        return this.tarefa;
    }
    async executar(config, versao) {
        const vigente = () => this.versao === versao && !this.abort.signal.aborted && this.acoes.ativa();
        try {
            this.preparacao = await this.acoes.workspace.prepararRedimensionamento(config, {
                signal: this.abort.signal,
                validarImpressao: dados => validarCompatibilidadeImpressao(dados, {
                    abrir: (bytes, opcoes) => this.acoes.previa.preparar(bytes, opcoes),
                    fechar: candidato => this.acoes.previa.descartar(candidato)
                })
            });
            if (!vigente()) throw new DOMException("Cancelada", "AbortError");
            this.vista = await this.acoes.previa.preparar(await this.preparacao.pdf.save(), { signal: this.abort.signal });
            if (!this.vista || !vigente()) throw new DOMException("Cancelada", "AbortError");
            if (!await this.acoes.previa.previsualizar(this.vista.documento, config.paginaPrevia) || !vigente()) throw new DOMException("Cancelada", "AbortError");
            this.config = config;
            this.acoes.rotulo(`Prévia: ${config.paginaPrevia} / ${this.vista.documento.numPages}`);
            return this.preparacao;
        } catch (erro) {
            await this.restaurar();
            if (!vigente()) throw new DOMException("Cancelada", "AbortError");
            throw erro;
        }
    }
    async restaurar() {
        if (this.restauracao) return this.restauracao;
        if (!this.vista && !this.preparacao) return;
        const vista = this.vista; this.vista = this.preparacao = null;
        this.restauracao = (async () => {
            try {
                await this.acoes.previa.atualizarDocumento(this.acoes.documento());
                if (this.acoes.ativa() && this.acoes.documento()) await this.acoes.mostrarOriginal();
            } catch (erro) {
                this.acoes.previa.limparCanvas(); // Não mostrar uma prévia descartada como documento definitivo.
                throw erro;
            } finally { await this.acoes.previa.descartar(vista); }
        })();
        try { await this.restauracao; } finally { this.restauracao = null; }
    }
    async cancelar() {
        if (this.aplicando) { try { await this.aplicacao; } catch {} }
        ++this.versao;
        if (this.ativa) { this.abort?.abort(); this.acoes.previa.invalidar(); }
        try { await this.tarefa; } catch { /* O solicitante informa a falha. */ }
        try { await this.restaurar(); } finally { this.acoes.atualizar(); }
    }
    async aplicar() {
        if (!this.preparacao || !this.vista || this.preparando || !this.acoes.iniciar()) throw new Error("Gere uma prévia válida antes de aplicar.");
        this.aplicando = true; this.acoes.atualizar();
        try {
            this.aplicacao = this.acoes.aplicar(this.preparacao, this.vista);
            await this.aplicacao;
            this.vista = this.preparacao = null;
        } catch (erro) { await this.restaurar(); throw erro; }
        finally { this.aplicando = false; this.aplicacao = null; this.acoes.liberar(); }
    }
    async reexibir() {
        if (this.vista && !this.preparando && !this.aplicando) await this.acoes.previa.mostrar(this.config.paginaPrevia);
    }
}
