// Interrupção por tarefa: getPage não oferece cancelamento próprio e uma promise
// de renderização que não termina não pode reter toda a fila ou a desmontagem.
export class PDFTrabalhoMiniatura {
    constructor(tempoLimite) {
        this.interrupcao = new Promise((_, reject) => { this.rejeitar = reject; });
        this.interrupcao.catch(() => {});
        this.timer = setTimeout(() => this.interromper(Object.assign(
            new Error("A prévia demorou demais. Tente novamente."), { name: "MiniaturaTempoEsgotado" }
        )), tempoLimite);
    }

    esperar(promise) { return Promise.race([promise, this.interrupcao]); }

    interromper(erro) {
        if (this.erro) return;
        this.erro = erro;
        clearTimeout(this.timer);
        try { this.render?.cancel(); } catch {}
        this.rejeitar(erro);
    }

    cancelar() {
        this.interromper(Object.assign(new Error("Prévia cancelada"), { name: "RenderingCancelledException" }));
    }

    concluir() { clearTimeout(this.timer); }
}
