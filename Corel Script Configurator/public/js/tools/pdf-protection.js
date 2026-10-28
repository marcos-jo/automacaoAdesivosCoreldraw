// Confirmação e avisos do navegador; o workspace é a única fonte de pendências.
export class PDFProtecao {
    constructor(raiz, workspace, janela = window) {
        this.workspace = workspace;
        this.janela = janela;
        this.dialogo = raiz.querySelector("#pdf-confirmar-descarte");
        this.indicador = raiz.querySelector("#pdf-alteracoes-pendentes");
        this.continuar = raiz.querySelector("#pdf-continuar-editando");
        this.descartar = raiz.querySelector("#pdf-descartar-sair");
        this.antesDeFechar = evento => {
            if (this.encerrada || !this.workspace.alterado) return;
            evento.preventDefault();
            evento.returnValue = "";
        };
        this.cancelar = evento => { evento.preventDefault(); this.finalizar(false); };
        this.permanecer = () => this.finalizar(false);
        // O evento close é enfileirado: pode chegar depois de uma nova abertura.
        this.aoFechar = () => { if (!this.dialogo.open) this.finalizar(false); };
        this.sair = () => this.finalizar(true);
        this.continuar.addEventListener("click", this.permanecer);
        this.descartar.addEventListener("click", this.sair);
        this.dialogo.addEventListener("cancel", this.cancelar);
        this.dialogo.addEventListener("close", this.aoFechar);
        this.atualizar();
    }

    atualizar() {
        if (this.encerrada) return;
        const pendente = this.workspace.alterado;
        this.indicador.hidden = !pendente;
        if (pendente && !this.observando) this.janela.addEventListener("beforeunload", this.antesDeFechar);
        if (!pendente && this.observando) this.janela.removeEventListener("beforeunload", this.antesDeFechar);
        this.observando = pendente;
    }

    confirmarDescarte() {
        if (this.encerrada) return Promise.resolve(false);
        if (this.confirmacao) return this.confirmacao;
        if (!this.workspace.alterado) return Promise.resolve(true);
        this.confirmacao = new Promise(resolve => { this.resolver = resolve; });
        const resposta = this.confirmacao;
        try {
            this.dialogo.showModal();
            this.continuar.focus();
        } catch {
            // Se a janela não puder abrir, preservar o documento.
            this.finalizar(false);
        }
        return resposta;
    }

    finalizar(descartar) {
        if (!this.resolver) return;
        const resolver = this.resolver;
        this.resolver = this.confirmacao = null;
        this.dialogo.close();
        resolver(descartar);
    }

    destruir() {
        if (this.encerrada) return;
        this.encerrada = true;
        this.finalizar(false);
        this.janela.removeEventListener("beforeunload", this.antesDeFechar);
        this.observando = false;
        this.continuar.removeEventListener("click", this.permanecer);
        this.descartar.removeEventListener("click", this.sair);
        this.dialogo.removeEventListener("cancel", this.cancelar);
        this.dialogo.removeEventListener("close", this.aoFechar);
    }
}
