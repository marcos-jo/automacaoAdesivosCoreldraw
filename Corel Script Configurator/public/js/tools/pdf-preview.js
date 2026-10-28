// Dono das tarefas PDF.js e do canvas principal; não altera bytes do workspace.
export class PDFPreview {
    constructor(container) {
        this.container = container;
        this.documento = null;
        this.atual = null;
        this.tarefas = new Set();
        this.encerramentos = new WeakMap();
        this.render = null;
        this.versao = 0;
        this.encerrada = false;
        this.exibicao = null;
    }

    async preparar(bytes) {
        const pdfjs = await import("/vendor/pdfjs/pdf.js");
        if (this.encerrada) return null;
        pdfjs.GlobalWorkerOptions.workerSrc = "/vendor/pdfjs/pdf.worker.js";
        const tarefa = pdfjs.getDocument({ data: new Uint8Array(bytes) });
        this.tarefas.add(tarefa);
        try {
            const documento = await tarefa.promise;
            if (this.encerrada) {
                await this.descartar({ tarefa });
                return null;
            }
            return { tarefa, documento };
        } catch (erro) {
            await this.descartar({ tarefa });
            if (!this.encerrada) throw erro;
            return null;
        }
    }

    descartar(candidato) {
        const tarefa = candidato?.tarefa;
        if (!tarefa) return Promise.resolve();
        if (!this.encerramentos.has(tarefa)) {
            // Memoiza antes de iniciar: substituição e encerramento podem se encontrar.
            this.encerramentos.set(tarefa, Promise.resolve().then(() => tarefa.destroy())
                .finally(() => this.tarefas.delete(tarefa)));
        }
        return this.encerramentos.get(tarefa);
    }

    invalidar() {
        ++this.versao;
        this.render?.cancel();
    }

    async cancelarRender() {
        this.invalidar();
        const render = this.render;
        if (!render) return;
        // O solicitante já trata falhas de renderização; elas não impedem a limpeza.
        try { await render.promise; } catch {}
        if (this.render === render) this.render = null;
    }

    async adotar(candidato, { preservarCanvas = false } = {}) {
        await this.cancelarRender();
        await this.descartar(this.atual);
        if (this.encerrada) {
            await this.descartar(candidato);
            return false;
        }
        this.atual = candidato;
        this.documento = candidato.documento;
        if (!preservarCanvas) {
            this.exibicao = null;
            for (const canvas of this.container.querySelectorAll("canvas")) canvas.width = canvas.height = 0;
            this.container.replaceChildren();
        }
        return true;
    }

    async previsualizar(documento, numero) {
        const anterior = this.documento;
        if (!await this.atualizarDocumento(documento)) return false;
        try {
            const exibida = await this.mostrar(numero);
            if (!exibida) await this.atualizarDocumento(anterior);
            return exibida;
        } catch (erro) {
            await this.atualizarDocumento(anterior);
            throw erro;
        }
    }

    async atualizarDocumento(documento) {
        await this.cancelarRender();
        if (this.encerrada) return false;
        this.documento = documento;
        return true;
    }

    limparCanvas() {
        this.exibicao = null;
        for (const canvas of this.container.querySelectorAll("canvas")) canvas.width = canvas.height = 0;
        this.container.replaceChildren();
    }

    async mostrar(numero, giro = 0) {
        if (!this.documento || this.encerrada) return false;
        const pdf = this.documento;
        const fonte = pdf.fonte || pdf;
        const chave = pdf.obterChaveMiniatura?.(numero) ?? numero;
        const largura = this.container.clientWidth;
        if (!this.render && this.exibicao?.fonte === fonte && this.exibicao.chave === chave &&
            this.exibicao.giro === giro && this.exibicao.largura === largura &&
            this.container.querySelector("canvas")?.width > 0) {
            // Pode haver um getPage antigo aguardando, ainda sem RenderTask.
            this.invalidar();
            return true;
        }
        const anterior = this.render;
        this.invalidar();
        const versao = this.versao;
        if (anterior) {
            try { await anterior.promise; } catch (erro) {
                if (erro.name !== "RenderingCancelledException") throw erro;
            }
        }
        const vigente = () => !this.encerrada && versao === this.versao && pdf === this.documento;
        if (!vigente()) return false;
        const pagina = await pdf.getPage(numero);
        if (!vigente()) return false;
        const angulo = ((pagina.rotate + giro) % 360 + 360) % 360;
        const base = pagina.getViewport({ scale: 1, rotation: angulo });
        const larguraDisponivel = Math.max(120, (this.container.clientWidth || 898) - 48);
        const escala = Math.min(1.8, larguraDisponivel / base.width, 1100 / base.height);
        const viewport = pagina.getViewport({ scale: escala, rotation: angulo });
        const canvas = document.createElement("canvas");
        canvas.className = "pdf-pagina-canvas";
        canvas.width = Math.ceil(viewport.width);
        canvas.height = Math.ceil(viewport.height);
        const render = pagina.render({ canvasContext: canvas.getContext("2d"), viewport, canvas });
        this.render = render;
        let exibido = false;
        try {
            await render.promise;
            if (!vigente()) return false;
            for (const anterior of this.container.querySelectorAll("canvas")) {
                anterior.width = anterior.height = 0;
            }
            this.container.replaceChildren(canvas);
            this.exibicao = { fonte, chave, giro, largura };
            exibido = true;
            return true;
        } catch (erro) {
            if (erro.name === "RenderingCancelledException" || !vigente()) return false;
            throw erro;
        } finally {
            if (!exibido) canvas.width = canvas.height = 0;
            if (this.render === render) this.render = null;
        }
    }

    destruir() {
        if (!this.encerramento) {
            this.encerrada = true;
            this.encerramento = (async () => {
                await this.cancelarRender();
                await Promise.all([...this.tarefas].map(tarefa => this.descartar({ tarefa })));
                this.documento = this.atual = null;
                this.exibicao = null;
                for (const canvas of this.container.querySelectorAll("canvas")) canvas.width = canvas.height = 0;
                this.container.replaceChildren();
            })();
        }
        return this.encerramento;
    }
}
