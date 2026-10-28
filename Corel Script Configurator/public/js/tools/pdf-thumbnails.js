const ALTURA_ITEM = 166;
const MARGEM = 2;

// Janela virtual: somente páginas próximas da área visível têm botões/canvases.
export class PDFMiniaturas {
    constructor(container, aoSelecionar) {
        this.container = container;
        this.aoSelecionar = aoSelecionar;
        this.itens = new Map();
        this.geracao = 0;
        this.selecionada = 1;
        this.rolar = () => this.atualizarJanela();
        this.clicar = evento => {
            const botao = evento.target.closest("button[data-pagina]");
            if (botao && !this.ocupado) this.aoSelecionar(Number(botao.dataset.pagina));
        };
        container.addEventListener("scroll", this.rolar, { passive: true });
        container.addEventListener("click", this.clicar);
        if (typeof ResizeObserver !== "undefined") {
            this.observador = new ResizeObserver(this.rolar);
            this.observador.observe(container);
        }
    }

    definirDocumento(pdf, selecionada = 1) {
        if (this.encerrada) return;
        this.pdf = pdf;
        this.selecionada = selecionada;
        this.container.scrollTop = (selecionada - 1) * ALTURA_ITEM;
        this.lista = document.createElement("div");
        this.lista.className = "pdf-miniaturas-lista";
        this.lista.style.height = `${pdf.numPages * ALTURA_ITEM}px`;
        this.container.replaceChildren(this.lista);
        this.atualizarJanela();
    }

    definirOcupado(valor) {
        this.ocupado = valor;
        for (const item of this.itens.values()) item.botao.disabled = valor;
    }

    selecionar(numero) {
        this.selecionada = numero;
        const topo = (numero - 1) * ALTURA_ITEM;
        const altura = this.container.clientHeight || 600;
        if (topo < this.container.scrollTop || topo + ALTURA_ITEM > this.container.scrollTop + altura) {
            this.container.scrollTop = topo;
        }
        this.atualizarJanela();
    }

    removerItem(numero, item) {
        if (this.atual === item) this.tarefa?.cancel();
        item.botao.remove();
        // Um canvas em renderização é liberado apenas depois da tarefa terminar.
        if (this.atual !== item && item.canvas) item.canvas.width = item.canvas.height = 0;
        this.itens.delete(numero);
    }

    atualizarJanela() {
        if (!this.pdf || this.encerrada) return;
        const total = this.pdf.numPages;
        const visivel = Math.floor((this.container.scrollTop || 0) / ALTURA_ITEM);
        const quantidade = Math.min(16, Math.ceil((this.container.clientHeight || 600) / ALTURA_ITEM) + 2 * MARGEM);
        const inicio = Math.max(1, Math.min(total - quantidade + 1, visivel + 1 - MARGEM));
        const fim = Math.min(total, inicio + quantidade - 1);
        for (const [numero, item] of this.itens) {
            if (numero < inicio || numero > fim) this.removerItem(numero, item);
        }
        for (let numero = inicio; numero <= fim; numero++) {
            if (!this.itens.has(numero)) {
                const botao = document.createElement("button");
                botao.type = "button";
                botao.className = "pdf-miniatura";
                botao.dataset.pagina = String(numero);
                botao.setAttribute("aria-label", `Abrir página ${numero}`);
                botao.style.top = `${(numero - 1) * ALTURA_ITEM}px`;
                const imagem = document.createElement("span");
                imagem.className = "pdf-miniatura-imagem";
                imagem.textContent = "Carregando…";
                const legenda = document.createElement("span");
                legenda.textContent = `Página ${numero}`;
                botao.append(imagem, legenda);
                this.lista.appendChild(botao);
                this.itens.set(numero, { botao, imagem, estado: "pendente" });
            }
            const item = this.itens.get(numero);
            item.botao.disabled = !!this.ocupado;
            item.botao.setAttribute("aria-current", String(numero === this.selecionada));
        }
        if (!this.processamento) {
            // Inicia na próxima microtarefa para a promessa existir antes do trabalho.
            this.processamento = Promise.resolve().then(() => this.processarFila())
                .finally(() => { this.processamento = null; });
        }
    }

    async processarFila() {
        const geracao = this.geracao;
        const pdf = this.pdf;
        while (pdf && this.pdf === pdf && geracao === this.geracao && !this.encerrada) {
            const proximo = [...this.itens].find(([, item]) => item.estado === "pendente");
            if (!proximo) return;
            const [numero, item] = proximo;
            const vigente = () => this.pdf === pdf && geracao === this.geracao && this.itens.get(numero) === item;
            item.estado = "renderizando";
            this.atual = item;
            try {
                const pagina = await pdf.getPage(numero);
                if (!vigente()) continue;
                const base = pagina.getViewport({ scale: 1 });
                const viewport = pagina.getViewport({ scale: Math.min(112 / base.width, 126 / base.height, 1) });
                const canvas = document.createElement("canvas");
                canvas.className = "pdf-miniatura-canvas";
                item.canvas = canvas;
                canvas.width = Math.ceil(viewport.width);
                canvas.height = Math.ceil(viewport.height);
                canvas.setAttribute("aria-hidden", "true");
                this.tarefa = pagina.render({ canvasContext: canvas.getContext("2d"), viewport, canvas });
                await this.tarefa.promise;
                if (vigente()) {
                    item.imagem.replaceChildren(canvas);
                    item.estado = "pronto";
                }
            } catch (erro) {
                if (vigente()) {
                    item.estado = "erro";
                    item.imagem.textContent = "Prévia indisponível";
                    item.botao.title = erro.name === "RenderingCancelledException" ? "Prévia cancelada" : "Clique para abrir a página";
                }
            } finally {
                if (item.estado !== "pronto" && item.canvas) item.canvas.width = item.canvas.height = 0;
                this.tarefa = this.atual = null;
            }
        }
    }

    async limpar() {
        ++this.geracao;
        this.pdf = null;
        this.tarefa?.cancel();
        await this.processamento;
        for (const [numero, item] of this.itens) this.removerItem(numero, item);
        this.container.replaceChildren();
        this.lista = null;
    }

    destruir() {
        if (!this.encerramento) {
            this.encerrada = true;
            this.container.removeEventListener("scroll", this.rolar);
            this.container.removeEventListener("click", this.clicar);
            this.observador?.disconnect();
            this.encerramento = this.limpar();
        }
        return this.encerramento;
    }
}
