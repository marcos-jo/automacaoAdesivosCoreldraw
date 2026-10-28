// Interação de arraste apenas. O controlador aplica o movimento no workspace.
export class PDFArrasteMiniaturas {
    constructor(container, { alturaItem, obterLista, obterTotal, bloqueado, aoMover, aoRolar }) {
        Object.assign(this, { container, alturaItem, obterLista, obterTotal, bloqueado, aoMover, aoRolar });
        this.handlers = {
            dragstart: evento => this.iniciar(evento),
            dragover: evento => this.sobre(evento),
            drop: evento => this.soltar(evento),
            dragend: () => this.cancelar(),
            dragleave: evento => {
                const limites = container.getBoundingClientRect();
                const dentro = Number.isFinite(evento.clientY) && evento.clientY >= limites.top && evento.clientY <= limites.bottom &&
                    (!Number.isFinite(limites.left) || evento.clientX >= limites.left && evento.clientX <= limites.right);
                if (!container.contains(evento.relatedTarget) && !dentro) {
                    this.pararRolagem();
                    if (this.marcador) this.marcador.hidden = true;
                }
            }
        };
        for (const [nome, handler] of Object.entries(this.handlers)) container.addEventListener(nome, handler);
    }

    iniciar(evento) {
        const botao = evento.target.closest("button[data-pagina]");
        const numero = Number(botao?.dataset.pagina);
        if (this.bloqueado() || !evento.dataTransfer || !Number.isInteger(numero) ||
            numero < 1 || numero > this.obterTotal() || this.obterTotal() < 2) {
            evento.preventDefault();
            return;
        }
        this.cancelar();
        this.origem = numero;
        this.botaoOrigem = botao;
        botao.classList.toggle("pdf-miniatura-arrastando", true);
        evento.dataTransfer.setData("application/x-4printshops-pagina", String(numero));
        evento.dataTransfer.effectAllowed = "move";
    }

    marcar(clientY) {
        const lista = this.obterLista();
        if (!lista) return;
        const y = clientY - this.container.getBoundingClientRect().top + this.container.scrollTop;
        this.posicao = Math.max(0, Math.min(this.obterTotal(), Math.round(y / this.alturaItem)));
        if (!this.marcador) {
            this.marcador = document.createElement("span");
            this.marcador.className = "pdf-miniatura-insercao";
            this.marcador.setAttribute("aria-hidden", "true");
            lista.appendChild(this.marcador);
        }
        this.marcador.hidden = false;
        this.marcador.style.top = `${Math.min(this.posicao * this.alturaItem, this.obterTotal() * this.alturaItem - 4)}px`;
    }

    sobre(evento) {
        if (!this.origem || this.bloqueado() || !evento.dataTransfer) return;
        evento.preventDefault();
        evento.dataTransfer.dropEffect = "move";
        this.ponteiroY = evento.clientY;
        this.marcar(this.ponteiroY);
        if (!this.frame) this.frame = requestAnimationFrame(() => this.rolar());
    }

    rolar() {
        this.frame = null;
        if (!this.origem || this.bloqueado()) return;
        const limites = this.container.getBoundingClientRect();
        const topo = Math.max(0, limites.top);
        const fundo = Math.min(limites.bottom, globalThis.innerHeight || limites.bottom);
        const distancia = this.ponteiroY < topo + 44
            ? this.ponteiroY - topo - 44
            : this.ponteiroY > fundo - 44 ? this.ponteiroY - fundo + 44 : 0;
        if (!distancia) return;
        const anterior = this.container.scrollTop;
        const maximo = Math.max(0, this.container.scrollHeight - this.container.clientHeight);
        this.container.scrollTop = Math.max(0, Math.min(maximo, anterior + Math.sign(distancia) * Math.min(44, Math.abs(distancia))));
        if (this.container.scrollTop === anterior) return;
        this.aoRolar();
        this.marcar(this.ponteiroY);
        this.frame = requestAnimationFrame(() => this.rolar());
    }

    soltar(evento) {
        if (!this.origem || this.bloqueado()) return;
        evento.preventDefault();
        this.marcar(evento.clientY);
        const origem = this.origem;
        // Posição é uma divisória entre páginas; destino é o índice final, 1-based.
        const destino = this.posicao > origem - 1 ? this.posicao : this.posicao + 1;
        this.cancelar();
        if (destino !== origem) this.aoMover(origem, destino);
    }

    pararRolagem() {
        if (this.frame) cancelAnimationFrame(this.frame);
        this.frame = null;
    }

    cancelar() {
        this.pararRolagem();
        this.botaoOrigem?.classList.toggle("pdf-miniatura-arrastando", false);
        this.marcador?.remove();
        this.marcador = this.botaoOrigem = null;
        this.origem = this.posicao = null;
        this.aoRolar();
    }

    destruir() {
        this.cancelar();
        for (const [nome, handler] of Object.entries(this.handlers)) this.container.removeEventListener(nome, handler);
    }
}
