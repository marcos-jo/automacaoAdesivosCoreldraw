// Navegação visual: não altera bytes ou revisões do workspace.
export class PDFControlesVisualizacao {
    constructor(buscar, previa, atualizar, permitido) {
        Object.assign(this, { previa, atualizar, permitido });
        this.container = buscar("#pdf-visualizacao"); this.modo = "pagina"; this.zoom = 1; this.listeners = [];
        this.evento = (alvo, tipo, acao, opcoes) => {
            alvo.addEventListener(tipo, acao, opcoes);
            this.listeners.push(() => alvo.removeEventListener(tipo, acao, opcoes));
        };
        this.rotulo = buscar("#pdf-zoom-valor");
        this.status = buscar("#pdf-status");
        previa.calcularEscala = base => {
            this.medida = [this.container.clientWidth, this.container.clientHeight];
            const w = Math.max(80, (this.container.clientWidth || 898) - 32), h = Math.max(80, (this.container.clientHeight || 650) - 32);
            const escala = this.modo === "pagina" ? Math.min(w / base.width, h / base.height) : this.modo === "largura" ? w / base.width : this.zoom;
            this.zoom = Math.max(Number.EPSILON, Math.min(6, escala, Math.sqrt(16000000 / (base.width * base.height))));
            this.rotulo.textContent = `${Math.round(this.zoom * 100)}%`;
            return this.zoom;
        };
        for (const modo of ["pagina", "largura", "real"]) this.evento(buscar("#pdf-zoom-" + modo), "click", () => {
            if (!permitido()) return;
            this.modo = modo; if (modo === "real") this.zoom = 1; this.ancora = null; this.solicitar();
        });
        for (const [id, fator] of [["mais", 1.2], ["menos", 1 / 1.2]]) this.evento(buscar("#pdf-zoom-" + id), "click", () => this.ampliar(fator));
        this.evento(this.container, "wheel", e => {
            if (!permitido()) return;
            e.preventDefault();
            const delta = e.deltaY * (e.deltaMode === 1 ? 16 : e.deltaMode === 2 ? this.container.clientHeight : 1);
            this.ampliar(Math.exp(-Math.max(-200, Math.min(200, delta)) * .002), { x: e.clientX, y: e.clientY });
        }, { passive: false });
        this.evento(this.container, "pointerdown", e => {
            if (!permitido() || ![0, 1].includes(e.button) || e.target.tagName !== "CANVAS") return;
            e.preventDefault(); this.pan = { id: e.pointerId, x: e.clientX, y: e.clientY, left: this.container.scrollLeft, top: this.container.scrollTop };
            this.container.setPointerCapture(e.pointerId); this.container.classList.toggle("pdf-arrastando-pagina", true);
        });
        this.evento(this.container, "pointermove", e => {
            if (this.pan?.id !== e.pointerId) return;
            this.container.scrollLeft = this.pan.left + this.pan.x - e.clientX;
            this.container.scrollTop = this.pan.top + this.pan.y - e.clientY;
        });
        const soltar = () => { this.pan = null; this.container.classList.toggle("pdf-arrastando-pagina", false); };
        for (const tipo of ["pointerup", "pointercancel", "lostpointercapture"]) this.evento(this.container, tipo, soltar);
        if (typeof ResizeObserver !== "undefined") {
            this.observador = new ResizeObserver(() => { if (permitido()) this.solicitar(true); }); this.observador.observe(this.container);
        }
    }
    ampliar(fator, ponteiro) {
        if (!this.permitido()) return;
        const c = this.container, r = c.querySelector("canvas")?.getBoundingClientRect();
        const limites = c.getBoundingClientRect(), x = ponteiro?.x ?? limites.left + c.clientWidth / 2, y = ponteiro?.y ?? limites.top + c.clientHeight / 2;
        if (r) this.ancora = { x, y, fx: (x - r.left) / r.width, fy: (y - r.top) / r.height };
        this.modo = "manual"; this.zoom = Math.max(.05, Math.min(6, this.zoom * fator)); this.solicitar();
    }
    solicitar(redimensionamento = false) {
        clearTimeout(this.timer);
        const geracao = this.geracao = (this.geracao || 0) + 1;
        this.timer = setTimeout(async () => {
            if (this.encerrada || !this.permitido()) return;
            const exibicao = this.previa.exibicao;
            if (redimensionamento && exibicao?.base && this.medida?.[0] === this.container.clientWidth && this.medida?.[1] === this.container.clientHeight && exibicao.escala === this.previa.calcularEscala(exibicao.base)) return;
            const ancora = this.ancora; this.ancora = null;
            const documento = this.previa.documento;
            try { await this.atualizar(); } catch (erro) { if (!this.encerrada && this.status) this.status.textContent = erro.message; return; }
            if (this.encerrada || geracao !== this.geracao || !this.permitido() || documento !== this.previa.documento) return;
            const r = this.container.querySelector("canvas")?.getBoundingClientRect();
            if (ancora && r) {
                this.container.scrollLeft += r.left + ancora.fx * r.width - ancora.x;
                this.container.scrollTop += r.top + ancora.fy * r.height - ancora.y;
            } else if (this.modo !== "manual") this.container.scrollTop = this.container.scrollLeft = 0;
        }, 60);
    }
    paginaExibida(chave, fonte) {
        if (chave !== this.pagina || fonte !== this.fonte) {
            this.container.scrollTop = this.container.scrollLeft = 0; this.pan = this.ancora = null;
            this.pagina = chave; this.fonte = fonte;
        }
    }
    pausar() { clearTimeout(this.timer); this.geracao = (this.geracao || 0) + 1; this.pan = this.ancora = null; this.container.classList.toggle("pdf-arrastando-pagina", false); }
    destruir() { this.encerrada = true; this.pausar(); this.observador?.disconnect(); for (const remover of this.listeners) remover(); delete this.previa.calcularEscala; }
}
