const CHAVE = "4printshops.pdf.paineis.v1";
const limites = { esquerda: [134, 320], direita: [190, 420] };
export class PDFPreferenciasPaineis {
    constructor(armazenamento = {
        ler: chave => globalThis.localStorage?.getItem(chave),
        gravar: (chave, valor) => globalThis.localStorage?.setItem(chave, valor)
    }) {
        this.armazenamento = armazenamento; this.valor = { esquerda: 150, direita: 230 };
        try {
            const dados = JSON.parse(armazenamento.ler(CHAVE));
            if (dados?.versao === 1) for (const lado of Object.keys(limites)) {
                const n = dados[lado], [min, max] = limites[lado];
                if (typeof n === "number" && Number.isFinite(n) && n >= min && n <= max) this.valor[lado] = n;
            }
        } catch { /* Preferências inválidas/indisponíveis: padrões em memória. */ }
    }
    salvar(lado, valor) {
        const [min, max] = limites[lado];
        if (!Number.isFinite(valor)) return;
        this.valor[lado] = Math.round(Math.max(min, Math.min(max, valor)));
        try { this.armazenamento.gravar(CHAVE, JSON.stringify({ versao: 1, ...this.valor })); } catch { /* Sessão em memória. */ }
    }
}

export class PDFPaineis {
    constructor(buscar, atualizar, preferencias = new PDFPreferenciasPaineis()) {
        Object.assign(this, { buscar, atualizar, preferencias }); this.listeners = [];
        this.area = buscar("#pdf-area-trabalho"); this.raiz = buscar("#pdf-ferramenta");
        const evento = (elemento, tipo, acao) => { elemento.addEventListener(tipo, acao); this.listeners.push(() => elemento.removeEventListener(tipo, acao)); };
        for (const lado of Object.keys(limites)) {
            const divisor = buscar("#pdf-divisor-" + lado);
            evento(divisor, "pointerdown", e => {
                if (e.button !== 0) return; e.preventDefault();
                this.arraste = { lado, id: e.pointerId, x: e.clientX, valor: this.efetivas[lado] };
                divisor.setPointerCapture(e.pointerId);
            });
            evento(divisor, "pointermove", e => {
                if (this.arraste?.id !== e.pointerId) return;
                preferencias.salvar(lado, this.arraste.valor + (e.clientX - this.arraste.x) * (lado === "esquerda" ? 1 : -1)); this.aplicar();
            });
            for (const tipo of ["pointerup", "pointercancel", "lostpointercapture"]) evento(divisor, tipo, () => { this.arraste = null; });
            evento(divisor, "keydown", e => {
                if (!["ArrowLeft", "ArrowRight", "Home", "End"].includes(e.key)) return;
                e.preventDefault(); const [min, max] = limites[lado];
                const valor = e.key === "Home" ? min : e.key === "End" ? max : this.efetivas[lado] + (e.key === "ArrowRight" ? 10 : -10) * (lado === "esquerda" ? 1 : -1);
                preferencias.salvar(lado, valor); this.aplicar();
            });
        }
        if (typeof ResizeObserver !== "undefined") { this.observador = new ResizeObserver(() => this.aplicar()); this.observador.observe(this.area); }
        this.aplicar();
    }
    aplicar() {
        if (this.suspensa) return;
        const total = this.area.clientWidth || 1000, estreita = total < 560;
        const esquerda = Math.min(this.preferencias.valor.esquerda, Math.max(100, total - 216));
        const direita = Math.min(this.preferencias.valor.direita, Math.max(190, total - esquerda - 236));
        this.efetivas = { esquerda: Math.min(esquerda, Math.max(134, total - direita - 236)), direita };
        if (estreita && !this.estreita && !this.buscar("#pdf-painel-ferramentas").hidden) this.buscar("#pdf-painel-alternar").click();
        this.estreita = estreita; this.raiz.classList.toggle("pdf-janela-estreita", estreita);
        for (const lado of Object.keys(limites)) {
            this.raiz.style.setProperty?.("--pdf-" + lado, this.efetivas[lado] + "px");
            const divisor = this.buscar("#pdf-divisor-" + lado), [min, max] = limites[lado];
            divisor.setAttribute("aria-valuemin", String(min)); divisor.setAttribute("aria-valuemax", String(max)); divisor.setAttribute("aria-valuenow", String(Math.round(this.efetivas[lado])));
        }
    }
    destruir() { this.observador?.disconnect(); for (const remover of this.listeners) remover(); this.arraste = null; }
}
