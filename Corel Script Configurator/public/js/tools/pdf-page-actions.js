// Menu contextual e acesso equivalente por botão/teclado, restritos à instância.
export class PDFAcoesPaginas {
    constructor(raiz, miniaturas, acoes) {
        Object.assign(this, { raiz, miniaturas, acoes }); this.listeners = [];
        this.menu = raiz.querySelector("#pdf-paginas-menu");
        const evento = (alvo, tipo, acao) => { alvo.addEventListener(tipo, acao); this.listeners.push(() => alvo.removeEventListener(tipo, acao)); };
        evento(miniaturas.container, "contextmenu", e => {
            const botao = e.target.closest("button[data-pagina]"); if (!botao || miniaturas.ocupado) return;
            e.preventDefault(); const n = Number(botao.dataset.pagina);
            if (!miniaturas.selecao.numeros().includes(n)) miniaturas.selecao.clicar(n);
            miniaturas.atualizarJanela(); this.abrir(e.clientX, e.clientY, botao);
        });
        evento(raiz.querySelector("#pdf-paginas-acoes"), "click", e => {
            const r = e.target.getBoundingClientRect(); this.abrir(r.left, r.bottom, e.target);
        });
        for (const [id, acao] of [["duplicar", acoes.duplicar], ["excluir-selecao", acoes.excluir]]) {
            evento(raiz.querySelector("#pdf-" + id), "click", () => { this.fechar(); acao(); });
        }
        evento(raiz, "pointerdown", e => { if (!this.menu.contains(e.target)) this.fechar(false); });
        evento(this.menu, "focusout", e => { if (!this.menu.contains(e.relatedTarget)) this.fechar(false); });
        evento(raiz, "keydown", e => {
            if (e.shiftKey && e.key === "F10" && e.target.closest("button[data-pagina]")) {
                const n = Number(e.target.closest("button[data-pagina]").dataset.pagina);
                if (!miniaturas.selecao.numeros().includes(n)) { miniaturas.selecao.clicar(n); miniaturas.atualizarJanela(); }
                e.preventDefault(); const r = e.target.getBoundingClientRect(); this.abrir(r.left, r.bottom, e.target);
            }
            if (this.menu.hidden) return;
            const botoes = [...this.menu.querySelectorAll("button")];
            if (e.key === "Escape") { e.preventDefault(); this.fechar(); }
            if (["ArrowDown", "ArrowUp"].includes(e.key)) {
                e.preventDefault(); const i = botoes.indexOf(e.target), passo = e.key === "ArrowDown" ? 1 : -1;
                botoes[(i + passo + botoes.length) % botoes.length].focus();
            }
        });
    }
    abrir(x, y, origem) {
        if (this.miniaturas.ocupado || !this.miniaturas.pdf) return;
        this.origem = origem; this.menu.hidden = false;
        this.raiz.querySelector("#pdf-paginas-acoes").setAttribute("aria-expanded", "true");
        this.menu.style.left = `${Math.max(4, Math.min(x, (globalThis.innerWidth || 1000) - 200))}px`;
        this.menu.style.top = `${Math.max(4, Math.min(y, (globalThis.innerHeight || 800) - 120))}px`;
        this.menu.querySelector("button").focus();
    }
    fechar(focar = true) { this.menu.hidden = true; this.raiz.querySelector("#pdf-paginas-acoes").setAttribute("aria-expanded", "false"); if (focar) this.origem?.focus(); }
    destruir() { this.fechar(false); for (const remover of this.listeners) remover(); }
}
