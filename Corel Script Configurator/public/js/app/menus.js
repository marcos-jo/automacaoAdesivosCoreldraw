// Descritores chamam comandos reais do módulo; estado é consultado ao abrir.
export class MenusAplicacao {
    constructor(container, informar) {
        this.container = container; this.informar = informar; this.itens = new Map();
        container.addEventListener("toggle", e => {
            if (e.target.open) for (const detalhes of container.querySelectorAll("details")) if (detalhes !== e.target) detalhes.open = false;
            this.atualizar();
        }, true);
        container.addEventListener("click", e => {
            const resumo = e.target.closest("summary");
            if (resumo) { for (const d of container.querySelectorAll("details")) if (d !== resumo.closest("details")) d.open = false; return; }
            const botao = e.target.closest("button[data-command]");
            if (!botao) return;
            const item = this.itens.get(botao.dataset.command)?.item;
            if (!item || item.podeExecutar?.() === false) return;
            this.fechar();
            Promise.resolve().then(() => item.executar()).catch(erro => informar(erro.message));
        });
        container.addEventListener("keydown", e => {
            const detalhes = e.target.closest("details"); if (!detalhes) return;
            if (e.key === "Escape") { e.preventDefault(); detalhes.open = false; detalhes.querySelector("summary").focus(); }
            if (!["ArrowDown", "ArrowUp"].includes(e.key)) return;
            e.preventDefault();
            for (const d of container.querySelectorAll("details")) if (d !== detalhes) d.open = false;
            detalhes.open = true; this.atualizar();
            const botoes = [...detalhes.querySelectorAll("button")].filter(b => !b.disabled);
            if (!botoes.length) return;
            const indice = botoes.indexOf(e.target), passo = e.key === "ArrowDown" ? 1 : -1;
            const destino = indice < 0 ? (passo > 0 ? 0 : botoes.length - 1) : (indice + passo + botoes.length) % botoes.length;
            botoes[destino].focus();
        });
        container.closest("#app-shell").addEventListener("pointerdown", e => { if (!container.contains(e.target)) this.fechar(); });
    }
    renderizar(aba, acoes) {
        this.container.replaceChildren(); this.itens.clear();
        this.container.hidden = aba.id === "inicio";
        if (aba.id === "inicio") return;
        const grupos = (aba.controle.menus?.() || []).map(g => ({ ...g, itens:[...g.itens] }));
        let arquivo = grupos.find(g => g.id === "arquivo");
        if (!arquivo) { arquivo = { id:"arquivo", titulo:"Arquivo", itens:[] }; grupos.unshift(arquivo); }
        arquivo.itens.push({ id:"fechar-aba", titulo:"Fechar aba", executar:acoes.fechar });
        let exibir = grupos.find(g => g.id === "exibir");
        if (!exibir) { exibir = { id:"exibir", titulo:"Exibir", itens:[] }; grupos.push(exibir); }
        exibir.itens.push({ id:"inicio", titulo:"Início", executar:acoes.inicio });
        for (const grupo of grupos) {
            const detalhes = document.createElement("details"); detalhes.className = "app-menu";
            const resumo = document.createElement("summary"); resumo.id = "app-menu-" + grupo.id; resumo.textContent = grupo.titulo; resumo.setAttribute("aria-haspopup","menu");
            const lista = document.createElement("div"); lista.setAttribute("role","menu"); lista.setAttribute("aria-label",grupo.titulo);
            for (const item of grupo.itens) {
                const botao = document.createElement("button"); botao.type = "button"; botao.dataset.command = item.id; botao.setAttribute("role","menuitem");
                this.itens.set(item.id,{item,botao}); lista.append(botao);
            }
            detalhes.append(resumo,lista); this.container.append(detalhes);
        }
        this.atualizar();
    }
    atualizar() {
        for (const {item,botao} of this.itens.values()) {
            botao.textContent = typeof item.titulo === "function" ? item.titulo() : item.titulo;
            botao.disabled = item.podeExecutar?.() === false;
        }
    }
    fechar() { for (const detalhes of this.container.querySelectorAll("details")) detalhes.open = false; }
}
