export class InterfacePrincipal {
    constructor(raiz, modulos) {
        this.raiz = raiz; this.modulos = modulos;
        raiz.innerHTML = `<header class="app-cabecalho"><strong class="app-marca">4PrintShops</strong><nav id="app-abas" role="tablist" aria-label="Ferramentas abertas"></nav></header>
            <nav id="app-menus" aria-label="Menus da ferramenta"></nav>
            <main id="app" class="app-workspaces"></main><p id="app-mensagem" role="status" hidden></p>`;
        this.conteudo = raiz.querySelector("#app"); this.barra = raiz.querySelector("#app-abas");
        this.inicio = this.criarPainel("inicio");
        this.inicio.innerHTML = `<section class="app-inicio"><header class="app-inicio-heading"><div><p class="app-saudacao">ÁREA DE TRABALHO</p><h1>Seu próximo trabalho começa aqui</h1><p>Importe PDFs, imagens ou Word para organizar, preparar e imprimir.</p></div><button type="button" data-open-pdf class="app-abrir-pdf">Importar</button></header><h2>Ferramentas da gráfica</h2><div class="app-cartoes"></div><p class="app-futuras">Seus documentos abertos permanecem acessíveis nas abas superiores.</p></section>`;
        for (const modulo of modulos) {
            const botao = document.createElement("button"); botao.type = "button"; botao.dataset.tool = modulo.id; botao.className = "app-cartao";
            const titulo = document.createElement("strong"), descricao = document.createElement("span");
            titulo.textContent = modulo.titulo; descricao.textContent = modulo.descricao;
            botao.append(titulo, descricao); this.inicio.querySelector(".app-cartoes").append(botao);
        }
        raiz.addEventListener("click", e => {
            const botao = e.target.closest("button");
            if (botao?.hasAttribute('data-open-pdf')) Promise.resolve(this.aoAbrirPDF?.()).catch(erro=>this.informar(erro.message));
            else if (botao?.dataset.closeTool) this.aoFechar?.(botao.dataset.closeTool);
            else if (botao?.dataset.tool) this.aoAtivar?.(botao.dataset.tool);
        });
        this.barra.addEventListener("keydown", e => {
            const botoes = [...this.barra.querySelectorAll('[role="tab"]')], i = botoes.indexOf(e.target);
            if (i < 0) return;
            if (e.key === "Delete" && e.target.dataset.tool !== "inicio") { e.preventDefault(); this.aoFechar?.(e.target.dataset.tool); }
            if (!["ArrowLeft", "ArrowRight", "Home", "End"].includes(e.key)) return;
            e.preventDefault(); const indice = e.key === "Home" ? 0 : e.key === "End" ? botoes.length - 1 : (i + (e.key === "ArrowRight" ? 1 : -1) + botoes.length) % botoes.length;
            this.aoAtivar?.(botoes[indice].dataset.tool).then(() => this.barra.querySelector('[aria-selected="true"]')?.focus(), () => {});
        });
    }
    criarPainel(id) {
        const painel = document.createElement("section"); painel.id = "workspace-" + id; painel.className = "app-workspace";
        painel.setAttribute("role", "tabpanel"); painel.setAttribute("aria-labelledby", "aba-" + id); painel.hidden = true;
        this.conteudo.append(painel); return painel;
    }
    removerPainel(painel) { painel.remove(); }
    exibir(ativa, abas) {
        this.barra.replaceChildren();
        for (const aba of abas) {
            aba.painel.hidden = aba !== ativa;
            const grupo = document.createElement("div"); grupo.className = "app-aba";
            const botao = document.createElement("button"); botao.type = "button"; botao.id = "aba-" + aba.id; botao.dataset.tool = aba.id;
            botao.textContent = aba.id === "inicio" ? "Início" : aba.titulo;
            botao.setAttribute("role", "tab"); botao.setAttribute("aria-controls", aba.painel.id); botao.setAttribute("aria-selected", String(aba === ativa)); botao.tabIndex = aba === ativa ? 0 : -1;
            grupo.append(botao);
            if (aba.id !== "inicio") {
                const fechar = document.createElement("button"); fechar.type = "button"; fechar.dataset.closeTool = aba.id;
                fechar.textContent = "×"; fechar.setAttribute("aria-label", "Fechar aba " + aba.titulo); grupo.append(fechar);
            }
            this.barra.append(grupo);
        }
        this.raiz.dataset.activeTool = ativa.id; this.informar(""); this.aoExibir?.(ativa);
    }
    informar(texto) { const mensagem = this.raiz.querySelector("#app-mensagem"); mensagem.textContent = texto; mensagem.hidden = !texto; }
}
