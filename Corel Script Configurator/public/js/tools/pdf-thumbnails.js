import { PDFArrasteMiniaturas } from "./pdf-thumbnail-drag.js";
import { PDFTrabalhoMiniatura } from "./pdf-thumbnail-job.js";
import { PDFSelecaoPaginas } from "./pdf-page-selection.js";

const ALTURA_ITEM = 166;
const MARGEM = 2;

// Janela virtual: somente páginas próximas da área visível têm botões/canvases.
export class PDFMiniaturas {
    constructor(container, aoSelecionar, aoMover = () => {}, { tempoLimite = 10000 } = {}) {
        this.container = container;
        this.aoSelecionar = aoSelecionar;
        this.itens = new Map();
        // Inclui canvases visíveis e fora da janela; cerca de 1,8 MiB no máximo.
        this.cache = new Map();
        this.limiteCache = 32;
        this.geracao = 0;
        this.selecionada = 1;
        this.selecao = new PDFSelecaoPaginas();
        this.tempoLimite = tempoLimite;
        this.versaoRolagem = 0;
        this.arraste = new PDFArrasteMiniaturas(container, {
            alturaItem: ALTURA_ITEM, obterLista: () => this.lista,
            obterTotal: () => this.pdf?.numPages || 0,
            bloqueado: () => !!this.ocupado || !!this.encerrada,
            obterSelecao: numero => {
                if (!this.selecao.numeros().includes(numero)) this.selecao.clicar(numero);
                this.atualizarJanela(); return this.selecao.numeros();
            }, aoMover, aoRolar: () => this.atualizarJanela()
        });
        this.rolar = () => { ++this.versaoRolagem; this.atualizarJanela(); };
        this.redimensionar = () => { this.ajustarAltura(); this.atualizarJanela(); };
        this.clicar = evento => {
            const botao = evento.target.closest("button[data-pagina]");
            if (botao && !this.ocupado && !this.encerrada) {
                const numero = Number(botao.dataset.pagina);
                evento.preventDefault();
                this.selecao.clicar(numero, evento); this.atualizarJanela();
                if (this.itens.get(numero)?.estado === "erro") this.tentarNovamente(numero);
                this.aoSelecionar(numero);
            }
        };
        container.addEventListener("scroll", this.rolar, { passive: true });
        container.addEventListener("click", this.clicar);
        if (typeof ResizeObserver !== "undefined") {
            this.observador = new ResizeObserver(this.redimensionar);
            this.observador.observe(container);
        }
        this.janela = globalThis.window;
        this.janela?.addEventListener("resize", this.redimensionar);
    }

    ajustarAltura() {
        const altura = globalThis.innerHeight;
        if (!altura) return;
        const topo = Math.max(12, this.container.getBoundingClientRect().top);
        const valor = `${Math.max(166, Math.min(780, altura - topo - 20))}px`;
        if (this.container.style.height !== valor) this.container.style.height = valor;
    }

    definirDocumento(pdf, selecionada = 1) {
        if (this.encerrada) return;
        this.pdf = pdf;
        this.ajustarAltura();
        this.selecionada = selecionada;
        this.selecao.atualizar(Array.from({ length: pdf.numPages }, (_, i) => pdf.paginas?.[i]?.id ?? String(i + 1)), selecionada);
        this.lista = document.createElement("div");
        this.lista.className = "pdf-miniaturas-lista";
        this.lista.style.height = `${pdf.numPages * ALTURA_ITEM}px`;
        this.container.replaceChildren(this.lista);
        // A área rolável precisa existir antes de posicionar a página selecionada.
        this.container.scrollTop = (selecionada - 1) * ALTURA_ITEM;
        this.atualizarJanela();
    }

    chave(numero) { return this.pdf.obterChaveMiniatura?.(numero) ?? String(numero); }

    invalidarIdentidades(ids) {
        for (const [chave, canvas] of this.cache) if (ids.has(chave.split(":")[0])) {
            this.cache.delete(chave); canvas.width = canvas.height = 0;
        }
    }

    guardarCanvas(chave, canvas) {
        this.cache.delete(chave);
        this.cache.set(chave, canvas);
        while (this.cache.size > this.limiteCache) {
            const remover = [...this.cache].find(([, c]) => ![...this.itens.values()].some(i => i.canvas === c));
            if (!remover) break;
            this.cache.delete(remover[0]);
            remover[1].width = remover[1].height = 0;
        }
    }

    async atualizarDocumento(pdf, selecionada) {
        await this.limpar({ preservarCache: true });
        const validas = new Set(Array.from({ length: pdf.numPages }, (_, i) => pdf.obterChaveMiniatura?.(i + 1) ?? String(i + 1)));
        for (const [chave, canvas] of this.cache) {
            if (!validas.has(chave)) {
                this.cache.delete(chave);
                canvas.width = canvas.height = 0;
            }
        }
        this.definirDocumento(pdf, selecionada);
    }

    definirOcupado(valor) {
        this.ocupado = valor;
        if (valor) this.arraste.cancelar();
        for (const item of this.itens.values()) {
            item.botao.disabled = valor;
            item.botao.draggable = !valor && this.pdf?.numPages > 1;
        }
    }

    focarSelecionada() { this.itens.get(this.selecionada)?.botao.focus(); }

    selecionar(numero, { versaoRolagem = this.versaoRolagem } = {}) {
        this.selecionada = numero;
        const topo = (numero - 1) * ALTURA_ITEM;
        const altura = this.container.clientHeight || 600;
        if (!this.arraste.origem && versaoRolagem === this.versaoRolagem &&
            (topo < this.container.scrollTop || topo + ALTURA_ITEM > this.container.scrollTop + altura)) {
            this.container.scrollTop = topo;
        }
        this.atualizarJanela();
    }

    removerItem(numero, item) {
        if (this.atual === item) this.trabalho?.cancelar();
        item.botao.remove();
        // O executor libera o canvas ativo depois de concluir/interromper a espera.
        if (this.atual !== item && item.canvas) {
            if (this.cache.get(item.chave) === item.canvas) item.canvas.remove();
            else item.canvas.width = item.canvas.height = 0;
        }
        this.itens.delete(numero);
    }

    atualizarJanela() {
        if (!this.pdf || this.encerrada || this.suspensa) return;
        const total = this.pdf.numPages;
        const visivel = Math.floor((this.container.scrollTop || 0) / ALTURA_ITEM);
        const quantidade = Math.min(this.arraste.origem ? 15 : 16,
            Math.ceil((this.container.clientHeight || 600) / ALTURA_ITEM) + 2 * MARGEM);
        const inicio = Math.max(1, Math.min(total - quantidade + 1, visivel + 1 - MARGEM));
        const fim = Math.min(total, inicio + quantidade - 1);
        for (const [numero, item] of this.itens) {
            // Manter a origem no DOM evita cancelar o arraste nativo ao rolar.
            if ((numero < inicio || numero > fim) && numero !== this.arraste.origem) this.removerItem(numero, item);
        }
        for (let numero = inicio; numero <= fim; numero++) {
            if (!this.itens.has(numero)) {
                const botao = document.createElement("button");
                botao.type = "button";
                botao.className = "pdf-miniatura";
                botao.dataset.pagina = String(numero);
                botao.setAttribute("aria-label", `Abrir página ${numero}. Arraste para mover.`);
                botao.style.top = `${(numero - 1) * ALTURA_ITEM}px`;
                const imagem = document.createElement("span");
                imagem.className = "pdf-miniatura-imagem";
                imagem.textContent = "Carregando…";
                const legenda = document.createElement("span");
                legenda.textContent = `Página ${numero}`;
                botao.append(imagem, legenda);
                this.lista.appendChild(botao);
                const chave = this.chave(numero);
                const canvas = this.cache.get(chave);
                const item = { numero, botao, imagem, chave, canvas };
                this.definirEstado(item, canvas ? "pronto" : "pendente");
                this.itens.set(numero, item);
                if (canvas) {
                    imagem.replaceChildren(canvas);
                    this.guardarCanvas(chave, canvas);
                }
            }
            const item = this.itens.get(numero);
            item.botao.disabled = !!this.ocupado;
            item.botao.draggable = !this.ocupado && total > 1;
            item.botao.setAttribute("aria-current", String(numero === this.selecionada));
            item.botao.setAttribute("aria-pressed", String(this.selecao.ids.has(this.selecao.lista[numero - 1])));
        }
        this.garantirFila();
    }

    definirEstado(item, estado) {
        item.estado = estado;
        item.botao.dataset.estado = estado;
        item.botao.setAttribute("aria-busy", String(estado === "pendente" || estado === "renderizando"));
        item.botao.classList.toggle("pdf-miniatura-erro", estado === "erro");
    }

    falhar(item, erro) {
        this.definirEstado(item, "erro");
        item.imagem.replaceChildren();
        item.imagem.textContent = "Prévia indisponível · Tentar novamente";
        item.botao.title = erro.message;
        item.botao.setAttribute("aria-label", `Página ${item.numero}. Prévia indisponível. Clique para tentar novamente e abrir a página.`);
    }

    tentarNovamente(numero) {
        const item = this.itens.get(numero);
        if (!item || item.estado !== "erro" || this.ocupado || this.encerrada) return;
        this.definirEstado(item, "pendente");
        item.imagem.textContent = "Carregando…";
        this.garantirFila();
    }

    garantirFila() {
        if (this.processamento || !this.pdf || this.encerrada || this.suspensa ||
            ![...this.itens.values()].some(item => item.estado === "pendente")) return;
        this.processamento = Promise.resolve().then(() => this.processarFila()).catch(erro => {
            for (const item of this.itens.values()) {
                if (item.estado === "pendente" || item.estado === "renderizando") this.falhar(item, erro);
            }
        }).finally(() => {
            this.processamento = null;
            // A seleção/rolagem pode ter criado trabalho enquanto a fila terminava.
            this.garantirFila();
        });
    }

    async processarFila() {
        const geracao = this.geracao;
        const pdf = this.pdf;
        while (pdf && this.pdf === pdf && geracao === this.geracao && !this.encerrada && !this.suspensa) {
            const proximo = [...this.itens].find(([, item]) => item.estado === "pendente");
            if (!proximo) return;
            const [numero, item] = proximo;
            const vigente = () => this.pdf === pdf && geracao === this.geracao && this.itens.get(numero) === item;
            this.definirEstado(item, "renderizando");
            this.atual = item;
            const trabalho = new PDFTrabalhoMiniatura(this.tempoLimite);
            this.trabalho = trabalho;
            try {
                const pagina = await trabalho.esperar(pdf.getPage(numero));
                if (!vigente()) continue;
                const base = pagina.getViewport({ scale: 1 });
                const viewport = pagina.getViewport({ scale: Math.min(112 / base.width, 126 / base.height, 1) });
                const canvas = document.createElement("canvas");
                canvas.className = "pdf-miniatura-canvas";
                item.canvas = canvas;
                canvas.width = Math.ceil(viewport.width);
                canvas.height = Math.ceil(viewport.height);
                canvas.setAttribute("aria-hidden", "true");
                trabalho.render = pagina.render({ canvasContext: canvas.getContext("2d"), viewport, canvas });
                await trabalho.esperar(trabalho.render.promise);
                if (vigente()) {
                    item.imagem.replaceChildren(canvas);
                    this.definirEstado(item, "pronto");
                    this.guardarCanvas(item.chave, canvas);
                    item.botao.title = "";
                    item.botao.setAttribute("aria-label", `Abrir página ${numero}. Arraste para mover.`);
                }
            } catch (erro) {
                if (vigente()) {
                    this.falhar(item, erro);
                }
            } finally {
                trabalho.concluir();
                if (item.estado !== "pronto" && item.canvas) item.canvas.width = item.canvas.height = 0;
                this.trabalho = this.atual = null;
            }
            // Ceder ao navegador evita monopolizar eventos durante muitas prévias.
            await new Promise(resolve => setTimeout(resolve, 0));
        }
    }

    async limpar({ preservarCache = false } = {}) {
        ++this.geracao;
        this.pdf = null;
        this.arraste.cancelar();
        this.trabalho?.cancelar();
        await this.processamento;
        for (const [numero, item] of this.itens) this.removerItem(numero, item);
        this.container.replaceChildren();
        this.lista = null;
        if (!preservarCache) {
            for (const canvas of this.cache.values()) canvas.width = canvas.height = 0;
            this.cache.clear();
        }
    }

    async pausar() {
        this.suspensa = true; ++this.geracao; this.arraste.cancelar(); this.trabalho?.cancelar();
        await this.processamento;
        for (const item of this.itens.values()) if (item.estado === "renderizando") {
            this.definirEstado(item, "pendente"); item.canvas = null;
            item.imagem.replaceChildren(); item.imagem.textContent = "Carregando…";
        }
    }
    retomar() { this.suspensa = false; this.ajustarAltura(); this.atualizarJanela(); }
    destruir() {
        if (!this.encerramento) {
            this.encerrada = true;
            this.arraste.destruir();
            this.container.removeEventListener("scroll", this.rolar);
            this.container.removeEventListener("click", this.clicar);
            this.observador?.disconnect();
            this.janela?.removeEventListener("resize", this.redimensionar);
            this.encerramento = this.limpar();
        }
        return this.encerramento;
    }
}
