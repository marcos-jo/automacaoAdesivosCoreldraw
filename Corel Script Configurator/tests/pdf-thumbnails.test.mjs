import assert from "node:assert/strict";
import test from "node:test";
import { setTimeout as pausa } from "node:timers/promises";
import { Elemento, JanelaTeste } from "./helpers/pdf-dom.mjs";
import { PDFMiniaturas } from "../public/js/tools/pdf-thumbnails.js";

async function aguardar(condicao) {
    const limite = Date.now() + 1000;
    while (!condicao()) {
        assert.ok(Date.now() < limite, "Estado da miniatura não concluiu");
        await pausa(2);
    }
}

function documento(total, giros = 0) {
    const renderizadas = [];
    const pdf = {
        numPages: total,
        async getPage(numero) {
            return {
                getViewport({ scale }) {
                    return giros === 90
                        ? { width: 300 * scale, height: 200 * scale }
                        : { width: 200 * scale, height: 300 * scale };
                },
                render({ canvas }) {
                    renderizadas.push({ numero, canvas });
                    return { promise: Promise.resolve(), cancel() {} };
                }
            };
        }
    };
    return { pdf, renderizadas };
}

test("pausa interrompe fila sem perder seleção/scroll e retomada só renderiza itens pendentes", async () => {
    globalThis.document = { createElement: tag => new Elemento(tag) };
    globalThis.window = new JanelaTeste();
    const container = new Elemento("div"); container.clientHeight=600;
    const miniaturas = new PDFMiniaturas(container,()=>{}), {pdf,renderizadas}=documento(350);
    try {
        miniaturas.definirDocumento(pdf); await miniaturas.processamento;
        miniaturas.selecao.clicar(2); const ids=new Set(miniaturas.selecao.ids);
        await miniaturas.pausar(); const total=renderizadas.length;
        container.scrollTop=100*166; container.emitir("scroll"); await pausa(30);
        assert.equal(renderizadas.length,total); assert.deepEqual(miniaturas.selecao.ids,ids);
        miniaturas.retomar(); await miniaturas.processamento;
        assert.ok(renderizadas.length>total); assert.ok(miniaturas.itens.size<=16);
        assert.equal(container.scrollTop,100*166);
    } finally {await miniaturas.destruir();delete globalThis.document;delete globalThis.window;}
});

test("350 páginas: renderização sob demanda limitada, sem duplicação na mesma janela", async () => {
    globalThis.document = { createElement: tag => new Elemento(tag) };
    globalThis.window = new JanelaTeste();
    const container = new Elemento("div");
    container.clientHeight = 600;
    const selecionadas = [];
    const miniaturas = new PDFMiniaturas(container, numero => selecionadas.push(numero));
    try {
        const { pdf, renderizadas } = documento(350);
        miniaturas.definirDocumento(pdf);
        await miniaturas.processamento;
        assert.ok(renderizadas.length < 12, "Não deve renderizar as 350 páginas de uma vez");
        assert.ok(container.querySelectorAll("canvas").length < 12);
        const iniciais = renderizadas.length;
        miniaturas.atualizarJanela();
        await miniaturas.processamento;
        assert.equal(renderizadas.length, iniciais);

        container.scrollTop = 349 * 166;
        container.emitir("scroll");
        await miniaturas.processamento;
        assert.ok(renderizadas.length < 24);
        assert.equal(renderizadas.at(-1).numero, 350);
        assert.ok(renderizadas[0].canvas.width > 0, "Cache limitado conserva a prévia para reutilização");
        assert.ok(miniaturas.cache.size <= miniaturas.limiteCache);
        const pagina350 = container.querySelectorAll("button").find(botao => botao.dataset.pagina === "350");
        pagina350.click();
        assert.deepEqual(selecionadas, [350]);
        miniaturas.selecionar(350);
        assert.equal(pagina350.getAttribute("aria-current"), "true");
        miniaturas.definirOcupado(true);
        pagina350.emitir("click");
        assert.deepEqual(selecionadas, [350], "Operação em andamento bloqueia navegação");
    } finally {
        await miniaturas.destruir();
        assert.equal(container.querySelectorAll("canvas").length, 0);
        assert.equal(container.listeners.get("scroll").length, 0);
        assert.equal(window.listeners.get("resize").size, 0);
        const reentrada = new PDFMiniaturas(container, () => {});
        assert.equal(window.listeners.get("resize").size, 1);
        await reentrada.destruir();
        assert.equal(window.listeners.get("resize").size, 0);
        delete globalThis.document;
        delete globalThis.window;
    }
});

test("substituição após giro/exclusão/união refaz prévias com nova orientação e quantidade", async () => {
    globalThis.document = { createElement: tag => new Elemento(tag) };
    const container = new Elemento("div");
    const miniaturas = new PDFMiniaturas(container, () => {});
    try {
        const a = documento(3);
        miniaturas.definirDocumento(a.pdf);
        await miniaturas.processamento;
        const canvasAntigo = a.renderizadas[0].canvas;
        assert.ok(canvasAntigo.height > canvasAntigo.width);
        await miniaturas.limpar();
        const b = documento(2, 90);
        miniaturas.definirDocumento(b.pdf, 2);
        await miniaturas.processamento;
        assert.equal(canvasAntigo.width, 0);
        assert.equal(container.querySelectorAll("button").length, 2);
        assert.ok(b.renderizadas[0].canvas.width > b.renderizadas[0].canvas.height);
        assert.equal(container.querySelectorAll("button")[1].getAttribute("aria-current"), "true");
        await miniaturas.limpar();
        miniaturas.definirDocumento(documento(5).pdf);
        await miniaturas.processamento;
        assert.equal(container.querySelectorAll("button").length, 5);
    } finally {
        await miniaturas.destruir();
        delete globalThis.document;
    }
});

test("encerra com renderização pendente, cancela e libera canvas sem aceitar cliques antigos", async () => {
    globalThis.document = { createElement: tag => new Elemento(tag) };
    const container = new Elemento("div");
    let cancelamentos = 0;
    let canvas;
    const miniaturas = new PDFMiniaturas(container, () => assert.fail("Listener deveria ter sido removido"));
    const pdf = {
        numPages: 1,
        async getPage() {
            return {
                getViewport: ({ scale }) => ({ width: 200 * scale, height: 300 * scale }),
                render(opcoes) {
                    canvas = opcoes.canvas;
                    let rejeitar;
                    return {
                        promise: new Promise((_, reject) => { rejeitar = reject; }),
                        cancel() {
                            cancelamentos++;
                            rejeitar(Object.assign(new Error("Cancelada"), { name: "RenderingCancelledException" }));
                        }
                    };
                }
            };
        }
    };
    try {
        miniaturas.definirDocumento(pdf);
        await aguardar(() => canvas);
        assert.ok(canvas);
        const botao = container.querySelector("button");
        const encerramento = miniaturas.destruir();
        assert.equal(miniaturas.destruir(), encerramento);
        await encerramento;
        assert.equal(cancelamentos, 1);
        assert.equal(canvas.width, 0);
        botao.emitir("click");
        assert.equal(container.querySelectorAll("button").length, 0);
    } finally {
        await miniaturas.destruir();
        delete globalThis.document;
    }
});

test("arraste indica a inserção, solicita a posição final e ignora movimento sem efeito", async () => {
    globalThis.document = { createElement: tag => new Elemento(tag) };
    globalThis.requestAnimationFrame = () => 1;
    globalThis.cancelAnimationFrame = () => {};
    const container = new Elemento("div");
    const movimentos = [];
    const miniaturas = new PDFMiniaturas(container, () => {}, (...movimento) => movimentos.push(movimento));
    const dataTransfer = { setData() {} };
    try {
        miniaturas.definirDocumento(documento(4).pdf);
        await miniaturas.processamento;
        const origem = container.querySelectorAll("button")[1];
        assert.equal(origem.draggable, true);
        origem.emitir("dragstart", origem, { dataTransfer });
        const sobre = container.emitir("dragover", container, { dataTransfer, clientY: 4 * 166 });
        assert.equal(sobre.defaultPrevented, true);
        assert.equal(miniaturas.arraste.marcador.hidden, false);
        assert.equal(miniaturas.arraste.marcador.style.top, "660px");
        container.emitir("drop", container, { dataTransfer, clientY: 4 * 166 });
        assert.deepEqual(movimentos, [[2, 4]]);
        assert.equal(miniaturas.arraste.marcador, null);
        origem.emitir("dragstart", origem, { dataTransfer });
        container.emitir("drop", container, { dataTransfer, clientY: 2 * 166 });
        assert.deepEqual(movimentos, [[2, 4]]);
        miniaturas.definirOcupado(true);
        assert.equal(origem.draggable, false);
        assert.equal(origem.emitir("dragstart", origem, { dataTransfer }).defaultPrevented, true);
        container.emitir("drop", container, { dataTransfer, clientY: 0 });
        assert.deepEqual(movimentos, [[2, 4]]);
    } finally {
        await miniaturas.destruir();
        for (const nome of ["dragstart", "dragover", "drop", "dragend", "dragleave"]) {
            assert.equal(container.listeners.get(nome).length, 0);
        }
        delete globalThis.document;
        delete globalThis.requestAnimationFrame;
        delete globalThis.cancelAnimationFrame;
    }
});

test("arrastar entre extremos de 350 páginas mantém a origem e limita a janela", async () => {
    globalThis.document = { createElement: tag => new Elemento(tag) };
    globalThis.cancelAnimationFrame = () => {};
    const container = new Elemento("div");
    container.clientHeight = 600;
    const movimentos = [];
    const miniaturas = new PDFMiniaturas(container, () => {}, (...movimento) => movimentos.push(movimento));
    const dataTransfer = { setData() {} };
    try {
        miniaturas.definirDocumento(documento(350).pdf);
        await miniaturas.processamento;
        const origem = container.querySelector("button");
        origem.emitir("dragstart", origem, { dataTransfer });
        container.scrollTop = 349 * 166;
        container.emitir("scroll");
        await miniaturas.processamento;
        assert.equal(container.contains(origem), true);
        assert.ok(container.querySelectorAll("button").length <= 16);
        assert.ok(container.querySelectorAll("canvas").length <= 16);
        assert.ok(container.querySelectorAll("button").some(b => b.dataset.pagina === "350"));
        container.emitir("drop", container, { dataTransfer, clientY: 600 });
        assert.deepEqual(movimentos, [[1, 350]]);
        assert.equal(container.contains(origem), false);
    } finally {
        await miniaturas.destruir();
        delete globalThis.document;
        delete globalThis.cancelAnimationFrame;
    }
});

test("rolagem nas bordas durante arraste é cancelada ao trocar de documento", async () => {
    globalThis.document = { createElement: tag => new Elemento(tag) };
    const quadros = new Map();
    let id = 0;
    globalThis.requestAnimationFrame = callback => { quadros.set(++id, callback); return id; };
    globalThis.cancelAnimationFrame = atual => quadros.delete(atual);
    const container = new Elemento("div");
    container.clientHeight = 600;
    container.scrollHeight = 350 * 166;
    const miniaturas = new PDFMiniaturas(container, () => {});
    const dataTransfer = { setData() {} };
    try {
        miniaturas.definirDocumento(documento(350).pdf);
        await miniaturas.processamento;
        const origem = container.querySelector("button");
        origem.emitir("dragstart", origem, { dataTransfer });
        container.emitir("dragover", container, { dataTransfer, clientY: 595 });
        const [atual, callback] = [...quadros][0];
        quadros.delete(atual);
        callback();
        assert.ok(container.scrollTop > 0);
        assert.ok(quadros.size > 0);
        await miniaturas.limpar();
        assert.equal(quadros.size, 0);
        assert.equal(miniaturas.arraste.origem, null);
    } finally {
        await miniaturas.destruir();
        delete globalThis.document;
        delete globalThis.requestAnimationFrame;
        delete globalThis.cancelAnimationFrame;
    }
});

test("regressão: cancelar render que não responde não bloqueia limpeza e próxima geração", async () => {
    globalThis.document = { createElement: tag => new Elemento(tag) };
    const container = new Elemento("div");
    const miniaturas = new PDFMiniaturas(container, () => {}, () => {}, { tempoLimite: 5000 });
    let iniciada = false;
    let liberar;
    const pendente = new Promise(resolve => { liberar = resolve; });
    const pdf = { numPages: 1, getPage: async () => ({
        getViewport: ({ scale }) => ({ width: 200 * scale, height: 300 * scale }),
        render: () => { iniciada = true; return { promise: pendente, cancel() {} }; }
    }) };
    let limpeza;
    try {
        miniaturas.definirDocumento(pdf);
        await aguardar(() => iniciada);
        limpeza = miniaturas.limpar();
        const resultado = await Promise.race([limpeza.then(() => "limpa"), pausa(100).then(() => "bloqueada")]);
        assert.equal(resultado, "limpa");
        miniaturas.definirDocumento(documento(2).pdf);
        await miniaturas.processamento;
        assert.equal(container.querySelectorAll("canvas").length, 2);
    } finally {
        liberar();
        await limpeza;
        await miniaturas.destruir();
        delete globalThis.document;
    }
});

test("regressão: janela alterada na conclusão da fila não deixa itens pendentes sem executor", async () => {
    globalThis.document = { createElement: tag => new Elemento(tag) };
    const container = new Elemento("div");
    const miniaturas = new PDFMiniaturas(container, () => {});
    const executar = miniaturas.processarFila.bind(miniaturas);
    let trocar = true;
    miniaturas.processarFila = () => {
        const fim = executar();
        fim.then(() => {
            if (!trocar) return;
            trocar = false;
            container.scrollTop = 49 * 166;
            miniaturas.atualizarJanela();
        });
        return fim;
    };
    try {
        miniaturas.definirDocumento(documento(100).pdf);
        await miniaturas.processamento;
        await aguardar(() => [...miniaturas.itens.values()].every(item => item.estado === "pronto"));
        assert.ok([...miniaturas.itens.values()].every(item => item.estado === "pronto"));
    } finally {
        await miniaturas.destruir();
        delete globalThis.document;
    }
});

test("regressão: arraste rola pela borda visível quando o painel ultrapassa o viewport", async () => {
    globalThis.document = { createElement: tag => new Elemento(tag) };
    globalThis.innerHeight = 800;
    const quadros = new Map();
    let id = 0;
    globalThis.requestAnimationFrame = callback => { quadros.set(++id, callback); return id; };
    globalThis.cancelAnimationFrame = atual => quadros.delete(atual);
    const container = new Elemento("div");
    container.clientHeight = 600;
    container.scrollHeight = 100 * 166;
    container.getBoundingClientRect = () => ({ top: 400, bottom: 1000 });
    const miniaturas = new PDFMiniaturas(container, () => {});
    try {
        miniaturas.definirDocumento(documento(100).pdf);
        await miniaturas.processamento;
        const origem = container.querySelector("button");
        const dataTransfer = { setData() {} };
        origem.emitir("dragstart", origem, { dataTransfer });
        container.emitir("dragover", container, { dataTransfer, clientY: 795 });
        const [atual, callback] = [...quadros][0];
        quadros.delete(atual);
        callback();
        assert.ok(container.scrollTop > 0);
    } finally {
        await miniaturas.destruir();
        delete globalThis.document;
        delete globalThis.innerHeight;
        delete globalThis.requestAnimationFrame;
        delete globalThis.cancelAnimationFrame;
    }
});

test("falha isolada tem erro e retry, mantendo as outras miniaturas e a navegação", async () => {
    globalThis.document = { createElement: tag => new Elemento(tag) };
    const container = new Elemento("div");
    const selecoes = [];
    const miniaturas = new PDFMiniaturas(container, numero => selecoes.push(numero));
    const normal = documento(3).pdf;
    let falhar = true;
    const pdf = { numPages: 3, getPage: async numero => {
        if (numero === 2 && falhar) { falhar = false; throw new Error("Falha simulada"); }
        return normal.getPage(numero);
    } };
    try {
        miniaturas.definirDocumento(pdf);
        await miniaturas.processamento;
        assert.deepEqual([...miniaturas.itens.values()].map(i => i.estado), ["pronto", "erro", "pronto"]);
        const item = miniaturas.itens.get(2);
        assert.equal(item.botao.getAttribute("aria-busy"), "false");
        assert.match(item.imagem.textContent, /Tentar novamente/);
        item.botao.click();
        await miniaturas.processamento;
        assert.equal(item.estado, "pronto");
        assert.deepEqual(selecoes, [2]);
    } finally {
        await miniaturas.destruir();
        delete globalThis.document;
    }
});

test("timeout de uma prévia não bloqueia a seguinte e permite nova tentativa", async () => {
    globalThis.document = { createElement: tag => new Elemento(tag) };
    const container = new Elemento("div");
    const miniaturas = new PDFMiniaturas(container, () => {}, () => {}, { tempoLimite: 20 });
    const normal = documento(2).pdf;
    let liberar;
    let primeira = true;
    const espera = new Promise(resolve => { liberar = resolve; });
    const pdf = { numPages: 2, getPage: async numero => {
        const pagina = await normal.getPage(numero);
        if (numero === 1 && primeira) {
            primeira = false;
            pagina.render = () => ({ promise: espera, cancel() {} });
        }
        return pagina;
    } };
    try {
        miniaturas.definirDocumento(pdf);
        await miniaturas.processamento;
        assert.equal(miniaturas.itens.get(1).estado, "erro");
        assert.equal(miniaturas.itens.get(2).estado, "pronto");
        miniaturas.tentarNovamente(1);
        await miniaturas.processamento;
        assert.equal(miniaturas.itens.get(1).estado, "pronto");
        liberar();
        await pausa(2);
        assert.equal(container.querySelectorAll("canvas").length, 2);
    } finally {
        liberar();
        await miniaturas.destruir();
        delete globalThis.document;
    }
});

test("getPage antigo que termina depois da substituição não publica canvas no novo documento", async () => {
    globalThis.document = { createElement: tag => new Elemento(tag) };
    const container = new Elemento("div");
    const miniaturas = new PDFMiniaturas(container, () => {});
    let liberar;
    const espera = new Promise(resolve => { liberar = resolve; });
    try {
        miniaturas.definirDocumento({ numPages: 1, getPage: () => espera });
        await Promise.resolve();
        await miniaturas.limpar();
        miniaturas.definirDocumento(documento(2).pdf);
        await miniaturas.processamento;
        liberar(await documento(1, 90).pdf.getPage(1));
        await pausa(2);
        assert.equal(container.querySelectorAll("canvas").length, 2);
        assert.ok(container.querySelectorAll("canvas").every(c => c.height > c.width));
    } finally {
        await miniaturas.destruir();
        delete globalThis.document;
    }
});

test("prévia principal atrasada não desfaz a rolagem manual ou a posição durante arraste", async () => {
    globalThis.document = { createElement: tag => new Elemento(tag) };
    const container = new Elemento("div");
    const miniaturas = new PDFMiniaturas(container, () => {});
    try {
        miniaturas.definirDocumento(documento(100).pdf);
        await miniaturas.processamento;
        const versaoRolagem = miniaturas.versaoRolagem;
        container.scrollTop = 49 * 166;
        container.emitir("scroll");
        miniaturas.selecionar(1, { versaoRolagem });
        assert.equal(container.scrollTop, 49 * 166);
        miniaturas.arraste.origem = 50;
        miniaturas.selecionar(1);
        assert.equal(container.scrollTop, 49 * 166);
    } finally {
        await miniaturas.destruir();
        delete globalThis.document;
    }
});

test("cache reutiliza canvases pela identidade após mover e invalida apenas a página girada", async () => {
    globalThis.document = { createElement: tag => new Elemento(tag) };
    const container = new Elemento("div");
    const miniaturas = new PDFMiniaturas(container, () => {});
    const { pdf, renderizadas } = documento(3);
    let chaves = ["a:0", "b:0", "c:0"];
    pdf.obterChaveMiniatura = numero => chaves[numero - 1];
    try {
        miniaturas.definirDocumento(pdf);
        await miniaturas.processamento;
        const a = miniaturas.itens.get(1).canvas;
        const b = miniaturas.itens.get(2).canvas;
        chaves = ["c:0", "a:0", "b:0"];
        await miniaturas.atualizarDocumento(pdf, 2);
        await miniaturas.processamento;
        assert.equal(renderizadas.length, 3);
        assert.equal(miniaturas.itens.get(2).canvas, a);
        assert.equal(miniaturas.itens.get(3).canvas, b);
        chaves = ["c:0", "a:90", "b:0"];
        await miniaturas.atualizarDocumento(pdf, 2);
        await miniaturas.processamento;
        assert.equal(renderizadas.length, 4);
        assert.equal(a.width, 0);
        assert.equal(miniaturas.itens.get(3).canvas, b);
    } finally {
        await miniaturas.destruir();
        assert.equal(miniaturas.cache.size, 0);
        delete globalThis.document;
    }
});

test("cache LRU permanece limitado ao percorrer 350 páginas e libera todos os canvases", async () => {
    globalThis.document = { createElement: tag => new Elemento(tag) };
    const container = new Elemento("div");
    const miniaturas = new PDFMiniaturas(container, () => {});
    const { pdf, renderizadas } = documento(350);
    try {
        miniaturas.definirDocumento(pdf);
        await miniaturas.processamento;
        const canvasInicial = renderizadas[0].canvas;
        for (const posicao of [40, 80, 120, 160, 200, 300]) {
            container.scrollTop = posicao * 166;
            container.emitir("scroll");
            await miniaturas.processamento;
            assert.ok(miniaturas.cache.size <= 32);
            assert.ok(container.querySelectorAll("canvas").length <= 16);
        }
        assert.equal(canvasInicial.width, 0, "A prévia mais antiga foi expulsa do cache");
        await miniaturas.destruir();
        assert.ok(renderizadas.every(r => r.canvas.width === 0));
    } finally {
        await miniaturas.destruir();
        delete globalThis.document;
    }
});
