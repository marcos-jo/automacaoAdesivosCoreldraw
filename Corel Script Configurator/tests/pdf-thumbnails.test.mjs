import assert from "node:assert/strict";
import test from "node:test";
import { Elemento } from "./helpers/pdf-dom.mjs";
import { PDFMiniaturas } from "../public/js/tools/pdf-thumbnails.js";

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

test("350 páginas: renderização sob demanda limitada, sem duplicação na mesma janela", async () => {
    globalThis.document = { createElement: tag => new Elemento(tag) };
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
        assert.equal(renderizadas[0].canvas.width, 0, "Canvas fora da janela deve ser liberado");
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
        delete globalThis.document;
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
        await Promise.resolve();
        await Promise.resolve();
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
