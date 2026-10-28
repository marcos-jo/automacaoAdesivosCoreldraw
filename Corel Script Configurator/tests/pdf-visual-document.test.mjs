import assert from "node:assert/strict";
import test from "node:test";
import { PDFDocumentoVisual } from "../public/js/tools/pdf-visual-document.js";

function contexto(total = 3) {
    let paginas = Array.from({ length: total }, (_, i) => ({ id: `ref-${i + 1}`, rotacao: i === 0 ? 90 : 0 }));
    const leituras = [], renders = [];
    const workspace = { obterPaginas: () => paginas, obterBytes: () => new Uint8Array([paginas.length, paginas[0].rotacao / 90]) };
    const fonte = { numPages: total, getPage: async numero => {
        leituras.push(numero);
        return { getViewport: ({ scale = 1, rotation = 0 }) => ({
            width: (rotation % 180 ? 600 : 400) * scale,
            height: (rotation % 180 ? 400 : 600) * scale,
            rotation
        }), render: opcoes => { renders.push({ numero, rotacao: opcoes.viewport.rotation }); return { promise: Promise.resolve() }; } };
    }, getData: () => assert.fail("Não usar os bytes da fonte antiga") };
    return { workspace, fonte, leituras, renders, paginas: () => paginas, definir: novas => { paginas = novas; } };
}

test("mapa mantém identidade e rotação absoluta após movimentos, sem reler dimensões", async () => {
    const c = contexto(350);
    let doc = await PDFDocumentoVisual.criar(c.fonte, c.workspace);
    assert.equal(c.leituras.length, 350);
    const chave = doc.obterChaveMiniatura(1);
    for (let i = 0; i < 8; i++) {
        const paginas = c.paginas().slice();
        paginas.splice(49, 0, ...paginas.splice(4, 1));
        paginas[49] = { ...paginas[49], rotacao: (i % 4) * 90 };
        c.definir(paginas);
        doc = doc.atualizar(c.workspace);
        assert.equal(doc.obterChaveMiniatura(1), chave);
        assert.equal(doc.obterDimensoes().length, 350);
    }
    assert.equal(c.leituras.length, 350, "Atualizações usam somente metadados já conhecidos");
    const pagina = await doc.getPage(50);
    await pagina.render({ viewport: pagina.getViewport({ scale: 1 }) }).promise;
    assert.equal(c.renders.at(-1).rotacao, 270);
    assert.equal(pagina.pageNumber, 50);
});

test("snapshot antigo não muda enquanto uma busca assíncrona termina", async () => {
    const c = contexto();
    const antigo = await PDFDocumentoVisual.criar(c.fonte, c.workspace);
    const busca = antigo.getPage(1);
    c.definir(c.paginas().slice().reverse());
    const novo = antigo.atualizar(c.workspace);
    const paginaAntiga = await busca;
    assert.equal(paginaAntiga.rotate, 90);
    assert.equal((await novo.getPage(1)).rotate, 0);
    assert.equal(antigo.obterChaveMiniatura(1), "ref-1:90");
});

test("extração recebe bytes atuais; exclusão e recarga conservam dimensões por identidade", async () => {
    const c = contexto();
    const antigo = await PDFDocumentoVisual.criar(c.fonte, c.workspace);
    c.definir(c.paginas().slice(1).reverse());
    const novo = await PDFDocumentoVisual.criar({ ...c.fonte, numPages: 2 }, c.workspace, antigo);
    assert.equal(c.leituras.length, 3);
    assert.deepEqual(await novo.getData(), new Uint8Array([2, 0]));
    assert.equal(novo.obterChaveMiniatura(1), "ref-3:0");
    await novo.getPage(1);
    assert.equal(c.leituras.at(-1), 1, "Recarga renova índices da fonte PDF.js");
});

test("página desconhecida exige caminho completo e não modifica mapa existente", async () => {
    const c = contexto();
    const antigo = await PDFDocumentoVisual.criar(c.fonte, c.workspace);
    c.definir([{ id: "nova", rotacao: 0 }]);
    assert.throws(() => antigo.atualizar(c.workspace), /recarregar/);
    assert.equal(antigo.numPages, 3);
    assert.equal(antigo.obterChaveMiniatura(1), "ref-1:90");
});
