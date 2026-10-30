import test from "node:test";
import assert from "node:assert/strict";
import { PDFWorkspace } from "./helpers/pdf-resize-modules.mjs";
import { PDFDocument, PDFName, degrees } from "../public/vendor/pdf-lib/pdf-lib.js";
import { criarPDFAssinado } from "./helpers/pdf-signature-fixture.mjs";
const nome = PDFName.of;
const arquivo = bytes => ({ name: 'copiar.pdf', arrayBuffer: async () => bytes.slice().buffer });
async function criar() {
    const p = await PDFDocument.create();
    for (let i = 0; i < 5; i++) { const pagina = p.addPage([200 + i, 300]); pagina.drawText('Pagina ' + i); pagina.setRotation(degrees(i % 2 * 90)); }
    return p;
}
test("duplica grupo após última selecionada com fontes, conteúdo, giro e identidades independentes", async () => {
    const pdf = await criar(), bytes = await pdf.save(), original = bytes.slice(), ws = new PDFWorkspace(); await ws.abrir(arquivo(bytes));
    const ids = ws.obterPaginas().map(p => p.id), r = ws.revisao;
    assert.deepEqual(await ws.duplicarPaginas([4, 2]), [5, 6]);
    assert.equal(ws.revisao, r + 1); assert.equal(ws.alterado, true);
    assert.equal(new Set(ws.obterPaginas().map(p => p.id)).size, 7);
    assert.deepEqual(ws.obterPaginas().slice(0, 4).map(p => p.id), ids.slice(0, 4));
    const reaberto = await PDFDocument.load(await ws.obterBytes());
    assert.deepEqual(reaberto.getPages().map(p => p.getWidth()), [200, 201, 202, 203, 201, 203, 204]);
    assert.deepEqual(reaberto.getPages().map(p => p.getRotation().angle), [0, 90, 0, 90, 90, 90, 0]);
    for (const n of [4, 5]) {
        assert.ok(reaberto.getPage(n).node.Resources().lookup(nome('Font')).keys().length);
        assert.ok(reaberto.getPage(n).node.Contents());
    }
    assert.deepEqual(bytes, original);
    await ws.moverPagina(6, 1); ws.excluirPaginas([3]);
    await ws.adicionarArquivos([arquivo(await (await criar()).save())]);
    const preparacao = await ws.prepararRedimensionamento({ largura:210, altura:297, modo:'ajustar',paginas:[1],paginaPrevia:1 });
    ws.aplicarRedimensionamento(preparacao);
    assert.equal((await PDFDocument.load(await ws.obterBytes())).getPageCount(), 11);
});
test("bloqueia estruturas interativas/assinaturas e seleções inválidas sem modificar workspace", async () => {
    for (const recurso of ['Annots', 'AcroForm', 'Outlines', 'Names', 'StructTreeRoot', 'PageLabels', 'PrintPageRange', 'PresSteps']) {
        const pdf = await criar();
        if (recurso === 'Annots') pdf.getPage(0).node.set(nome(recurso), pdf.context.obj([{ Type:'Annot',Subtype:'Link',Rect:[0,0,10,10] }]));
        else if (recurso === 'PresSteps') pdf.getPage(0).node.set(nome(recurso), pdf.context.obj({}));
        else pdf.catalog.set(nome(recurso), pdf.context.obj({}));
        const ws = new PDFWorkspace(); await ws.abrir(arquivo(await pdf.save())); const antes = ws.capturarEstado();
        await assert.rejects(ws.duplicarPaginas([1]), /bloqueada/); assert.deepEqual(ws.capturarEstado(), antes);
    }
    const ws = new PDFWorkspace(); await ws.abrir(arquivo(await criarPDFAssinado()));
    await assert.rejects(ws.duplicarPaginas([1]), /assinatura/);
    await assert.rejects(ws.duplicarPaginas([0]), /inválida/);
});
test("concorrência rejeita duplicação preparada sobre revisão antiga", async () => {
    const ws = new PDFWorkspace(); await ws.abrir(arquivo(await (await criar()).save()));
    const trabalho = ws.duplicarPaginas([2]); ws.girarPagina(1, 90);
    await assert.rejects(trabalho, /mudou/); assert.equal(ws.getTotalPaginas(), 5);
    assert.equal(ws.obterPaginas()[0].rotacao, 90);
});
