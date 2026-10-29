import assert from "node:assert/strict";
import test from "node:test";
import { PDFDocument, PDFName, PDFString, degrees, rgb, cmyk } from "../public/vendor/pdf-lib/pdf-lib.js";
import { PDFWorkspace, calcularTransformacao, mmParaPontos as pt } from "./helpers/pdf-resize-modules.mjs";

async function abrir(tamanhos = [[210, 297]]) {
    const pdf = await PDFDocument.create();
    for (const [w, h, r = 0] of tamanhos) {
        const p = pdf.addPage([pt(w), pt(h)]); p.setRotation(degrees(r));
        p.drawText("Texto vetorial", { x: 10, y: 20, size: 12, color: cmyk(0, 1, 0, 0) });
        p.drawRectangle({ x: 2, y: 2, width: 30, height: 40, color: rgb(1, 0, 0) });
    }
    const workspace = new PDFWorkspace();
    const bytes = await pdf.save();
    await workspace.abrir({ name: "original.pdf", arrayBuffer: async () => bytes.slice().buffer });
    return workspace;
}
const perto = (a, b) => assert.ok(Math.abs(a - b) < 1e-6, `${a} != ${b}`);
for (const [origem, destino] of [[[210, 297], [148, 210]], [[148, 210], [210, 297]], [[297, 420], [210, 297]], [[210, 297], [123.45, 234.56]], [[210, 297], [297, 210]]]) {
    test(`redimensionar ${origem.join('×')} para ${destino.join('×')}: prévia transacional e aplicação`, async () => {
        const workspace = await abrir([origem]);
        const estado = workspace.capturarEstado();
        const previo = await workspace.prepararRedimensionamento({ largura: destino[0], altura: destino[1], modo: "ajustar", paginas: [1] });
        assert.deepEqual(workspace.capturarEstado(), estado);
        assert.equal(workspace.alterado, false);
        workspace.aplicarRedimensionamento(previo);
        const exportado = await PDFDocument.load(await workspace.obterBytes());
        perto(exportado.getPage(0).getWidth(), pt(destino[0]));
        perto(exportado.getPage(0).getHeight(), pt(destino[1]));
        assert.equal(workspace.alterado, true);
        assert.throws(() => workspace.aplicarRedimensionamento(previo), /vencida/);
        const repetida = await workspace.prepararRedimensionamento({ largura: destino[0], altura: destino[1], modo: "ajustar", paginas: [1] });
        assert.equal(workspace.aplicarRedimensionamento(repetida), false, "Não duplicar uma transformação sem efeito");
    });
}
test("modos: ajuste proporcional, escala 100%, preenchimento com corte e rotações", () => {
    const crop = { x: 20, y: 30, width: 400, height: 600 };
    const fit = calcularTransformacao(crop, 0, 1, 100, 100, "ajustar");
    assert.equal(fit.corte, false);
    perto(fit.escala, pt(100) / 600);
    const centro = calcularTransformacao(crop, 0, 1, 100, 100, "centralizar");
    assert.equal(centro.escala, 1); assert.equal(centro.corte, true);
    const fill = calcularTransformacao(crop, 0, 1, 100, 100, "preencher");
    perto(fill.escala, pt(100) / 400); assert.equal(fill.corte, true);
    const girada = calcularTransformacao(crop, 90, 2, 100, 100, "ajustar");
    perto(girada.original.largura, 1200 * 25.4 / 72);
    assert.equal(girada.matriz[0], 0); assert.ok(girada.matriz[1] < 0);
});
test("página atual, intervalo, todas; páginas mistas e já giradas", async () => {
    const workspace = await abrir([[210, 297, 90], [148, 210, 180], [297, 420, 270]]);
    for (const paginas of [[2], [1, 2], [1, 2, 3]]) {
        const preparo = await workspace.prepararRedimensionamento({ largura: 150, altura: 200, modo: "ajustar", paginas });
        workspace.aplicarRedimensionamento(preparo);
        for (const n of paginas) {
            perto(workspace.pdf.getPage(n - 1).getWidth(), pt(150));
            assert.equal(workspace.pdf.getPage(n - 1).getRotation().angle, 0);
        }
    }
});
test("preserva Media/Crop e transforma caixas existentes sem criar sangria", async () => {
    const workspace = await abrir();
    const p = workspace.pdf.getPage(0);
    p.setMediaBox(10, 20, 400, 600); p.setCropBox(20, 30, 350, 500);
    p.setTrimBox(30, 40, 300, 450); p.setBleedBox(25, 35, 320, 470); p.setArtBox(40, 50, 200, 300);
    const preparo = await workspace.prepararRedimensionamento({ largura: 100, altura: 150, modo: "ajustar", paginas: [1] });
    const transformada = preparo.pdf.getPage(0);
    assert.equal(transformada.getMediaBox().x, 0);
    assert.deepEqual(transformada.getMediaBox(), transformada.getCropBox());
    assert.ok(transformada.getBleedBox().width > transformada.getTrimBox().width);
    assert.match(preparo.avisos.join(" "), /Sangria não/);
    const simples = await abrir();
    const copia = await simples.prepararRedimensionamento({ largura: 148, altura: 210, modo: "ajustar", paginas: [1] });
    assert.equal(copia.pdf.getPage(0).node.has(PDFName.of("BleedBox")), false);
});
test("links URI e notas simples mudam posição; recursos e streams originais permanecem", async () => {
    const workspace = await abrir([[200, 300, 90]]), pdf = workspace.pdf, p = pdf.getPage(0);
    const link = pdf.context.obj({ Type: "Annot", Subtype: "Link", Rect: [10, 10, 40, 40], A: { S: "URI", URI: PDFString.of("https://example.com") }, P: p.ref });
    p.node.addAnnot(pdf.context.register(link));
    p.node.addAnnot(pdf.context.register(pdf.context.obj({ Type: "Annot", Subtype: "Text", Rect: [50, 50, 70, 70], Contents: PDFString.of("Nota"), P: p.ref })));
    const refs = p.node.Contents().asArray().map(r => r.toString());
    const preparo = await workspace.prepararRedimensionamento({ largura: 100, altura: 150, modo: "ajustar", paginas: [1] });
    const novo = preparo.pdf.getPage(0), annots = novo.node.Annots();
    assert.equal(annots.size(), 2);
    assert.notEqual(annots.lookup(0).lookup(PDFName.of("Rect")).toString(), link.lookup(PDFName.of("Rect")).toString());
    assert.equal(annots.lookup(0).lookup(PDFName.of("A")).lookup(PDFName.of("URI")).decodeText(), "https://example.com");
    for (const ref of refs) assert.ok(novo.node.Contents().asArray().some(r => r.toString() === ref));
    assert.equal(novo.node.Resources().toString(), p.node.Resources().toString());
});
test("erro em outra página, recursos não suportados, cancelamento e versão vencida são atômicos", async () => {
    const workspace = await abrir([[210, 297], [210, 297]]);
    const p = workspace.pdf.getPage(1);
    p.node.addAnnot(workspace.pdf.context.register(workspace.pdf.context.obj({ Type: "Annot", Subtype: "Widget", Rect: [0, 0, 10, 10] })));
    const estado = workspace.capturarEstado();
    await assert.rejects(workspace.prepararRedimensionamento({ largura: 100, altura: 150, modo: "ajustar", paginas: [1, 2] }), /Widget/);
    assert.deepEqual(workspace.capturarEstado(), estado);
    const controller = new AbortController(); controller.abort();
    await assert.rejects(workspace.prepararRedimensionamento({ largura: 100, altura: 150, modo: "ajustar", paginas: [1] }, { signal: controller.signal }), { name: "AbortError" });
    const preparado = await workspace.prepararRedimensionamento({ largura: 100, altura: 150, modo: "ajustar", paginas: [1] });
    workspace.girarPagina(1, 90);
    assert.throws(() => workspace.aplicarRedimensionamento(preparado), /vencida/);
});
test("redimensiona após unir, excluir e reorganizar; exporta a sequência e rotação restantes", async () => {
    const workspace = await abrir([[210, 297], [148, 210]]), outro = await abrir([[100, 200]]);
    await workspace.adicionarArquivos([{ name: "novo.pdf", arrayBuffer: () => outro.obterBytes() }]);
    workspace.excluirPaginas([2]); await workspace.moverPagina(2, 1); workspace.girarPagina(2, 90);
    const preparado = await workspace.prepararRedimensionamento({ largura: "123,45", altura: "234,56", modo: "centralizar", paginas: [1] });
    workspace.aplicarRedimensionamento(preparado);
    const pdf = await PDFDocument.load(await workspace.obterBytes());
    assert.equal(pdf.getPageCount(), 2); perto(pdf.getPage(0).getWidth(), pt(123.45));
    assert.equal(pdf.getPage(1).getRotation().angle, 90);
});

test("UserUnit/origem negativa e links cortados continuam válidos em redimensionamentos posteriores", async () => {
    const workspace = await abrir(), pdf = workspace.pdf, p = pdf.getPage(0);
    p.setMediaBox(-20, -30, 400, 600); p.setCropBox(-10, -20, 380, 580);
    p.node.set(PDFName.of("UserUnit"), pdf.context.obj(2));
    p.node.addAnnot(pdf.context.register(pdf.context.obj({ Type: "Annot", Subtype: "Link", Rect: [350, 500, 370, 550], A: { S: "URI", URI: PDFString.of("https://example.com") } })));
    const primeiro = await workspace.prepararRedimensionamento({ largura: 100, altura: 100, modo: "centralizar", paginas: [1] });
    perto(primeiro.detalhes[0].original.largura, 760 * 25.4 / 72);
    workspace.aplicarRedimensionamento(primeiro);
    const segundo = await workspace.prepararRedimensionamento({ largura: 150, altura: 150, modo: "ajustar", paginas: [1] });
    workspace.aplicarRedimensionamento(segundo);
    assert.equal(workspace.pdf.getPage(0).node.lookup(PDFName.of("UserUnit")).asNumber(), 1);
});

test("destinos internos, XFA, assinaturas e aparências complexas bloqueiam sem publicar alteração", async () => {
    for (const recurso of ["destino", "XFA", "SigFlags", "AP"]) {
        const workspace = await abrir(), pdf = workspace.pdf, p = pdf.getPage(0);
        if (recurso === "XFA" || recurso === "SigFlags") pdf.catalog.set(PDFName.of("AcroForm"), pdf.context.obj({ [recurso]: recurso === "SigFlags" ? 1 : "dados" }));
        else {
            const anotacao = pdf.context.obj({ Type: "Annot", Subtype: "Link", Rect: [0, 0, 20, 20],
                ...(recurso === "AP" ? { AP: {} } : { Dest: [p.ref, "Fit"] }) });
            p.node.addAnnot(pdf.context.register(anotacao));
        }
        const estado = workspace.capturarEstado();
        await assert.rejects(workspace.prepararRedimensionamento({ largura: 100, altura: 100, modo: "ajustar", paginas: [1] }));
        assert.deepEqual(workspace.capturarEstado(), estado);
        assert.equal(workspace.alterado, false);
    }
});
