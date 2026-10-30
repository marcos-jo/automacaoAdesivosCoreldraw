import assert from "node:assert/strict";
import test from "node:test";
import { PDFDocument, PDFName, PDFStream } from "../public/vendor/pdf-lib/pdf-lib.js";
import * as pdfjs from "../public/vendor/pdfjs/pdf.js";
import { PDFWorkspace, impressaoUrl, mmParaPontos } from "./helpers/pdf-resize-modules.mjs";
import { criarFixtureImpressao, marcarEstrutura, textoImpressao } from "./helpers/pdf-print-fixture.mjs";
const { diagnosticarImpressao, prepararParaImpressao } = await import(impressaoUrl);
pdfjs.GlobalWorkerOptions.workerSrc = new URL("../public/vendor/pdfjs/pdf.worker.js", import.meta.url).href;
Uint8Array.prototype.toHex ??= function () { return Buffer.from(this).toString("hex"); };
const nome = PDFName.of;
const config = { largura: 210, altura: 297, modo: "ajustar", paginas: [1], modoImpressao: true };
async function abrir(bytes) {
    const workspace = new PDFWorkspace();
    await workspace.abrir({ name: "powerpoint.pdf", arrayBuffer: async () => bytes.slice().buffer });
    return workspace;
}
async function textos(bytes) {
    const tarefa = pdfjs.getDocument({ data: bytes.slice() });
    try { const pdf = await tarefa.promise; return (await (await pdf.getPage(1)).getTextContent()).items.map(i => i.str).join(" "); }
    finally { await tarefa.destroy(); }
}
function streams(pdf) {
    return new Map(pdf.context.enumerateIndirectObjects().filter(([, o]) => o instanceof PDFStream).map(([r, o]) => [r.toString(), Buffer.from(o.getContents()).toString("base64")]));
}

test("diagnóstico separa marcação, marcadores, destinos e links; preserva streams e recursos ao preparar", async () => {
    const bytes = await criarFixtureImpressao({ navegacao: true, links: true });
    const pdf = await PDFDocument.load(bytes);
    const metadata = pdf.context.register(pdf.context.flateStream("<metadata>produção</metadata>", { Type: "Metadata", Subtype: "XML" }));
    pdf.catalog.set(nome("Metadata"), metadata);
    pdf.catalog.set(nome("OutputIntents"), pdf.context.obj([{ Type: "OutputIntent", S: "GTS_PDFX" }]));
    pdf.catalog.set(nome("OCProperties"), pdf.context.obj({ OCGs: [], D: { Order: [] } }));
    const preservados = ["Metadata", "OutputIntents", "OCProperties"].map(chave => [chave, pdf.catalog.get(nome(chave)).toString()]);
    const anterior = streams(pdf), recursos = pdf.getPage(0).node.Resources().toString();
    const diagnostico = diagnosticarImpressao(pdf, [1]);
    assert.equal(diagnostico.podePreparar, true);
    for (const codigo of ["acessibilidade", "marcadores", "destinos", "abertura", "links-internos"]) assert.ok(diagnostico.removiveis.some(r => r.codigo === codigo));
    prepararParaImpressao(pdf, diagnostico);
    assert.deepEqual(streams(pdf), anterior, "Todos os bytes de conteúdo, fontes e imagens preservados");
    assert.equal(pdf.getPage(0).node.Resources().toString(), recursos);
    for (const chave of ["StructTreeRoot", "MarkInfo", "Outlines", "Dests", "OpenAction", "PageMode"]) assert.equal(pdf.catalog.has(nome(chave)), false, chave);
    assert.equal(pdf.catalog.lookup(nome("Names")).has(nome("EmbeddedFiles")), true);
    assert.equal(pdf.catalog.lookup(nome("Names")).has(nome("Dests")), false);
    assert.equal(pdf.catalog.lookup(nome("Lang")).decodeText(), "pt-BR");
    for (const [chave, valor] of preservados) assert.equal(pdf.catalog.get(nome(chave)).toString(), valor, chave);
    assert.equal(pdf.getPage(0).node.has(nome("StructParents")), false);
    assert.equal(pdf.getPage(0).node.has(nome("Tabs")), false);
    const annots = pdf.getPage(0).node.Annots();
    assert.equal(annots.size(), 3);
    for (const i of [0, 1]) { assert.equal(annots.lookup(i).has(nome("Dest")), false); assert.equal(annots.lookup(i).has(nome("A")), false); }
    assert.equal(annots.lookup(2).lookup(nome("A")).lookup(nome("URI")).decodeText(), "https://example.com");
    assert.equal(annots.lookup(0).lookup(nome("Border")).toString(), "[ 0 0 2 ]");
    assert.equal(diagnosticarImpressao(pdf, [1]).removiveis.length, 0, "Nenhuma referência estrutural ativa restante");
    const textoAntes = await textos(bytes), textoDepois = await textos(await pdf.save());
    assert.ok(textoAntes.includes(textoImpressao));
    assert.equal(textoDepois, textoAntes, "Inclui símbolos e ActualText no conteúdo");
});

test("PowerPoint marcado exige consentimento; prévia/cancelamento preservam original e aplicar marca pendência", async () => {
    const bytes = await criarFixtureImpressao(), workspace = await abrir(bytes), estado = workspace.capturarEstado();
    await assert.rejects(workspace.prepararRedimensionamento({ ...config, modoImpressao: false }), e => e.diagnostico.podePreparar);
    const preparado = await workspace.prepararRedimensionamento(config);
    assert.deepEqual(workspace.capturarEstado(), estado);
    assert.equal(workspace.alterado, false);
    const textoAntes = await textos(bytes), textoDepois = await textos(await preparado.pdf.save());
    assert.equal(textoDepois, textoAntes);
    for (const [ref, dados] of streams(workspace.pdf)) assert.equal(streams(preparado.pdf).get(ref), dados, ref);
    assert.equal(workspace.pdf.catalog.has(nome("StructTreeRoot")), true);
    workspace.aplicarRedimensionamento(preparado); assert.equal(workspace.alterado, true);
    const exportado = await PDFDocument.load(await workspace.obterBytes());
    assert.ok(Math.abs(exportado.getPage(0).getWidth() - mmParaPontos(210)) < 1e-6);
    assert.equal(exportado.catalog.has(nome("StructTreeRoot")), false);
    assert.equal((await PDFDocument.load(bytes)).catalog.has(nome("StructTreeRoot")), true);
});

test("preparação sem mudança geométrica pode ser aplicada e exportada; consentimento não é persistido", async () => {
    const workspace = await abrir(await criarFixtureImpressao());
    const preparo = await workspace.prepararRedimensionamento({ ...config, largura: 960 * 25.4 / 72, altura: 540 * 25.4 / 72 });
    assert.equal(preparo.alteradas.length, 0); assert.equal(preparo.preparadoParaImpressao, true);
    assert.equal(workspace.aplicarRedimensionamento(preparo), true); assert.equal(workspace.alterado, true);
    assert.equal((await workspace.prepararRedimensionamento({ ...config, modoImpressao: false })).preparadoParaImpressao, false);
});

test("modos, formatos, 90°, dimensões mistas e edições anteriores mantêm identidade e texto", async () => {
    const bytes = await criarFixtureImpressao({ navegacao: true, links: true, rotacao: 90 });
    for (const modo of ["ajustar", "centralizar", "preencher"]) for (const [largura, altura] of [[210, 297], [297, 420], [148, 210], [105, 148], [180.5, 260.25]]) {
        const workspace = await abrir(bytes);
        const outro = await PDFDocument.create(); outro.addPage([200, 400]); outro.addPage([300, 300]);
        const b = await outro.save(); await workspace.adicionarArquivos([{ name: "outro.pdf", arrayBuffer: async () => b.slice().buffer }]);
        await workspace.moverPagina(1, 3); workspace.excluirPaginas([1]);
        const ids = workspace.obterPaginas().map(p => p.id);
        const preparado = await workspace.prepararRedimensionamento({ ...config, modo, largura, altura, paginas: [1, 2] });
        workspace.aplicarRedimensionamento(preparado);
        assert.deepEqual(workspace.obterPaginas().map(p => p.id), ids);
        assert.equal(workspace.pdf.getPage(1).getRotation().angle, 0);
        assert.ok(Math.abs(workspace.pdf.getPage(1).getWidth() - mmParaPontos(largura)) < 1e-6);
        assert.equal(diagnosticarImpressao(workspace.pdf, [1, 2]).removiveis.length, 0);
        // Com corte explícito o PDF.js pode omitir texto fora da MediaBox.
        if (modo === "ajustar") {
            const tarefa = pdfjs.getDocument({ data: (await workspace.obterBytes()).slice() });
            try {
                const pagina = await (await tarefa.promise).getPage(2);
                assert.ok((await pagina.getTextContent()).items.map(i => i.str).join(" ").includes(textoImpressao));
            } finally { await tarefa.destroy(); }
        }
    }
});

test("bloqueios visuais e ActualText na árvore continuam atômicos mesmo com modo de impressão", async () => {
    for (const risco of ["Widget", "Fields", "AP", "XFA", "SigFlags", "AA", "JavaScript", "ActualText"]) {
        const workspace = await abrir(await criarFixtureImpressao()), pdf = workspace.pdf, p = pdf.getPage(0), c = pdf.context;
        if (risco === "ActualText") marcarEstrutura(pdf, p, true);
        else if (risco === "Fields") pdf.catalog.set(nome("AcroForm"), c.obj({ Fields: [c.register(c.obj({ FT: "Tx", T: "campo" }))] }));
        else if (["XFA", "SigFlags"].includes(risco)) pdf.catalog.set(nome("AcroForm"), c.obj({ [risco]: risco === "XFA" ? "dados" : 1 }));
        else if (risco === "AA") p.node.set(nome("AA"), c.obj({ O: { S: "JavaScript" } }));
        else if (risco === "JavaScript") pdf.catalog.set(nome("Names"), c.obj({ JavaScript: { Names: [] } }));
        else p.node.addAnnot(c.register(c.obj({ Type: "Annot", Subtype: risco === "Widget" ? "Widget" : "Link", Rect: [0, 0, 20, 20], ...(risco === "AP" ? { AP: { N: c.register(c.flateStream("0 0 10 10 re f", { BBox: [0, 0, 10, 10] })) } } : {}) })));
        const estado = workspace.capturarEstado();
        await assert.rejects(workspace.prepararRedimensionamento(config), e => e.diagnostico.bloqueios.length > 0 && !e.diagnostico.podePreparar);
        assert.deepEqual(workspace.capturarEstado(), estado); assert.equal(workspace.alterado, false);
    }
});

test("PDF sem marcação com URI externa não precisa perder estruturas", async () => {
    const workspace = await abrir(await criarFixtureImpressao({ marcada: false }));
    const pdf = workspace.pdf, p = pdf.getPage(0);
    p.node.addAnnot(pdf.context.register(pdf.context.obj({ Type: "Annot", Subtype: "Link", Rect: [0, 0, 20, 20], A: { S: "URI" } })));
    const preparado = await workspace.prepararRedimensionamento({ ...config, modoImpressao: false });
    assert.deepEqual(preparado.removidos, []);
    assert.equal(preparado.preparadoParaImpressao, false);
});
