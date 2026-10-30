import assert from "node:assert/strict";
import test from "node:test";
import { PDFDocument, PDFName, PDFHexString } from "../public/vendor/pdf-lib/pdf-lib.js";
import * as pdfjs from "../public/vendor/pdfjs/pdf.js";
import { PDFWorkspace, impressaoUrl } from "./helpers/pdf-resize-modules.mjs";
import { indexarTextoMarcado, verificarRedundanciaActualText, validarCompatibilidadeImpressao } from "../public/js/tools/pdf-print-validation.js";
import { criarFixtureImpressao, textoImpressao } from "./helpers/pdf-print-fixture.mjs";
const { diagnosticarImpressao, validarActualTextImpressao, prepararParaImpressao } = await import(impressaoUrl);
pdfjs.GlobalWorkerOptions.workerSrc = new URL("../public/vendor/pdfjs/pdf.worker.js", import.meta.url).href;
Uint8Array.prototype.toHex ??= function () { return Buffer.from(this).toString("hex"); };
// PDF.js também é conferido no Chrome, com implementação nativa de sumPrecise.
Math.sumPrecise ??= valores => Array.from(valores).reduce((a, b) => a + b, 0);
const nome = PDFName.of;
const config = { largura: 210, altura: 297, modo: "ajustar", paginas: [1], modoImpressao: true };

async function compararTextos({ antes, depois, substitutos }) {
    const marcas = new Map();
    async function extrair(bytes, guardar) {
        const tarefa = pdfjs.getDocument({ data: bytes.slice() });
        try {
            const pdf = await tarefa.promise, textos = [];
            for (let n = 1; n <= pdf.numPages; n++) {
                const conteudo = await (await pdf.getPage(n)).getTextContent({ includeMarkedContent: true, disableNormalization: true });
                if (guardar) marcas.set(n, indexarTextoMarcado(conteudo));
                textos.push(conteudo.items.filter(i => typeof i.str === "string").map(i => i.str));
            }
            return textos;
        } finally { await tarefa.destroy(); }
    }
    assert.deepEqual(await extrair(antes, true), await extrair(depois, false));
    return verificarRedundanciaActualText(substitutos, marcas);
}

test("ActualText redundante com acentos/símbolos exige prova e confirmação, sem perder edições anteriores", async () => {
    const bytes = await criarFixtureImpressao({ actualText: textoImpressao }), workspace = new PDFWorkspace();
    await workspace.abrir({ name: "redundante.pdf", arrayBuffer: async () => bytes.slice().buffer });
    workspace.girarPagina(1, 90);
    const estado = workspace.capturarEstado();
    const diagnostico = diagnosticarImpressao(workspace.pdf, [1]);
    assert.equal(diagnostico.actualText.length, 1);
    assert.deepEqual(diagnostico.actualText[0].conteudos, [{ pagina: 1, mcid: 0 }]);
    assert.throws(() => prepararParaImpressao(workspace.pdf, diagnostico), /ActualText/);
    await assert.rejects(workspace.prepararRedimensionamento(config), /ActualText/);
    // Este callback verifica texto real; a renderização real é validada no Chrome.
    await assert.rejects(workspace.prepararRedimensionamento({ ...config, modoImpressao: false }, { validarImpressao: compararTextos }), e => e.diagnostico.podePreparar && !!e.diagnostico.validacaoActualText);
    assert.deepEqual(workspace.capturarEstado(), estado);
    const previa = await workspace.prepararRedimensionamento(config, { validarImpressao: compararTextos });
    assert.equal(previa.validacaoActualText.compativel, true);
    assert.deepEqual(workspace.capturarEstado(), estado);
    workspace.aplicarRedimensionamento(previa);
    assert.equal(workspace.alterado, true);
    assert.equal(workspace.pdf.catalog.has(nome("StructTreeRoot")), false);
    assert.equal(workspace.pdf.getPage(0).getRotation().angle, 0);
});

test("texto substituto essencial continua bloqueado mesmo que extração global antes/depois seja igual", async () => {
    const pdf = await PDFDocument.load(await criarFixtureImpressao({ actualText: "Caracteres que só existem na estrutura" }));
    const diagnostico = diagnosticarImpressao(pdf, [1]);
    await assert.rejects(validarActualTextImpressao(pdf, diagnostico, compararTextos), /não é redundante/);
    assert.equal(pdf.catalog.has(nome("StructTreeRoot")), true);
    assert.throws(() => prepararParaImpressao(pdf, diagnostico));
});

test("ActualText sobre vetor e mistura normal/estruturado não usam coincidência de texto em outra região", async () => {
    const pdf = await PDFDocument.load(await criarFixtureImpressao({ actualText: textoImpressao })), p = pdf.getPage(0), c = pdf.context;
    p.node.addContentStream(c.register(c.flateStream("/Span << /MCID 42 >> BDC 0 0 30 40 re f EMC")));
    const raiz = pdf.catalog.lookup(nome("StructTreeRoot"));
    const vetor = c.register(c.obj({ Type: "StructElem", S: "Span", P: pdf.catalog.get(nome("StructTreeRoot")), Pg: p.ref, K: [42], ActualText: PDFHexString.fromText(textoImpressao) }));
    raiz.lookup(nome("K")).push(vetor);
    const diagnostico = diagnosticarImpressao(pdf, [1]);
    assert.equal(diagnostico.actualText.length, 2);
    await assert.rejects(validarActualTextImpressao(pdf, diagnostico, compararTextos), /ActualText 2/);
});

test("escopo de MCID, Unicode e repetições são verificados de maneira conservadora", () => {
    const conteudo = { items: [
        { type: "beginMarkedContentProps", id: "p0R_mc0" }, { str: "ação Ω", hasEOL: true }, { type: "endMarkedContent" },
        { type: "beginMarkedContentProps", id: "p0R_mc1" }, { str: "outro" }, { type: "endMarkedContent" }
    ] };
    const registros = [{ texto: "ac\u0327a\u0303o Ω", conteudos: [{ pagina: 1, mcid: 0 }] }];
    const marcas = new Map([[1, indexarTextoMarcado(conteudo)]]);
    assert.equal(verificarRedundanciaActualText(registros, marcas).compativel, true);
    assert.equal(verificarRedundanciaActualText([{ ...registros[0], texto: "outro" }], marcas).compativel, false);
    conteudo.items.push({ type: "beginMarkedContentProps", id: "p0R_mc0" }, { str: "ação Ω" }, { type: "endMarkedContent" });
    assert.equal(verificarRedundanciaActualText(registros, new Map([[1, indexarTextoMarcado(conteudo)]])).compativel, false);
});

test("cancelamento, erro visual e falha de leitura fecham tarefas e liberam canvases; não autorizam preparação", async () => {
    for (const cenario of ["igual", "pixel", "texto", "cancelar", "erro", "limite"]) {
        const abort = new AbortController(), fechadas = [], canvases = [];
        const pagina = indice => ({
            getViewport: () => ({ width: cenario === "limite" ? 9000 : 2, height: 2, rotation: 0 }),
            getTextContent: async () => ({ items: [{ type: "beginMarkedContentProps", id: "p0R_mc0" }, { str: indice && cenario === "texto" ? "diferente" : "ação" }, { type: "endMarkedContent" }] }),
            render: () => {
                if (cenario === "cancelar") queueMicrotask(() => abort.abort());
                return { promise: cenario === "cancelar" ? new Promise(() => {}) : cenario === "erro" ? Promise.reject(new Error("Render falhou")) : Promise.resolve(), cancel() {} };
            }, cleanup() {}
        });
        let abertas = 0;
        const dependencias = {
            abrir: async () => { const id = ++abertas; return { id, documento: { numPages: 1, getPage: async () => pagina(id - 1) } }; },
            fechar: async c => { fechadas.push(c.id); },
            criarCanvas: () => {
                const indice = canvases.length, canvas = { getContext: () => ({ getImageData: () => ({ data: new Uint8ClampedArray([0, 0, indice && cenario === "pixel" ? 1 : 0, 255]) }) }) };
                canvases.push(canvas); return canvas;
            }
        };
        const tarefa = validarCompatibilidadeImpressao({ antes: new Uint8Array(), depois: new Uint8Array(), substitutos: [{ texto: "ação", conteudos: [{ pagina: 1, mcid: 0 }] }], signal: abort.signal }, dependencias);
        if (["cancelar", "erro", "limite"].includes(cenario)) await assert.rejects(tarefa);
        else assert.equal((await tarefa).compativel, cenario === "igual");
        assert.deepEqual(fechadas.sort(), [1, 2]);
        assert.ok(canvases.every(c => c.width === 0 && c.height === 0));
    }
});

test("prova é vinculada à cópia e erro do comparador não elimina os bloqueios", async () => {
    const pdf = await PDFDocument.load(await criarFixtureImpressao({ actualText: textoImpressao }));
    const diagnostico = diagnosticarImpressao(pdf, [1]);
    await validarActualTextImpressao(pdf, diagnostico, compararTextos);
    const outra = await PDFDocument.load(await pdf.save());
    assert.throws(() => prepararParaImpressao(outra, diagnostico), /sem prova/);
    const novo = diagnosticarImpressao(pdf, [1]);
    await assert.rejects(validarActualTextImpressao(pdf, novo, async () => { throw new Error("Falha de renderização"); }), /Falha de renderização/);
    assert.throws(() => prepararParaImpressao(pdf, novo));
});

test("MCIDs acompanham as referências depois de união, movimento e exclusão anteriores", async () => {
    const bytes = await criarFixtureImpressao({ actualText: textoImpressao }), workspace = new PDFWorkspace();
    await workspace.abrir({ name: "marcado.pdf", arrayBuffer: async () => bytes.slice().buffer });
    const outro = await PDFDocument.create(); outro.addPage([300, 400]); outro.addPage([400, 500]); const b = await outro.save();
    await workspace.adicionarArquivos([{ name: "outro.pdf", arrayBuffer: async () => b.slice().buffer }]);
    await workspace.moverPagina(1, 3); workspace.excluirPaginas([1]);
    assert.deepEqual(diagnosticarImpressao(workspace.pdf, [1, 2]).actualText[0].conteudos, [{ pagina: 2, mcid: 0 }]);
    const estado = workspace.capturarEstado();
    const previa = await workspace.prepararRedimensionamento({ ...config, paginas: [1, 2] }, { validarImpressao: compararTextos });
    assert.deepEqual(workspace.capturarEstado(), estado);
    workspace.aplicarRedimensionamento(previa);
    assert.equal(workspace.getTotalPaginas(), 2);
});

test("árvore compartilhada/cíclica e conteúdo de XObject não dispensam a proteção", async () => {
    for (const risco of ["compartilhada", "ciclica", "xobject"]) {
        const pdf = await PDFDocument.load(await criarFixtureImpressao({ actualText: textoImpressao })), c = pdf.context;
        const raiz = pdf.catalog.lookup(nome("StructTreeRoot")), filhos = raiz.lookup(nome("K")), span = c.lookup(filhos.get(0));
        if (risco === "compartilhada") filhos.push(filhos.get(0));
        else if (risco === "ciclica") span.set(nome("K"), filhos);
        else span.set(nome("K"), c.obj({ Type: "MCR", Pg: pdf.getPage(0).ref, MCID: 0, Stm: c.register(c.flateStream("")) }));
        const diagnostico = diagnosticarImpressao(pdf, [1]);
        assert.ok(diagnostico.actualText.some(r => r.erro));
        await assert.rejects(validarActualTextImpressao(pdf, diagnostico, compararTextos));
        assert.equal(pdf.catalog.has(nome("StructTreeRoot")), true);
    }
});
