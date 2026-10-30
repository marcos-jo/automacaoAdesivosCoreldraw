import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";
import { PDFDocument, PDFName, PDFString, PDFPageTree, degrees } from "../public/vendor/pdf-lib/pdf-lib.js";

// Resolve apenas o caminho HTTP da biblioteca para executar o módulo real no Node.
const biblioteca = new URL("../public/vendor/pdf-lib/pdf-lib.js", import.meta.url).href;
const codigo = (await readFile(new URL("../public/js/tools/pdf-workspace.js", import.meta.url), "utf8"))
    .replace('"/vendor/pdf-lib/pdf-lib.js"', JSON.stringify(biblioteca))
    .replace('"./pdf-resize-engine.js"', JSON.stringify((await import("./helpers/pdf-resize-modules.mjs")).motorUrl))
    .replace('"./pdf-signatures.js"', JSON.stringify((await import("./helpers/pdf-resize-modules.mjs")).assinaturasUrl))
    .replace('"./pdf-print-copy.js"', JSON.stringify((await import("./helpers/pdf-resize-modules.mjs")).copiaImpressaoUrl));
const { PDFWorkspace } = await import("data:text/javascript;base64," + Buffer.from(codigo).toString("base64"));

async function arquivo(nome, tamanhos) {
    const pdf = await PDFDocument.create();
    for (const [largura, altura, giro = 0] of tamanhos) {
        pdf.addPage([largura, altura]).setRotation(degrees(giro));
    }
    const bytes = await pdf.save();
    return { name: nome, bytes, arrayBuffer: async () => bytes.slice().buffer };
}

function resumo(pdf) {
    return pdf.getPages().map(pagina => [
        pagina.getWidth(), pagina.getHeight(), pagina.getRotation().angle
    ]);
}

test("inicia documento na ordem selecionada e exporta todas as páginas", async () => {
    const a = await arquivo("a.pdf", [[200, 300], [210, 310, 90]]);
    const b = await arquivo("b.pdf", [[400, 500]]);
    const originais = [a.bytes.slice(), b.bytes.slice()];
    const workspace = new PDFWorkspace();

    assert.equal(await workspace.adicionarArquivos([b, a], "combinado.pdf"), 3);
    assert.equal(workspace.nomeArquivo, "combinado.pdf");
    assert.equal(workspace.alterado, true);
    const exportado = await PDFDocument.load(await workspace.obterBytes());
    assert.deepEqual(resumo(exportado), [[400, 500, 0], [200, 300, 0], [210, 310, 90]]);
    assert.deepEqual([a.bytes, b.bytes], originais);
});

test("acrescenta ao PDF aberto preservando rotação, exclusão, nome e metadados", async () => {
    const base = await arquivo("cliente.pdf", [[100, 200], [110, 210], [120, 220]]);
    const novo = await arquivo("novo.pdf", [[300, 400, 270]]);
    const workspace = new PDFWorkspace();
    await workspace.abrir(base);
    workspace.pdf.setTitle("Trabalho do cliente");
    workspace.girarPagina(1, 90);
    workspace.excluirPaginas([2]);

    assert.equal(await workspace.adicionarArquivos([novo], "outro.pdf"), 1);
    assert.equal(workspace.nomeArquivo, "cliente.pdf");
    const exportado = await PDFDocument.load(await workspace.obterBytes());
    assert.deepEqual(resumo(exportado), [[100, 200, 90], [120, 220, 0], [300, 400, 270]]);
    assert.equal(exportado.getTitle(), "Trabalho do cliente");

    // Continua editável depois da união, sem duplicar o giro da primeira página.
    workspace.girarPagina(3, 90);
    workspace.excluirPaginas([2]);
    const editado = await PDFDocument.load(await workspace.obterBytes());
    assert.deepEqual(resumo(editado), [[100, 200, 90], [300, 400, 0]]);
    assert.equal((await PDFDocument.load(base.bytes)).getPageCount(), 3);
});

test("falha no meio do lote preserva integralmente o documento anterior", async () => {
    const base = await arquivo("base.pdf", [[100, 200], [120, 220]]);
    const valido = await arquivo("valido.pdf", [[300, 400]]);
    const invalido = { name: "invalido.pdf", arrayBuffer: async () => new ArrayBuffer(0) };

    for (const comDocumento of [false, true]) {
        const workspace = new PDFWorkspace();
        if (comDocumento) {
            await workspace.abrir(base);
            workspace.girarPagina(1, 90);
            workspace.excluirPaginas([2]);
        }
        const anterior = { pdf: workspace.pdf, nome: workspace.nomeArquivo, alterado: workspace.alterado };
        await assert.rejects(workspace.adicionarArquivos([valido, invalido]), /invalido\.pdf/);
        assert.equal(workspace.pdf, anterior.pdf);
        assert.equal(workspace.nomeArquivo, anterior.nome);
        assert.equal(workspace.alterado, anterior.alterado);
        assert.equal(workspace.getTotalPaginas(), comDocumento ? 1 : 0);
    }
});

test("rejeita lista vazia sem criar ou modificar documento", async () => {
    const workspace = new PDFWorkspace();
    await assert.rejects(workspace.adicionarArquivos([]));
    assert.equal(workspace.getTotalPaginas(), 0);
    assert.equal(workspace.alterado, false);
});

test("visualização é limpa; exportação quita só a versão atual e nova edição volta a ficar pendente", async () => {
    const workspace = new PDFWorkspace();
    await workspace.abrir(await arquivo("base.pdf", [[100, 200], [120, 220]]));
    assert.equal(workspace.alterado, false);
    workspace.girarPagina(1, 90);
    const versao = workspace.capturarVersao();
    await workspace.obterBytes();
    assert.equal(workspace.alterado, true, "Serializar para a prévia não é exportar");
    assert.equal(workspace.marcarExportado(versao), true);
    assert.equal(workspace.alterado, false);
    workspace.excluirPaginas([2]);
    assert.equal(workspace.alterado, true);
    assert.equal(workspace.marcarExportado(versao), false);
    assert.equal(workspace.alterado, true);
});

test("união inicial é pendente e exportação assíncrona não quita edições posteriores", async () => {
    const workspace = new PDFWorkspace();
    const base = await arquivo("base.pdf", [[100, 200]]);
    await workspace.adicionarArquivos([base]);
    assert.equal(workspace.alterado, true);
    const versao = workspace.capturarVersao();
    let concluir;
    const pendente = new Promise(resolve => { concluir = resolve; });
    const exportacao = (async () => { await pendente; return workspace.marcarExportado(versao); })();
    workspace.girarPagina(1, 90);
    concluir();
    assert.equal(await exportacao, false);
    assert.equal(workspace.alterado, true);
    await workspace.abrir(base);
    workspace.girarPagina(1, 90);
    assert.equal(workspace.marcarExportado(versao), false, "Token de outro documento deve ser recusado");
    assert.equal(workspace.alterado, true);
});

test("restauração após falha preserva a revisão e o estado de exportação anteriores", async () => {
    const workspace = new PDFWorkspace();
    const base = await arquivo("base.pdf", [[100, 200]]);
    await workspace.abrir(base);
    workspace.girarPagina(1, 90);
    const anterior = workspace.capturarEstado();
    await workspace.abrir(base);
    assert.equal(workspace.alterado, false);
    workspace.restaurarEstado(anterior);
    assert.equal(workspace.alterado, true);
    assert.equal(workspace.pdf.getPage(0).getRotation().angle, 90);
});

test("reordena após rotação, exclusão e união, sem alterar os originais", async () => {
    const base = await arquivo("base.pdf", [[100, 200], [110, 210], [120, 220]]);
    const outra = await arquivo("outra.pdf", [[300, 400, 270]]);
    const workspace = new PDFWorkspace();
    await workspace.abrir(base);
    workspace.girarPagina(1, 90);
    workspace.excluirPaginas([2]);
    await workspace.adicionarArquivos([outra]);
    workspace.marcarExportado(workspace.capturarVersao());
    const revisao = workspace.revisao;
    assert.deepEqual(await workspace.moverPagina(3, 1), { alterada: true, ordem: [3, 1, 2] });
    assert.deepEqual(resumo(await PDFDocument.load(await workspace.obterBytes())),
        [[300, 400, 270], [100, 200, 90], [120, 220, 0]]);
    assert.equal(workspace.revisao, revisao + 1);
    assert.equal(workspace.alterado, true);
    assert.equal(workspace.nomeArquivo, "base.pdf");
    assert.deepEqual(resumo(await PDFDocument.load(base.bytes)), [[100, 200, 0], [110, 210, 0], [120, 220, 0]]);
});

test("ordem inválida e movimento sem efeito preservam documento, revisão e exportação", async () => {
    const workspace = new PDFWorkspace();
    await workspace.abrir(await arquivo("base.pdf", [[100, 200], [120, 220]]));
    const anterior = workspace.capturarEstado();
    for (const ordem of [[1], [1, 1], [0, 2], [2, 3], ["2", 1], [2, undefined], null]) {
        await assert.rejects(workspace.reordenarPaginas(ordem), /inválida/);
        assert.deepEqual(workspace.capturarEstado(), anterior);
    }
    for (const movimento of [[0, 1], [1, 3], [1.5, 2]]) {
        await assert.rejects(workspace.moverPagina(...movimento), /inválido/);
    }
    assert.equal(await workspace.reordenarPaginas([1, 2]), false);
    assert.equal((await workspace.moverPagina(1, 1)).alterada, false);
    assert.deepEqual(workspace.capturarEstado(), anterior);
});

test("350 páginas: ordem completa, nenhuma rasterização e revisão única", async () => {
    const workspace = new PDFWorkspace();
    await workspace.abrir(await arquivo("grande.pdf", Array.from({ length: 350 }, (_, i) => [200 + i, 600])));
    const ordem = Array.from({ length: 350 }, (_, i) => 350 - i);
    await workspace.reordenarPaginas(ordem);
    const pdf = await PDFDocument.load(await workspace.obterBytes());
    assert.deepEqual(pdf.getPages().map(p => p.getWidth()), ordem.map(n => 199 + n));
    assert.equal(pdf.getPageCount(), 350);
});

test("preserva referências de página, anotações, links, formulário e metadados", async () => {
    const pdf = await PDFDocument.create();
    const a = pdf.addPage([200, 300]);
    const b = pdf.addPage([400, 500]);
    const link = pdf.context.obj({ Type: "Annot", Subtype: "Link", Rect: [0, 0, 50, 50],
        Dest: [b.ref, PDFName.of("Fit")], P: a.ref });
    const nota = pdf.context.obj({ Type: "Annot", Subtype: "Text", Rect: [60, 0, 80, 30],
        Contents: PDFString.of("Nota de teste"), P: b.ref });
    a.node.addAnnot(pdf.context.register(link));
    b.node.addAnnot(pdf.context.register(nota));
    const campo = pdf.getForm().createTextField("cliente");
    campo.setText("4PrintShops");
    campo.addToPage(b, { x: 30, y: 100, width: 120, height: 30 });
    pdf.setTitle("Trabalho original");
    const bytes = await pdf.save();
    const workspace = new PDFWorkspace();
    await workspace.abrir({ name: "anotado.pdf", arrayBuffer: async () => bytes.slice().buffer });
    const referencias = workspace.pdf.getPages().map(p => p.ref.toString());
    await workspace.reordenarPaginas([2, 1]);
    const resultado = await PDFDocument.load(await workspace.obterBytes());
    assert.deepEqual(resultado.getPages().map(p => p.ref.toString()), referencias.reverse());
    assert.equal(resultado.getTitle(), "Trabalho original");
    assert.equal(resultado.getForm().getTextField("cliente").getText(), "4PrintShops");
    const annots = resultado.getPage(1).node.Annots();
    const linkFinal = resultado.context.lookup(annots.get(0));
    assert.equal(linkFinal.lookup(PDFName.of("Dest")).get(0).toString(), resultado.getPage(0).ref.toString());
    const anotacoes = resultado.getPage(0).node.Annots();
    const notaFinal = resultado.context.lookup(anotacoes.get(0));
    assert.equal(notaFinal.lookup(PDFName.of("Contents")).decodeText(), "Nota de teste");
    assert.equal(notaFinal.get(PDFName.of("P")).toString(), resultado.getPage(0).ref.toString());
});

test("preserva recursos, caixas e rotação herdados de um ramo da árvore de páginas", async () => {
    const pdf = await PDFDocument.create();
    const a = pdf.addPage([200, 300]);
    pdf.addPage([400, 500]);
    const root = pdf.catalog.Pages();
    const ramo = PDFPageTree.withContext(pdf.context, pdf.catalog.get(PDFName.of("Pages")));
    const ramoRef = pdf.context.register(ramo);
    for (const [nome, valor] of [["MediaBox", [0, 0, 230, 330]], ["CropBox", [10, 10, 200, 300]], ["Rotate", 90]]) {
        ramo.set(PDFName.of(nome), pdf.context.obj(valor));
        a.node.delete(PDFName.of(nome));
    }
    ramo.set(PDFName.of("Resources"), pdf.context.obj({ Teste: "herdado" }));
    a.node.delete(PDFName.of("Resources"));
    ramo.Kids().push(a.ref);
    ramo.set(PDFName.of("Count"), pdf.context.obj(1));
    root.Kids().set(0, ramoRef);
    a.node.setParent(ramoRef);
    const bytes = await pdf.save();
    const workspace = new PDFWorkspace();
    await workspace.abrir({ name: "heranca.pdf", arrayBuffer: async () => bytes.slice().buffer });
    await workspace.reordenarPaginas([2, 1]);
    const pagina = (await PDFDocument.load(await workspace.obterBytes())).getPage(1);
    assert.equal(pagina.getWidth(), 230);
    assert.equal(pagina.getRotation().angle, 90);
    assert.deepEqual(pagina.getCropBox(), { x: 10, y: 10, width: 190, height: 290 });
    assert.equal(pagina.node.Resources().get(PDFName.of("Teste")).toString(), "/herdado");
});

test("falha na cópia e edição concorrente não deixam uma ordem parcial no workspace", async () => {
    const workspace = new PDFWorkspace();
    await workspace.abrir(await arquivo("base.pdf", [[100, 200], [120, 220]]));
    const anterior = workspace.capturarEstado();
    const obterBytes = workspace.obterBytes.bind(workspace);
    workspace.obterBytes = async () => new Uint8Array();
    await assert.rejects(workspace.reordenarPaginas([2, 1]));
    assert.deepEqual(workspace.capturarEstado(), anterior);
    let liberar;
    const espera = new Promise(resolve => { liberar = resolve; });
    workspace.obterBytes = async () => { await espera; return obterBytes(); };
    const reordenacao = workspace.reordenarPaginas([2, 1]);
    workspace.girarPagina(1, 90);
    liberar();
    await assert.rejects(reordenacao, /mudou durante/);
    assert.deepEqual(resumo(workspace.pdf), [[100, 200, 90], [120, 220, 0]]);
    assert.equal(workspace.revisao, anterior.revisao + 1);
});

test("identidades e giro seguem a árvore real após exclusões sucessivas, sem índices em cache", async () => {
    const workspace = new PDFWorkspace();
    await workspace.abrir(await arquivo("base.pdf", [[100, 200], [200, 300], [300, 400]]));
    const ids = workspace.obterPaginas().map(p => p.id);
    workspace.excluirPaginas([1]);
    assert.deepEqual(workspace.obterPaginas().map(p => p.id), ids.slice(1));
    workspace.girarPagina(1, 90);
    assert.deepEqual(resumo(await PDFDocument.load(await workspace.obterBytes())), [[200, 300, 90], [300, 400, 0]]);
    workspace.excluirPaginas([1]);
    workspace.girarPagina(1, -90);
    assert.deepEqual(workspace.obterPaginas(), [{ id: ids[2], rotacao: 270 }]);
    assert.deepEqual(resumo(await PDFDocument.load(await workspace.obterBytes())), [[300, 400, 270]]);
});
