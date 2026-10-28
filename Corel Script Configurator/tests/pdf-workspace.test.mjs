import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";
import { PDFDocument, degrees } from "../public/vendor/pdf-lib/pdf-lib.js";

// Resolve apenas o caminho HTTP da biblioteca para executar o módulo real no Node.
const biblioteca = new URL("../public/vendor/pdf-lib/pdf-lib.js", import.meta.url).href;
const codigo = (await readFile(new URL("../public/js/tools/pdf-workspace.js", import.meta.url), "utf8"))
    .replace('"/vendor/pdf-lib/pdf-lib.js"', JSON.stringify(biblioteca));
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
