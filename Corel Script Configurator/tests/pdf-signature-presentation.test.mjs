import assert from "node:assert/strict";
import test from "node:test";
import { deflateSync } from "node:zlib";
import { Elemento } from "./helpers/pdf-dom.mjs";
import { PDFWorkspace } from "./helpers/pdf-resize-modules.mjs";
import { criarPDFAssinado } from "./helpers/pdf-signature-fixture.mjs";
import { renderPDFInterface } from "../public/js/tools/pdf-interface.js";
import { PDFInformacoesDocumento } from "../public/js/tools/pdf-document-info.js";

const arquivo = (bytes, name = "origem.pdf") => ({ name, arrayBuffer: async () => bytes.slice().buffer });
function interfaceResumo(ws) {
    const app = new Elemento("main");
    globalThis.document = { createElement: tag => new Elemento(tag) };
    const raiz = renderPDFInterface(app);
    const buscar = id => raiz.querySelector("#pdf-" + id);
    const informacoes = new PDFInformacoesDocumento(seletor => raiz.querySelector(seletor));
    const atualizar = () => informacoes.atualizar(ws.pdf.getPages().map((p, i) => ({ numero: i + 1, largura: p.getWidth(), altura: p.getHeight() })), ws.nomeArquivo, ws.obterResumoAssinaturas());
    atualizar();
    return { buscar, atualizar };
}
async function preparar(ws) {
    // Callback apenas estrutural para exercitar os estados reais do workspace.
    // A comparação de aparência e o fluxo completo são verificados no Chrome.
    const proposta = await ws.prepararCopiaImpressao({ modo: "necessarias", dpi: 300, paginas: [] }, {
        rasterizar: async () => ({ dados: new Uint8Array(deflateSync(new Uint8Array([255, 255, 255]))), largura: 1, altura: 1, area: [0, 0, 300, 400], dpi: 300 })
    });
    proposta.verificacao = { compativel: true };
    return proposta;
}

test("banner limita páginas visíveis e mantém diagnóstico completo nos detalhes e Informações", async () => {
    const ws = new PDFWorkspace();
    await ws.abrir(arquivo(await criarPDFAssinado({ total: 12, paginas: [1, 3, 5, 7, 9, 12] })));
    try {
        const ui = interfaceResumo(ws), estado = ws.capturarEstado();
        assert.match(ui.buscar("assinaturas-resumo").textContent, /1, 3, 5 e mais 3/);
        assert.ok(ui.buscar("assinaturas-resumo").textContent.length < 200);
        assert.equal(ui.buscar("assinaturas-resolver").textContent, "Resolver");
        assert.equal(ui.buscar("assinaturas-expandido").hidden, true);
        ui.buscar("assinaturas-expandir").click();
        assert.equal(ui.buscar("assinaturas-expandir").getAttribute("aria-expanded"), "true");
        assert.equal(ui.buscar("assinaturas-expandido").textContent, ui.buscar("assinaturas-detalhes").textContent);
        assert.match(ui.buscar("assinaturas-expandido").textContent, /1, 3, 5, 7, 9, 12/);
        assert.match(ui.buscar("assinaturas-detalhes").textContent, /aparências AP reconhecidas: 6/);
        ui.atualizar();
        assert.equal(ui.buscar("assinaturas-expandido").hidden, false);
        ui.buscar("assinaturas-expandir").click();
        assert.equal(ui.buscar("assinaturas-expandido").hidden, true);
        assert.deepEqual(ws.capturarEstado(), estado, "consultar/recolher não edita o documento");
    } finally { delete globalThis.document; }
});

test("prévia não muda o banner; aplicar separa cópia atual do histórico e recolhe detalhes", async () => {
    const ws = new PDFWorkspace(); await ws.abrir(arquivo(await criarPDFAssinado()));
    try {
        const ui = interfaceResumo(ws);
        ui.buscar("assinaturas-expandir").click();
        const proposta = await preparar(ws); ui.atualizar();
        assert.equal(ui.buscar("assinaturas-resolver").textContent, "Resolver");
        assert.equal(ws.alterado, false);
        ws.aplicarCopiaImpressao(proposta); ui.atualizar();
        assert.equal(ui.buscar("assinaturas-titulo").textContent, "Cópia para impressão preparada");
        assert.match(ui.buscar("assinaturas-resumo").textContent, /^Sem assinatura ativa/);
        assert.doesNotMatch(ui.buscar("assinaturas-resumo").textContent, /origem|detectada|validade|Páginas/);
        assert.equal(ui.buscar("assinaturas-resolver").textContent, "Refazer preparo");
        assert.equal(ui.buscar("assinaturas-expandido").hidden, true);
        assert.equal(ui.buscar("assinaturas-expandir").getAttribute("aria-expanded"), "false");
        assert.match(ui.buscar("assinaturas-detalhes").textContent, /Origem da cópia de impressão/);
        assert.match(ui.buscar("assinaturas-detalhes").textContent, /presentes 1/);
        assert.equal(ws.alterado, true);
    } finally { delete globalThis.document; }
});

test("nova origem assinada reativa Resolver sem contar assinaturas já preparadas no banner", async () => {
    const ws = new PDFWorkspace(); await ws.abrir(arquivo(await criarPDFAssinado()));
    ws.aplicarCopiaImpressao(await preparar(ws));
    try {
        const ui = interfaceResumo(ws);
        ui.buscar("assinaturas-expandir").click();
        await ws.adicionarArquivos([arquivo(await criarPDFAssinado({ total: 1 }), "nova.pdf")]); ui.atualizar();
        assert.equal(ui.buscar("assinaturas-titulo").textContent, "Assinatura digital detectada");
        assert.match(ui.buscar("assinaturas-resumo").textContent, /^1 assinatura/);
        assert.equal(ui.buscar("assinaturas-resolver").textContent, "Resolver");
        assert.equal(ui.buscar("assinaturas-expandido").hidden, true);
        assert.match(ui.buscar("assinaturas-detalhes").textContent, /origem.pdf/);
        assert.match(ui.buscar("assinaturas-detalhes").textContent, /nova.pdf/);
        assert.equal(ws.obterResumoAssinaturas().presentes, 2);
    } finally { delete globalThis.document; }
});

test("nova importação recolhe histórico anterior e documento sem assinatura não apresenta pendência fictícia", async () => {
    const ws = new PDFWorkspace(); await ws.abrir(arquivo(await criarPDFAssinado()));
    try {
        const ui = interfaceResumo(ws);
        ui.buscar("assinaturas-expandir").click();
        await ws.abrir(arquivo(await criarPDFAssinado({ invisivel: true }), "invisivel.pdf")); ui.atualizar();
        assert.equal(ui.buscar("assinaturas-expandido").hidden, true);
        assert.match(ui.buscar("assinaturas-resumo").textContent, /Invisíveis: 1/);
        assert.doesNotMatch(ui.buscar("assinaturas-detalhes").textContent, /origem.pdf/);
        await ws.abrir(arquivo(await criarPDFAssinado({ paginas: [] }), "limpo.pdf")); ui.atualizar();
        assert.equal(ui.buscar("assinaturas-alerta").hidden, true);
        assert.match(ui.buscar("assinaturas-detalhes").textContent, /nenhuma assinatura estrutural identificada/i);
        assert.doesNotMatch(ui.buscar("assinaturas-detalhes").textContent, /preparo pendente|invisivel.pdf/);
    } finally { delete globalThis.document; }
});
