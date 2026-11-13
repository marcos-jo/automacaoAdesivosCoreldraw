import assert from "node:assert/strict";
import test from "node:test";
import { deflateSync } from "node:zlib";
import { PDFDocument, PDFName } from "../public/vendor/pdf-lib/pdf-lib.js";
import { PDFWorkspace, assinaturasUrl, copiaImpressaoUrl } from "./helpers/pdf-resize-modules.mjs";
import { estimarRaster } from "../public/js/tools/pdf-print-raster.js";
import { criarPDFAssinado as criarFixtureAssinada } from "./helpers/pdf-signature-fixture.mjs";
const { planejarCopiaImpressao, reconstruirCopiaImpressao } = await import(copiaImpressaoUrl);
const { verificarSemAssinaturas } = await import(assinaturasUrl);
const criarPDFAssinado = (opcoes={}) => criarFixtureAssinada({...opcoes,vazia:true});
const config = { modo: "necessarias", dpi: 300, paginas: [] }, nome = PDFName.of;
const arquivo = bytes => ({ name: "assinado.pdf", arrayBuffer: async () => bytes.slice().buffer });
// Imagem pequena para testes ESTRUTURAIS do engine. Fidelidade/DPI reais são
// testados no Chrome, não por este callback injetado.
const rasterizar = async () => ({ dados: new Uint8Array(deflateSync(new Uint8Array([255, 255, 255, 0, 0, 0, 255, 255, 255, 0, 0, 0]))), largura: 2, altura: 2, area: [0, 0, 300, 400], escala: .005, dpi: 300 });
test("plano seletivo inclui widgets vazios/selos imprimíveis, permite páginas adicionais e preserva vetoriais", async () => {
    const pdf = await PDFDocument.load(await criarPDFAssinado({ anotacao: true }));
    const p = planejarCopiaImpressao(pdf, config);
    assert.deepEqual(p.rasterizadas, [1, 3]); assert.deepEqual(p.vetoriais, [2]);
    assert.deepEqual(planejarCopiaImpressao(pdf, { ...config, paginas: [2] }).rasterizadas, [1, 2, 3]);
    assert.deepEqual(planejarCopiaImpressao(pdf, { ...config, modo: "inteiro" }).rasterizadas, [1, 2, 3]);
});
test("reconstrói contexto novo sem assinatura aplicada; mantém caixas/rotação e elimina campos vazios", async () => {
    const origem = await PDFDocument.load(await criarPDFAssinado({ rotacao: 90 }));
    const resultado = await reconstruirCopiaImpressao(origem, config, { rasterizar });
    verificarSemAssinaturas(resultado.pdf);
    const reaberto = await PDFDocument.load(await resultado.pdf.save()); verificarSemAssinaturas(reaberto);
    assert.equal(reaberto.getPageCount(), 3);
    assert.deepEqual(reaberto.getPages().map(p => [p.getWidth(), p.getHeight(), p.getRotation().angle]), [[300, 400, 90], [300, 400, 90], [400, 300, 0]]);
    assert.ok(reaberto.getPage(0).node.Resources().has(nome("Font")));
    assert.equal(reaberto.getPage(2).node.Resources().lookup(nome("Font"))?.keys().length || 0, 0);
    assert.ok(reaberto.getPages().every(p => !p.node.Annots()?.size()));
});
test("widget vazio invisível não inventa páginas; aparências ausentes e conteúdo dinâmico mantêm bloqueio", async () => {
    const invisivel = await PDFDocument.load(await criarPDFAssinado({ invisivel: true }));
    // Widgets sem área não são necessários; a seleção manual é exigida.
    assert.throws(() => planejarCopiaImpressao(invisivel, config), /páginas visuais|Nenhuma página|Informe páginas|Não há páginas/);
    assert.deepEqual(planejarCopiaImpressao(invisivel, { ...config, paginas: [2] }).rasterizadas, [2]);
    for (const problema of ["sem-ap", "sem-print", "xfa", "camadas"]) {
        const pdf = await PDFDocument.load(await criarPDFAssinado({ semAparencia: problema === "sem-ap", imprimivel: problema !== "sem-print" }));
        if (problema === "xfa") pdf.catalog.lookup(nome("AcroForm")).set(nome("XFA"), pdf.context.obj("dados"));
        if (problema === "camadas") pdf.catalog.set(nome("OCProperties"), pdf.context.obj({ OCGs: [] }));
        assert.throws(() => planejarCopiaImpressao(pdf, config));
        if (problema === "camadas") assert.equal(planejarCopiaImpressao(pdf, { ...config, modo: "inteiro" }).rasterizadas.length, 3);
    }
});
test("PDF-008: campos vazios continuam distintos de assinatura aplicada",async()=>{
 const ws=new PDFWorkspace();await ws.abrir(arquivo(await criarPDFAssinado()));assert.equal(ws.obterResumoAssinaturas().presentes,0);assert.equal(ws.obterResumoAssinaturas().vazios,1);ws.girarPagina(3,90);assert.equal(ws.alterado,true);
});


test("preparar/cancelar/falhar preservam estado; aplicar exige verificação fresca e marca pendência", async () => {
    const ws = new PDFWorkspace(); await ws.abrir(arquivo(await criarPDFAssinado()));
    const antes = ws.capturarEstado();
    const proposta = await ws.prepararCopiaImpressao(config, { rasterizar });
    assert.deepEqual(ws.capturarEstado(), antes); assert.throws(() => ws.aplicarCopiaImpressao(proposta), /conferência/);
    proposta.verificacao = { compativel: true }; ws.aplicarCopiaImpressao(proposta);
    assert.equal(ws.alterado, true); assert.equal(ws.nomeArquivo, "assinado_IMPRESSAO.pdf");
    assert.equal(ws.obterNomeExportacao(), "assinado_IMPRESSAO.pdf");
    assert.equal(ws.obterResumoAssinaturas().preparada, true); assert.equal(ws.obterResumoAssinaturas().presentes, 0);
    verificarSemAssinaturas(ws.pdf);
    assert.throws(() => ws.aplicarCopiaImpressao(proposta), /vencida/);
    ws.restaurarEstado(antes);
    await assert.rejects(ws.prepararCopiaImpressao(config, { rasterizar: async () => { throw new Error("Render falhou"); } }), /Render falhou/);
    assert.deepEqual(ws.capturarEstado(), antes);
    const ac = new AbortController(); ac.abort();
    await assert.rejects(ws.prepararCopiaImpressao(config, { signal: ac.signal, rasterizar }), { name: "AbortError" });
});
test("limites de memória mantêm A4 300/600 e A3 300; A3 600 é bloqueado sem reduzir DPI", () => {
    const pt = mm => mm * 72 / 25.4;
    for (const [w, h, dpi] of [[210, 297, 300], [210, 297, 600], [297, 420, 300]]) {
        const l = estimarRaster(pt(w), pt(h), dpi); assert.ok(l.pixels <= 36000000); assert.equal(l.bytesCanvas, l.pixels * 4);
    }
    assert.throws(() => estimarRaster(pt(297), pt(420), 600), /36 milhões/);
    assert.equal(estimarRaster(570, 420, 300).largura, 2375, "UserUnit não deve gerar uma coluna extra por erro numérico");
});

test("PDF-008: origem assinada não é salva ao abrir, visualizar ou recusar preparo",async()=>{
 const bytes=await criarFixtureAssinada(),ws=new PDFWorkspace(),salvar=PDFDocument.prototype.save;let saves=0;PDFDocument.prototype.save=function(...args){++saves;return salvar.apply(this,args);};try{await ws.abrir(arquivo(bytes));assert.deepEqual(await ws.obterBytesVisualizacao(),bytes);await assert.rejects(ws.prepararCopiaImpressao(config,{rasterizar}),{codigo:'PDF_ASSINATURA_PROTEGIDA'});assert.equal(saves,0);}finally{PDFDocument.prototype.save=salvar;}
});


test("nome da cópia distingue inclusive um original já chamado _IMPRESSAO", async () => {
    const bytes = await criarPDFAssinado(), ws = new PDFWorkspace();
    await ws.abrir({ ...arquivo(bytes), name: "original_IMPRESSAO.pdf" });
    const proposta = await ws.prepararCopiaImpressao(config, { rasterizar }); proposta.verificacao = { compativel: true };
    ws.aplicarCopiaImpressao(proposta);
    assert.equal(ws.nomeArquivo, "original_IMPRESSAO_COPIA_IMPRESSAO.pdf");
});

test("PDF-008: assinatura aplicada bloqueia reconstrução mesmo com callbacks disponíveis",async()=>{
 const ws=new PDFWorkspace();await ws.abrir(arquivo(await criarFixtureAssinada()));let chamadas=0;await assert.rejects(ws.prepararCopiaImpressao(config,{rasterizar:()=>{chamadas++;}}),{codigo:'PDF_ASSINATURA_PROTEGIDA'});assert.equal(chamadas,0);assert.equal(ws.obterResumoAssinaturas().presentesPendentes,1);
});
