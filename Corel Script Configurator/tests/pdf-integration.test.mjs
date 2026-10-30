import { Elemento, JanelaTeste } from "./helpers/pdf-dom.mjs";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { setImmediate as proximoTurno } from "node:timers/promises";
import test from "node:test";
import { PDFDocument } from "../public/vendor/pdf-lib/pdf-lib.js";
import * as pdfjs from "../public/vendor/pdfjs/pdf.js";
import { PDFDocumentoVisual } from "../public/js/tools/pdf-visual-document.js";
import { marcarEstrutura, criarFixtureImpressao, textoImpressao } from "./helpers/pdf-print-fixture.mjs";
import { criarPDFAssinado } from "./helpers/pdf-signature-fixture.mjs";

// Adaptações exclusivas do teste no Node 24; nenhum arquivo de produção é modificado.
Uint8Array.prototype.toHex ??= function () { return Buffer.from(this).toString("hex"); };
const url = caminho => new URL(caminho, import.meta.url).href;
const modulo = codigo => "data:text/javascript;base64," + Buffer.from(codigo).toString("base64");
const biblioteca = url("../public/vendor/pdf-lib/pdf-lib.js");
async function carregar(caminho, caminhos) {
    let codigo = await readFile(new URL(caminho, import.meta.url), "utf8");
    for (const [origem, destino] of Object.entries(caminhos)) {
        codigo = codigo.replaceAll(JSON.stringify(origem), JSON.stringify(destino));
    }
    return modulo(codigo);
}
const impressao = await carregar("../public/js/tools/pdf-print-compatibility.js", { "/vendor/pdf-lib/pdf-lib.js": biblioteca });
const motorResize = await carregar("../public/js/tools/pdf-resize-engine.js", { "/vendor/pdf-lib/pdf-lib.js": biblioteca, "./pdf-formats.js": url("../public/js/tools/pdf-formats.js"), "./pdf-print-compatibility.js": impressao });
const assinaturas = await carregar("../public/js/tools/pdf-signatures.js", { "/vendor/pdf-lib/pdf-lib.js": biblioteca });
const copiaImpressao = await carregar("../public/js/tools/pdf-print-copy.js", { "/vendor/pdf-lib/pdf-lib.js": biblioteca, "./pdf-signatures.js": assinaturas });
const workspace = await carregar("../public/js/tools/pdf-workspace.js", { "/vendor/pdf-lib/pdf-lib.js": biblioteca, "./pdf-resize-engine.js": motorResize, "./pdf-signatures.js": assinaturas, "./pdf-print-copy.js": copiaImpressao });
const editor = await carregar("../public/js/tools/pdf-editor.js", { "/vendor/pdf-lib/pdf-lib.js": biblioteca });
const viewer = modulo(`
    export const GlobalWorkerOptions = { set workerSrc(_) {} };
    export function getDocument(options) { return globalThis.__pdfIntegration.load(options); }
`);
const preview = await carregar("../public/js/tools/pdf-preview.js", { "/vendor/pdfjs/pdf.js": viewer });
const controlador = await carregar("../public/js/tools/pdfs.js", {
    "../app/save-file.js": url("../public/js/app/save-file.js"),
    "../app/desktop-client.js": url("../public/js/app/desktop-client.js"),
    "../app/export-dialog.js": url("../public/js/app/export-dialog.js"),
    "./pdf-dimension-groups.js": url("../public/js/tools/pdf-dimension-groups.js"),
    "./pdf-view-controls.js": url("../public/js/tools/pdf-view-controls.js"),
    "./pdf-page-selection.js": url("../public/js/tools/pdf-page-selection.js"),
    "./pdf-page-actions.js": url("../public/js/tools/pdf-page-actions.js"),
    "./pdf-panel-layout.js": url("../public/js/tools/pdf-panel-layout.js"),
    "./pdf-print-session.js": url("../public/js/tools/pdf-print-session.js"),
    "./pdf-print-ui.js": url("../public/js/tools/pdf-print-ui.js"),
    "./pdf-resize-ui.js": url("../public/js/tools/pdf-resize-ui.js"),
    "./pdf-resize-session.js": url("../public/js/tools/pdf-resize-session.js"),
    "./pdf-visual-document.js": url("../public/js/tools/pdf-visual-document.js"),
    "./pdf-document-info.js": url("../public/js/tools/pdf-document-info.js"),
    "./pdf-protection.js": url("../public/js/tools/pdf-protection.js"),
    "./pdf-preview.js": preview,
    "./pdf-thumbnails.js": url("../public/js/tools/pdf-thumbnails.js"),
    "./pdf-interface.js": url("../public/js/tools/pdf-interface.js"),
    "./pdf-workspace.js": workspace,
    "./pdf-editor.js": editor,
    "./pdf-merge.js": url("../public/js/tools/pdf-merge.js"),
    "/vendor/pdfjs/pdf.js": viewer
});
const { renderPDFs } = await import(controlador);
const { PDFPreview } = await import(preview);
pdfjs.GlobalWorkerOptions.workerSrc = url("../public/vendor/pdfjs/pdf.worker.js");

// DOM mínimo para acionar os módulos reais. A rasterização do canvas é substituída;
// leitura, dimensões e extração usam PDF.js/pdf-lib locais com bytes PDF reais.
async function arquivo(nome, larguras) {
    const pdf = await PDFDocument.create();
    larguras.forEach(largura => pdf.addPage([largura, 300]));
    const bytes = await pdf.save();
    return { name: nome, arrayBuffer: async () => bytes.slice().buffer };
}

async function aguardar(condicao) {
    const limite = Date.now() + 5000;
    while (!condicao()) {
        assert.ok(Date.now() < limite, "Operação não concluiu");
        await proximoTurno();
    }
}

function iniciar() {
    const body = new Elemento("body");
    const app = new Elemento("main");
    body.append(app);
    globalThis.document = { body, getElementById: () => app, createElement: tag => new Elemento(tag) };
    globalThis.window = new JanelaTeste();
    const estado = globalThis.__pdfIntegration = {
        downloads: [], blobs: new Map(), tasks: [], previews: [], failNext: false,
        load(options) {
            if (this.failNext) {
                this.failNext = false;
                return { promise: Promise.reject(new Error("Falha simulada na prévia")), destroy: async () => {} };
            }
            const task = pdfjs.getDocument(options);
            const registro = { destruicoes: 0 };
            this.tasks.push(registro);
            return {
                destroy: async () => {
                    assert.equal(++registro.destruicoes, 1, "Destruição duplicada da mesma tarefa");
                    await task.destroy();
                },
                promise: task.promise.then(doc => {
                    const getPage = doc.getPage.bind(doc);
                    doc.getPage = async numero => {
                        const pagina = await getPage(numero);
                        pagina.render = ({ viewport, canvas }) => {
                            if (canvas.className === "pdf-miniatura-canvas") return { promise: Promise.resolve(), cancel() {} };
                            if (this.failRenderNext || this.failRenderAll) {
                                this.failRenderNext = false;
                                return { promise: Promise.reject(new Error("Falha de render simulada")), cancel() {} };
                            }
                            this.previews.push({ numero, giro: viewport.rotation });
                            this.renderStarted = true;
                            let cancelar;
                            const cancelada = new Promise((_, reject) => { cancelar = () => reject(Object.assign(new Error("Cancelada"), { name: "RenderingCancelledException" })); });
                            return { promise: Promise.race([this.renderGate || Promise.resolve(), cancelada]), cancel: cancelar };
                        };
                        return pagina;
                    };
                    return doc;
                })
            };
        }
    };
    let destruirAtual = renderPDFs();
    const destruir = () => destruirAtual();
    const remontar = async () => { await destruirAtual(); app.innerHTML = ""; destruirAtual = renderPDFs(); };
    const buscar = id => app.querySelector("#" + id);
    const pronto = () => aguardar(() => !buscar("pdf-arquivo").disabled);
    const clicar = async id => {
        buscar(id).click();
        if (id === "pdf-salvar-copia") { await proximoTurno(); buscar("app-exportacao-confirmar").click(); }
        await pronto(); await proximoTurno();
    };
    const selecionar = (id, arquivos) => { buscar(id).files = arquivos; buscar(id).emitir("change"); };
    const adicionar = async arquivos => {
        buscar("pdf-uniao-abrir").click();
        selecionar("pdf-uniao-arquivos", arquivos);
        await clicar("pdf-uniao-gerar");
    };
    const confirmarSaida = () => destruirAtual.confirmarSaida();
    return { ...estado, estado, app, buscar, pronto, clicar, selecionar, adicionar, destruir, remontar, confirmarSaida,
        suspender: () => destruirAtual.suspender(), retomar: () => destruirAtual.retomar() };
}

test("suspender/retomar conserva canvas, documento, navegação, painéis e proteção sem recriar PDF.js", async () => {
    const ui = iniciar();
    try {
        ui.selecionar("pdf-arquivo", [await arquivo("aba.pdf", [200,300,400])]); await ui.pronto();
        await ui.clicar("pdf-proxima"); await ui.clicar("pdf-girar-direita");
        ui.buscar("pdf-ferramenta-organizar").click();
        ui.buscar("pdf-visualizacao").scrollTop = 70;
        const canvas = ui.buscar("pdf-visualizacao").querySelector("canvas"), tarefas = ui.tasks.length;
        for (let i=0;i<4;i++) { await ui.suspender(); await proximoTurno(); await ui.retomar(); }
        assert.equal(ui.tasks.length,tarefas); assert.equal(ui.buscar("pdf-visualizacao").querySelector("canvas"),canvas);
        assert.equal(ui.buscar("pdf-pagina-atual").textContent,"2 / 3");
        assert.equal(ui.buscar("pdf-visualizacao").scrollTop,70);
        assert.equal(ui.buscar("pdf-opcoes-organizar").hidden,false);
        assert.equal(ui.buscar("pdf-alteracoes-pendentes").hidden,false);
        assert.equal(globalThis.window.listeners.get("beforeunload").size,1);
        assert.ok(ui.tasks.every(t => t.destruicoes === 0));
    } finally { await ui.destruir(); delete globalThis.document; delete globalThis.window; }
});

test("orientações: giro, reorganização, exclusão, união e resize atualizam indicador e diagnóstico de exportação", async () => {
    const ui = iniciar();
    const misturadas = esperado => {
        assert.equal(ui.buscar("pdf-informacoes")["pdf-alerta"], esperado);
        assert.equal(ui.buscar("pdf-orientacoes-alerta").hidden, !esperado);
        if (esperado) assert.match(ui.buscar("pdf-alertas").textContent, /Orientações diferentes detectadas/);
        else assert.doesNotMatch(ui.buscar("pdf-alertas").textContent, /Orientações diferentes detectadas/);
        ui.buscar("pdf-salvar-copia").click();
        const avisos = ui.buscar("app-exportacao-avisos").textContent;
        assert.equal(avisos.includes("Orientações diferentes detectadas"), esperado);
        if (esperado) assert.ok(avisos.includes(ui.buscar("pdf-orientacoes-detalhes").textContent));
        ui.buscar("app-exportacao-cancelar").click();
    };
    try {
        ui.selecionar("pdf-arquivo", [await arquivo("orientacoes.pdf", [200,200])]); await ui.pronto();
        misturadas(false);
        const fontes = ui.estado.tasks.length;
        await ui.clicar("pdf-girar-direita"); misturadas(true);
        assert.match(ui.buscar("pdf-orientacoes-detalhes").textContent, /Retrato: 1 página\(s\) — páginas 2.*Paisagem: 1 página\(s\) — páginas 1/);
        assert.doesNotMatch(ui.buscar("pdf-alertas").textContent, /tamanhos diferentes/);
        await ui.clicar("pdf-girar-esquerda"); misturadas(false);
        await ui.clicar("pdf-girar-esquerda"); misturadas(true); // 270°
        ui.buscar("pdf-mover-destino").value = "2"; await ui.clicar("pdf-mover-aplicar"); misturadas(true);
        assert.match(ui.buscar("pdf-orientacoes-detalhes").textContent, /Retrato: 1 página\(s\) — páginas 1.*Paisagem: 1 página\(s\) — páginas 2/);
        assert.equal(ui.estado.tasks.length, fontes, "Diagnóstico e giros/movimento não recarregam o PDF.js");
        ui.buscar("pdf-selecao-paginas").value = "2"; await ui.clicar("pdf-excluir"); misturadas(false);
        await ui.adicionar([await arquivo("paisagem.pdf", [450])]); misturadas(true);
        assert.match(ui.buscar("pdf-alertas").textContent, /tamanhos diferentes/);
        ui.buscar("pdf-ferramenta-redimensionar").click();
        ui.buscar("pdf-redim-alvo").value = "todas";
        ui.buscar("pdf-redim-prever").click();
        await aguardar(() => !ui.buscar("pdf-redim-aplicar").disabled);
        await ui.clicar("pdf-redim-aplicar"); misturadas(false);
        assert.equal(ui.buscar("pdf-alteracoes-pendentes").hidden, false);
        assert.equal(ui.downloads.length, 0, "Consultar diagnóstico nunca exporta automaticamente");
    } finally { await ui.destruir(); delete globalThis.document; delete globalThis.window; }
});

test("orientação normalizada não apaga o alerta de assinatura nem seu diagnóstico na exportação", async () => {
    const ui = iniciar();
    try {
        const bytes = await criarPDFAssinado({ total:2 });
        ui.selecionar("pdf-arquivo", [{name:"assinado.pdf", arrayBuffer:async()=>bytes.slice().buffer}]); await ui.pronto();
        await ui.clicar("pdf-girar-direita");
        assert.match(ui.buscar("pdf-informacoes").textContent, /2 alertas/);
        await ui.clicar("pdf-girar-esquerda");
        assert.equal(ui.buscar("pdf-orientacoes-alerta").hidden, true);
        assert.equal(ui.buscar("pdf-informacoes")["pdf-alerta"], true);
        assert.match(ui.buscar("pdf-informacoes").textContent, /1 alerta/);
        assert.equal(ui.buscar("pdf-assinaturas-alerta").hidden, false);
        ui.buscar("pdf-salvar-copia").click();
        assert.match(ui.buscar("app-exportacao-avisos").textContent, /invalidar assinaturas/);
        assert.doesNotMatch(ui.buscar("app-exportacao-avisos").textContent, /Orientações diferentes/);
        ui.buscar("app-exportacao-cancelar").click();
    } finally { await ui.destruir(); delete globalThis.document; delete globalThis.window; }
});

test("Desktop: cancelar/falhar no salvamento nativo preserva pendências e só a gravação concluída quita a versão", async () => {
    const ui=iniciar();
    try {
        ui.selecionar('pdf-arquivo',[await arquivo('desktop.pdf',[200,200])]);await ui.pronto();await ui.clicar('pdf-girar-direita');
        let comportamento=async()=>({saved:false}),chamadas=0;const estados=[];
        globalThis.window.desktop4PrintShops={salvarPDF:async(bytes,nome)=>{assert.ok(bytes instanceof Uint8Array);assert.equal(nome,'desktop_editado.pdf');chamadas++;return comportamento();},estadoDocumento:estado=>estados.push(estado)};
        await ui.clicar('pdf-salvar-copia');assert.equal(ui.buscar('pdf-alteracoes-pendentes').hidden,false);
        ui.buscar('app-exportacao-cancelar').click();comportamento=async()=>{throw new Error('Disco indisponível');};
        await ui.clicar('pdf-salvar-copia');assert.equal(ui.buscar('pdf-alteracoes-pendentes').hidden,false);
        assert.match(ui.buscar('app-exportacao-status').textContent,/Disco indisponível/);
        let concluir;comportamento=()=>new Promise(resolve=>{concluir=resolve;});ui.buscar('app-exportacao-confirmar').click();
        await aguardar(()=>!!concluir);assert.equal(ui.buscar('pdf-alteracoes-pendentes').hidden,false);
        concluir({saved:true});await ui.pronto();await proximoTurno();
        assert.equal(ui.buscar('pdf-alteracoes-pendentes').hidden,true);assert.equal(chamadas,3);assert.equal(ui.downloads.length,0);
        assert.deepEqual(estados.at(-1),{pending:false,busy:false});
    }finally{await ui.destruir();delete globalThis.document;delete globalThis.window;}
});

test("painel contextual mostra apenas a ferramenta escolhida e pode ser recolhido", async () => {
    const ui = iniciar();
    try {
        assert.equal(ui.buscar("pdf-opcoes-organizar").hidden, true);
        ui.buscar("pdf-ferramenta-organizar").click();
        assert.equal(ui.buscar("pdf-opcoes-organizar").hidden, false);
        assert.equal(ui.buscar("pdf-opcoes-rotacionar").hidden, true);
        ui.buscar("pdf-ferramenta-unir").click();
        assert.equal(ui.buscar("pdf-opcoes-unir").hidden, false);
        assert.equal(ui.buscar("pdf-opcoes-organizar").hidden, true);
        assert.equal(ui.buscar("pdf-ferramenta-unir").getAttribute("aria-expanded"), "true");
        ui.buscar("pdf-painel-alternar").click();
        assert.equal(ui.buscar("pdf-painel-ferramentas").hidden, true);
        assert.equal(ui.buscar("pdf-painel-alternar").getAttribute("aria-expanded"), "false");
        ui.buscar("pdf-painel-alternar").click();
        assert.equal(ui.buscar("pdf-painel-ferramentas").hidden, false);
    } finally {
        await ui.destruir();
        delete globalThis.document;
    }
});

test("redimensionar separa medidas e encaixe, com cadeado único entre os campos e descrições dos modos", async () => {
    const ui = iniciar();
    try {
        const cadeado = ui.buscar("pdf-redim-vincular"), grupo = cadeado.parent;
        assert.equal(ui.app.querySelectorAll("#pdf-redim-vincular").length, 1);
        assert.equal(grupo.children[1], cadeado);
        assert.equal(grupo.children[0].querySelector("#pdf-redim-largura"), ui.buscar("pdf-redim-largura"));
        assert.equal(grupo.children[2].querySelector("#pdf-redim-altura"), ui.buscar("pdf-redim-altura"));
        assert.equal(cadeado.querySelectorAll("svg").length, 2);
        assert.equal(ui.buscar("pdf-redim-proporcao").textContent, "210 × 297 mm · ≈ 1:1,41");
        for (const [modo, texto] of [['ajustar', 'Cabe inteiro; pode sobrar margem.'], ['centralizar', 'Mantém o tamanho original do conteúdo.'], ['preencher', 'Preenche toda a folha; pode cortar bordas.']]) {
            ui.buscar("pdf-redim-modo").value = modo; ui.buscar("pdf-redim-modo").emitir("change"); await proximoTurno();
            assert.equal(ui.buscar("pdf-redim-modo-descricao").textContent, texto);
        }
        assert.equal(ui.buscar("pdf-alteracoes-pendentes").hidden, true);
        assert.equal(ui.downloads.length, 0);
    } finally { await ui.destruir(); delete globalThis.document; delete globalThis.window; }
});

test("alerta de assinatura distingue origem e posição atual; Resolver abre/cancela sem editar ou baixar", async () => {
    const ui = iniciar();
    try {
        const bytes = await criarPDFAssinado();
        ui.selecionar("pdf-arquivo", [{ name: "assinada.pdf", arrayBuffer: async () => bytes.slice().buffer }]); await ui.pronto();
        assert.equal(ui.buscar("pdf-assinaturas-alerta").hidden, false);
        assert.equal(ui.buscar("pdf-assinaturas-titulo").textContent, "Assinatura digital detectada");
        assert.match(ui.buscar("pdf-assinaturas-resumo").textContent, /visuais atuais: 3/);
        assert.match(ui.buscar("pdf-assinaturas-detalhes").textContent, /Páginas visuais na origem: 3/);
        assert.equal(ui.buscar("pdf-assinaturas-resolver").textContent, "Resolver");
        assert.equal(ui.buscar("pdf-assinaturas-expandido").hidden, true);
        ui.buscar("pdf-assinaturas-expandir").click();
        assert.equal(ui.buscar("pdf-assinaturas-expandir").getAttribute("aria-expanded"), "true");
        assert.equal(ui.buscar("pdf-assinaturas-expandido").textContent, ui.buscar("pdf-assinaturas-detalhes").textContent);
        await ui.clicar("pdf-proxima"); await ui.clicar("pdf-proxima");
        ui.buscar("pdf-mover-destino").value = "1"; await ui.clicar("pdf-mover-aplicar");
        assert.match(ui.buscar("pdf-assinaturas-resumo").textContent, /visuais atuais: 1/);
        assert.equal(ui.buscar("pdf-assinaturas-expandido").hidden, false, "navegar e reorganizar preservam os detalhes abertos");
        ui.buscar("pdf-assinaturas-expandir").click();
        assert.equal(ui.buscar("pdf-assinaturas-expandido").hidden, true);
        assert.match(ui.buscar("pdf-assinaturas-detalhes").textContent, /Páginas visuais na origem: 3/);
        ui.buscar("pdf-assinaturas-resolver").click(); await aguardar(() => ui.buscar("pdf-impressao-dialogo").open);
        assert.equal(ui.buscar("pdf-arquivo").disabled, true);
        assert.equal(ui.buscar("pdf-salvar-copia").disabled, true);
        assert.equal(ui.buscar("pdf-impressao-detalhes").open, false);
        assert.equal(ui.buscar("pdf-impressao-controles-revisao").hidden, true);
        assert.equal(ui.buscar("pdf-impressao-aplicar").hidden, true);
        ui.buscar("pdf-impressao-detalhes").open = true;
        assert.equal(ui.downloads.length, 0);
        ui.buscar("pdf-impressao-cancelar").click(); await ui.pronto();
        ui.buscar("pdf-assinaturas-resolver").click(); await aguardar(() => ui.buscar("pdf-impressao-dialogo").open);
        assert.equal(ui.buscar("pdf-impressao-detalhes").open, false, "reabrir recolhe explicações anteriores");
        assert.equal(ui.buscar("pdf-impressao-resultado-tecnico").textContent, "");
        ui.buscar("pdf-impressao-cancelar").click(); await ui.pronto();
        assert.equal(ui.buscar("pdf-alteracoes-pendentes").hidden, false);
        assert.equal(ui.buscar("pdf-pagina-atual").textContent, "1 / 3");
        assert.equal(ui.downloads.length, 0);
    } finally { await ui.destruir(); delete globalThis.document; }
});

test("campo vazio não recebe título de assinatura detectada e reentradas não acumulam controles Resolver", async () => {
    const ui = iniciar();
    try {
        const bytes = await criarPDFAssinado({ vazia: true });
        ui.selecionar("pdf-arquivo", [{ name: "vazia.pdf", arrayBuffer: async () => bytes.slice().buffer }]); await ui.pronto();
        assert.equal(ui.buscar("pdf-assinaturas-titulo").textContent, "Campo de assinatura vazio");
        const resolverAntigo = ui.buscar("pdf-assinaturas-resolver");
        resolverAntigo.click(); await aguardar(() => ui.buscar("pdf-impressao-dialogo").open);
        await ui.remontar();
        resolverAntigo.click(); await proximoTurno();
        assert.ok(!ui.buscar("pdf-impressao-dialogo").open);
        assert.equal(ui.buscar("pdf-assinaturas-alerta").hidden, true);
        assert.equal(globalThis.window.listeners.get("beforeunload")?.size || 0, 0);
        assert.ok(ui.tasks.every(t => t.destruicoes === 1));
    } finally { await ui.destruir(); delete globalThis.document; }
});

test("duplicação recupera falha de prévia e permite repetir sem perder seleção ou criar pendência falsa", async () => {
    const ui = iniciar();
    try {
        ui.selecionar("pdf-arquivo", [await arquivo("duplicar.pdf", [200, 300])]); await ui.pronto();
        await ui.clicar("pdf-proxima"); const canvas = ui.buscar("pdf-visualizacao").querySelector("canvas");
        ui.buscar("pdf-paginas-acoes").click(); ui.estado.failNext = true;
        await ui.clicar("pdf-duplicar");
        assert.equal(ui.buscar("pdf-total").textContent, 2);
        assert.equal(ui.buscar("pdf-pagina-atual").textContent, "2 / 2");
        assert.equal(ui.buscar("pdf-alteracoes-pendentes").hidden, true);
        assert.equal(ui.buscar("pdf-visualizacao").querySelector("canvas"), canvas);
        assert.match(ui.buscar("pdf-status").textContent, /Falha simulada/);
        ui.buscar("pdf-paginas-acoes").click(); await ui.clicar("pdf-duplicar");
        assert.equal(ui.buscar("pdf-total").textContent, 3);
        assert.equal(ui.buscar("pdf-pagina-atual").textContent, "3 / 3");
        assert.equal(ui.buscar("pdf-alteracoes-pendentes").hidden, false);
        assert.equal(ui.downloads.length, 0);
    } finally { await ui.destruir(); delete globalThis.document; delete globalThis.window; }
});

test("modal inicia, ordena, acrescenta, edita, extrai e exporta sem baixar na união", async () => {
    const createURL = URL.createObjectURL;
    const revokeURL = URL.revokeObjectURL;
    const timer = globalThis.setTimeout;
    URL.createObjectURL = blob => {
        const id = "blob:" + globalThis.__pdfIntegration.blobs.size;
        globalThis.__pdfIntegration.blobs.set(id, blob);
        return id;
    };
    URL.revokeObjectURL = () => {};
    globalThis.setTimeout = (callback, ms, ...args) => ms === 60000 ? (callback(), 0) : timer(callback, ms, ...args);
    const ui = iniciar();
    try {
        const a = await arquivo("a.pdf", [200, 210]);
        const b = await arquivo("b.pdf", [400]);
        ui.buscar("pdf-uniao-abrir").click();
        ui.selecionar("pdf-uniao-arquivos", [a, b, b]);
        ui.buscar("pdf-uniao-lista").children[2].children[1].children[2].click();
        assert.equal(ui.buscar("pdf-uniao-lista").children.length, 2);
        ui.buscar("pdf-uniao-fechar").click();
        assert.equal(ui.buscar("pdf-uniao-janela").open, false);
        ui.buscar("pdf-uniao-abrir").click();
        const subir = ui.buscar("pdf-uniao-lista").children[1].children[1].children[0];
        subir.click();
        await ui.clicar("pdf-uniao-gerar");
        assert.equal(ui.downloads.length, 0);
        assert.equal(ui.buscar("pdf-uniao-janela").open, false);
        assert.equal(ui.buscar("pdf-uniao-lista").children.length, 0);
        assert.equal(ui.buscar("pdf-total").textContent, 3);
        assert.equal(ui.buscar("pdf-tamanho").textContent, "Variados");
        assert.equal(ui.buscar("pdf-informacoes")["pdf-alerta"], true);
        assert.equal(ui.buscar("pdf-pagina-atual").textContent, "1 / 3");
        assert.equal(ui.buscar("pdf-salvar-copia").disabled, false);

        await ui.clicar("pdf-girar-direita");
        ui.buscar("pdf-selecao-paginas").value = "2";
        await ui.clicar("pdf-excluir");
        await ui.adicionar([await arquivo("novo.pdf", [600])]);
        assert.equal(ui.buscar("pdf-total").textContent, 3);
        assert.equal(ui.previews.at(-1).giro, 90);
        assert.equal(ui.downloads.length, 0);
        await ui.clicar("pdf-proxima");
        await ui.clicar("pdf-proxima");
        assert.equal(ui.buscar("pdf-pagina-atual").textContent, "3 / 3");

        await ui.clicar("pdf-salvar-copia");
        const exportado = await PDFDocument.load(await ui.downloads[0].blob.arrayBuffer());
        assert.deepEqual(exportado.getPages().map(p => p.getWidth()), [400, 210, 600]);
        assert.equal(exportado.getPage(0).getRotation().angle, 90);
        ui.buscar("pdf-selecao-paginas").value = "1,3";
        await ui.clicar("pdf-extrair");
        const extraido = await PDFDocument.load(await ui.downloads[1].blob.arrayBuffer());
        assert.deepEqual(extraido.getPages().map(p => p.getWidth()), [400, 600]);
        assert.equal(extraido.getPage(0).getRotation().angle, 90);

        ui.selecionar("pdf-arquivo", [a]);
        await ui.pronto();
        await ui.adicionar([b]);
        assert.equal(ui.buscar("pdf-nome").textContent, "a.pdf");
        assert.equal(ui.buscar("pdf-total").textContent, 3);
    } finally {
        for (const task of ui.tasks) assert.ok(task.destruicoes <= 1);
        URL.createObjectURL = createURL;
        URL.revokeObjectURL = revokeURL;
        globalThis.setTimeout = timer;
        await ui.destruir();
        delete globalThis.document;
    }
});

test("bloqueia concorrência, recupera falhas e isola a instância anterior", async () => {
    const ui = iniciar();
    const valido = await arquivo("base.pdf", [200]);
    const logErro = console.error;
    const erros = [];
    try {
        ui.selecionar("pdf-arquivo", [valido]);
        await ui.pronto();
        const original = ui.buscar("pdf-nome").textContent;
        ui.estado.failNext = true;
        console.error = erro => erros.push(erro.message);
        await ui.adicionar([valido]);
        console.error = logErro;
        assert.deepEqual(erros, ["Falha simulada na prévia"]);
        assert.equal(ui.buscar("pdf-total").textContent, 1);
        assert.equal(ui.buscar("pdf-nome").textContent, original);
        assert.equal(ui.buscar("pdf-uniao-lista").children.length, 1);
        await ui.clicar("pdf-uniao-gerar");
        assert.equal(ui.buscar("pdf-total").textContent, 2, "Repetição não deve duplicar o lote que falhou");

        let liberar;
        const bloqueio = new Promise(resolve => { liberar = resolve; });
        const lento = { name: "lento.pdf", arrayBuffer: async () => { await bloqueio; return valido.arrayBuffer(); } };
        let liberarRender;
        ui.estado.renderGate = new Promise(resolve => { liberarRender = resolve; });
        ui.estado.renderStarted = false;
        ui.buscar("pdf-proxima").click();
        await aguardar(() => ui.estado.renderStarted);
        ui.buscar("pdf-uniao-abrir").click();
        ui.selecionar("pdf-uniao-arquivos", [lento]);
        ui.buscar("pdf-uniao-gerar").click();
        assert.equal(ui.buscar("pdf-arquivo").disabled, true);
        assert.equal(ui.buscar("pdf-girar-direita").disabled, true);
        assert.equal(ui.buscar("pdf-salvar-copia").disabled, true);
        ui.buscar("pdf-uniao-gerar").emitir("click");
        ui.buscar("pdf-girar-direita").emitir("click");
        liberarRender();
        ui.estado.renderGate = null;
        await proximoTurno();
        assert.equal(ui.buscar("pdf-pagina-atual").textContent, "1 / 2", "Renderização antiga não deve trocar a página durante a união");
        liberar();
        await ui.pronto();
        assert.equal(ui.buscar("pdf-total").textContent, 3);
        assert.equal(ui.previews.at(-1).giro, 0);

        let continuar;
        const espera = new Promise(resolve => { continuar = resolve; });
        ui.selecionar("pdf-uniao-arquivos", [{ name: "pausado.pdf", arrayBuffer: async () => { await espera; return valido.arrayBuffer(); } }]);
        ui.buscar("pdf-uniao-gerar").click();
        const statusAntigo = ui.buscar("pdf-status");
        await ui.remontar();
        const novoStatus = ui.buscar("pdf-status");
        continuar();
        await aguardar(() => !ui.app.querySelector("#pdf-arquivo").disabled);
        await aguardar(() => !ui.estado.tasks.at(-1) || ui.estado.tasks.at(-1).destruicoes === 1);
        assert.notEqual(novoStatus, statusAntigo);
        assert.equal(ui.buscar("pdf-conteudo").hidden, true);
        assert.equal(ui.buscar("pdf-salvar-copia").disabled, true);
        assert.equal(ui.app.listeners.get("click"), undefined);

        // Importação também impede iniciar a união durante a leitura.
        let concluirImportacao;
        const importacao = new Promise(resolve => { concluirImportacao = resolve; });
        ui.selecionar("pdf-arquivo", [{ name: "importacao.pdf", arrayBuffer: async () => { await importacao; return valido.arrayBuffer(); } }]);
        assert.equal(ui.buscar("pdf-uniao-arquivos").disabled, true);
        concluirImportacao();
        await ui.pronto();
        await ui.adicionar([valido]);
        assert.equal(ui.buscar("pdf-total").textContent, 2);
    } finally {
        console.error = logErro;
        await ui.destruir();
        delete globalThis.document;
    }
});

test("encerramento cancela render pendente e destrói cada tarefa PDF.js uma única vez", async () => {
    const ui = iniciar();
    let liberar;
    try {
        ui.selecionar("pdf-arquivo", [await arquivo("duas.pdf", [200, 300])]);
        await ui.pronto();
        ui.estado.renderStarted = false;
        ui.estado.renderGate = new Promise(resolve => { liberar = resolve; });
        ui.buscar("pdf-proxima").click();
        await aguardar(() => ui.estado.renderStarted);
        const fim = ui.destruir();
        assert.equal(ui.destruir(), fim);
        await fim;
        assert.ok(ui.tasks.every(tarefa => tarefa.destruicoes === 1));
        assert.equal(ui.buscar("pdf-miniaturas").querySelectorAll("canvas").length, 0);
    } finally {
        liberar?.();
        await ui.destruir();
        delete globalThis.document;
    }
});

test("falha de renderização não impede a destruição das tarefas de carregamento", async () => {
    const previa = new PDFPreview(new Elemento("div"));
    let destruicoes = 0;
    previa.tarefas.add({ destroy: async () => { destruicoes++; } });
    previa.render = { cancel() {}, promise: Promise.reject(new Error("Falha de rasterização")) };
    await previa.destruir();
    assert.equal(destruicoes, 1);
    assert.equal(previa.tarefas.size, 0);
});

test("saída limpa dispensa aviso; cancelar saída alterada preserva documento e miniaturas", async () => {
    const ui = iniciar();
    try {
        ui.selecionar("pdf-arquivo", [await arquivo("cliente.pdf", [200, 300])]);
        await ui.pronto();
        assert.equal(await ui.confirmarSaida(), true);
        assert.ok(!ui.buscar("pdf-confirmar-descarte").open);
        await ui.clicar("pdf-girar-direita");
        const canvas = ui.buscar("pdf-visualizacao").querySelector("canvas");
        const miniaturas = ui.buscar("pdf-miniaturas").querySelectorAll("button");
        const resposta = ui.confirmarSaida();
        await aguardar(() => ui.buscar("pdf-confirmar-descarte").open);
        ui.buscar("pdf-continuar-editando").click();
        assert.equal(await resposta, false);
        assert.equal(ui.buscar("pdf-visualizacao").querySelector("canvas"), canvas);
        assert.deepEqual(ui.buscar("pdf-miniaturas").querySelectorAll("button"), miniaturas);
        assert.equal(ui.buscar("pdf-alteracoes-pendentes").hidden, false);
        const descarte = ui.confirmarSaida();
        await aguardar(() => ui.buscar("pdf-confirmar-descarte").open);
        ui.buscar("pdf-descartar-sair").click();
        assert.equal(await descarte, true);
    } finally {
        await ui.destruir();
        assert.equal(window.listeners.get("beforeunload").size, 0);
        delete globalThis.document;
        delete globalThis.window;
    }
});

test("importar outro PDF exige confirmação; cancelar mantém nome, páginas e giro", async () => {
    const ui = iniciar();
    try {
        ui.selecionar("pdf-arquivo", [await arquivo("original.pdf", [200, 300])]);
        await ui.pronto();
        await ui.clicar("pdf-girar-direita");
        const canvas = ui.buscar("pdf-visualizacao").querySelector("canvas");
        const novo = await arquivo("novo.pdf", [600]);
        ui.selecionar("pdf-arquivo", [novo]);
        await aguardar(() => ui.buscar("pdf-confirmar-descarte").open);
        ui.buscar("pdf-continuar-editando").click();
        await ui.pronto();
        assert.equal(ui.buscar("pdf-nome").textContent, "original.pdf");
        assert.equal(ui.buscar("pdf-total").textContent, 2);
        assert.equal(ui.buscar("pdf-visualizacao").querySelector("canvas"), canvas);
        assert.equal(ui.previews.at(-1).giro, 90);
        ui.selecionar("pdf-arquivo", [novo]);
        await aguardar(() => ui.buscar("pdf-confirmar-descarte").open);
        ui.buscar("pdf-descartar-sair").click();
        await ui.pronto();
        assert.equal(ui.buscar("pdf-nome").textContent, "novo.pdf");
        assert.equal(ui.buscar("pdf-total").textContent, 1);
        assert.equal(ui.buscar("pdf-alteracoes-pendentes").hidden, true);
        assert.equal(await ui.confirmarSaida(), true);
    } finally {
        await ui.destruir();
        delete globalThis.document;
        delete globalThis.window;
    }
});

test("exportação normal só baixa após confirmar; cancelar não edita documento nem limpa pendências", async () => {
    const ui = iniciar();
    try {
        ui.selecionar("pdf-arquivo", [await arquivo("normal.pdf",[200,300])]); await ui.pronto();
        assert.equal(ui.buscar("pdf-salvar-copia").disabled,false);
        ui.buscar("pdf-salvar-copia").click();
        assert.equal(ui.buscar("app-exportacao-dialogo").open,true); assert.equal(ui.downloads.length,0);
        assert.equal(ui.buscar("app-exportacao-nome").value,"normal_editado.pdf");
        ui.buscar("app-exportacao-cancelar").click(); assert.equal(ui.downloads.length,0);
        assert.equal(ui.buscar("pdf-alteracoes-pendentes").hidden,true);
        await ui.clicar("pdf-girar-direita"); ui.buscar("pdf-salvar-copia").click();
        ui.buscar("app-exportacao-cancelar").click(); assert.equal(ui.buscar("pdf-alteracoes-pendentes").hidden,false);
        assert.equal(ui.buscar("pdf-nome").textContent,"normal.pdf");
    } finally { await ui.destruir(); delete globalThis.document; delete globalThis.window; }
});

test("cancelar exportação durante serialização impede download e conserva alterações", async () => {
    const ui = iniciar(), {PDFWorkspace} = await import(workspace), obterBytes = PDFWorkspace.prototype.obterBytes;
    let liberar;
    try {
        ui.selecionar("pdf-arquivo", [await arquivo("cancelar.pdf",[200,300])]); await ui.pronto(); await ui.clicar("pdf-girar-direita");
        const espera = new Promise(r => liberar=r);
        PDFWorkspace.prototype.obterBytes = async function() { const bytes = await obterBytes.call(this); await espera; return bytes; };
        ui.buscar("pdf-salvar-copia").click(); ui.buscar("app-exportacao-confirmar").click();
        await aguardar(() => ui.buscar("pdf-arquivo").disabled);
        ui.buscar("app-exportacao-cancelar").click(); liberar(); await ui.pronto(); await proximoTurno();
        assert.equal(ui.downloads.length,0); assert.equal(ui.buscar("pdf-alteracoes-pendentes").hidden,false);
        assert.equal(ui.buscar("app-exportacao-dialogo").open,false);
    } finally { liberar?.(); PDFWorkspace.prototype.obterBytes=obterBytes; await ui.destruir(); delete globalThis.document; delete globalThis.window; }
});

test("cancelar fechamento conserva também a prévia temporária e suas configurações", async () => {
    const ui = iniciar();
    try {
        ui.selecionar("pdf-arquivo", [await arquivo("previa-aba.pdf",[200])]); await ui.pronto(); await ui.clicar("pdf-girar-direita");
        ui.buscar("pdf-redim-prever").click(); await aguardar(() => !ui.buscar("pdf-redim-aplicar").disabled);
        const canvas = ui.buscar("pdf-visualizacao").querySelector("canvas");
        const saida = ui.confirmarSaida(); await aguardar(() => ui.buscar("pdf-confirmar-descarte").open);
        ui.buscar("pdf-continuar-editando").click(); assert.equal(await saida,false);
        assert.equal(ui.buscar("pdf-redim-aplicar").disabled,false); assert.equal(ui.buscar("pdf-visualizacao").querySelector("canvas"),canvas);
    } finally { await ui.destruir(); delete globalThis.document; delete globalThis.window; }
});

test("exportação limpa pendências; nova edição reativa; falha de exportação não quita", async () => {
    const ui = iniciar();
    const timer = globalThis.setTimeout;
    const criarURL = URL.createObjectURL;
    const logErro = console.error;
    globalThis.setTimeout = (callback, ms, ...args) => ms === 60000 ? (callback(), 0) : timer(callback, ms, ...args);
    try {
        await ui.adicionar([await arquivo("unido.pdf", [200, 300])]);
        assert.equal(ui.buscar("pdf-alteracoes-pendentes").hidden, false);
        await ui.clicar("pdf-salvar-copia");
        assert.equal(ui.buscar("pdf-alteracoes-pendentes").hidden, true);
        assert.equal(await ui.confirmarSaida(), true);
        await ui.clicar("pdf-girar-direita");
        assert.equal(ui.buscar("pdf-alteracoes-pendentes").hidden, false);
        ui.buscar("pdf-selecao-paginas").value = "1";
        await ui.clicar("pdf-extrair");
        assert.equal(ui.buscar("pdf-alteracoes-pendentes").hidden, false, "Extrair não exporta o documento completo");
        const erros = [];
        console.error = erro => erros.push(erro.message);
        URL.createObjectURL = () => { throw new Error("Download indisponível"); };
        await ui.clicar("pdf-salvar-copia");
        assert.deepEqual(erros, []);
        assert.match(ui.buscar("app-exportacao-status").textContent, /Download indisponível/);
        assert.equal(ui.buscar("app-exportacao-dialogo").open, true);
        assert.equal(ui.buscar("pdf-alteracoes-pendentes").hidden, false);
        assert.equal(window.listeners.get("beforeunload").size, 1);
    } finally {
        globalThis.setTimeout = timer;
        URL.createObjectURL = criarURL;
        console.error = logErro;
        await ui.destruir();
        delete globalThis.document;
        delete globalThis.window;
    }
});

test("saída durante união aguarda conclusão e então protege o novo trabalho", async () => {
    const ui = iniciar();
    const base = await arquivo("base.pdf", [200]);
    let liberar;
    const espera = new Promise(resolve => { liberar = resolve; });
    try {
        const uniao = ui.adicionar([{ name: "lento.pdf", arrayBuffer: async () => { await espera; return base.arrayBuffer(); } }]);
        const saida = ui.confirmarSaida();
        assert.ok(!ui.buscar("pdf-confirmar-descarte").open);
        liberar();
        await uniao;
        await aguardar(() => ui.buscar("pdf-confirmar-descarte").open);
        ui.buscar("pdf-continuar-editando").click();
        assert.equal(await saida, false);
        assert.equal(ui.buscar("pdf-total").textContent, 1);
        assert.equal(ui.buscar("pdf-alteracoes-pendentes").hidden, false);
    } finally {
        liberar();
        await ui.destruir();
        delete globalThis.document;
        delete globalThis.window;
    }
});

test("mover por botão ou arraste mantém a página exibida e recupera falha de prévia", async () => {
    const ui = iniciar();
    const logErro = console.error;
    globalThis.requestAnimationFrame = () => 1;
    globalThis.cancelAnimationFrame = () => {};
    try {
        ui.selecionar("pdf-arquivo", [await arquivo("ordem.pdf", [200, 300, 400])]);
        await ui.pronto();
        assert.equal(ui.buscar("pdf-mover-cima").disabled, true);
        await ui.clicar("pdf-proxima");
        await ui.clicar("pdf-girar-direita");
        await ui.clicar("pdf-mover-baixo");
        assert.equal(ui.buscar("pdf-pagina-atual").textContent, "3 / 3");
        assert.equal(ui.previews.at(-1).giro, 90);
        assert.equal(ui.buscar("pdf-mover-baixo").disabled, true);
        const painel = ui.buscar("pdf-miniaturas");
        const origem = painel.querySelectorAll("button").find(b => b.dataset.pagina === "1");
        const dataTransfer = { setData() {} };
        origem.emitir("dragstart", origem, { dataTransfer });
        painel.emitir("drop", painel, { dataTransfer, clientY: 3 * 166 - painel.scrollTop });
        await ui.pronto();
        assert.equal(ui.buscar("pdf-pagina-atual").textContent, "2 / 3");
        assert.equal(ui.previews.at(-1).giro, 90, "A seleção conserva a identidade, não a posição anterior");
        assert.equal(ui.buscar("pdf-alteracoes-pendentes").hidden, false);
        assert.equal(ui.downloads.length, 0);

        const canvas = ui.buscar("pdf-visualizacao").querySelector("canvas");
        const erros = [];
        console.error = erro => erros.push(erro.message);
        const atualizar = PDFDocumentoVisual.prototype.atualizar;
        PDFDocumentoVisual.prototype.atualizar = () => { throw new Error("Mapa incompatível simulado"); };
        ui.estado.failNext = true;
        try { await ui.clicar("pdf-mover-cima"); }
        finally { PDFDocumentoVisual.prototype.atualizar = atualizar; }
        assert.deepEqual(erros, ["Falha simulada na prévia"]);
        assert.equal(ui.buscar("pdf-visualizacao").querySelector("canvas"), canvas);
        assert.equal(ui.buscar("pdf-pagina-atual").textContent, "2 / 3");
        await ui.clicar("pdf-mover-cima");
        assert.equal(ui.buscar("pdf-pagina-atual").textContent, "1 / 3");
        assert.equal(ui.previews.at(-1).giro, 90);
    } finally {
        console.error = logErro;
        await ui.destruir();
        delete globalThis.document;
        delete globalThis.window;
        delete globalThis.requestAnimationFrame;
        delete globalThis.cancelAnimationFrame;
    }
});

test("Mover para posição valida destino e move 5 para 50 fora da janela virtualizada", async () => {
    const ui = iniciar();
    try {
        ui.selecionar("pdf-arquivo", [await arquivo("100.pdf", Array.from({ length: 100 }, (_, i) => 200 + i))]);
        await ui.pronto();
        for (let i = 0; i < 4; i++) await ui.clicar("pdf-proxima");
        assert.equal(ui.buscar("pdf-pagina-atual").textContent, "5 / 100");
        for (const valor of ["", "0", "101", "2.5", "abc"]) {
            ui.buscar("pdf-mover-destino").value = valor;
            await ui.clicar("pdf-mover-aplicar");
            assert.equal(ui.buscar("pdf-pagina-atual").textContent, "5 / 100");
            assert.equal(ui.buscar("pdf-alteracoes-pendentes").hidden, true);
        }
        ui.buscar("pdf-mover-destino").value = "50";
        await ui.clicar("pdf-mover-aplicar");
        assert.equal(ui.buscar("pdf-pagina-atual").textContent, "50 / 100");
        assert.equal(ui.buscar("pdf-alteracoes-pendentes").hidden, false);
        assert.ok(ui.buscar("pdf-miniaturas").querySelectorAll("button").length <= 16);
        assert.equal(ui.buscar("pdf-mover-aplicar").disabled, false);
        assert.equal(ui.downloads.length, 0);
    } finally {
        await ui.destruir();
        delete globalThis.document;
        delete globalThis.window;
    }
});

for (const total of [100, 350]) test(`${total} páginas: giros e movimentos sucessivos reutilizam PDF.js e exportam/extraem a ordem atual`, async () => {
    const ui = iniciar();
    const ClasseWorkspace = (await import(workspace)).PDFWorkspace;
    const obterBytes = ClasseWorkspace.prototype.obterBytes;
    const criarURL = URL.createObjectURL, revogarURL = URL.revokeObjectURL, timer = globalThis.setTimeout;
    URL.createObjectURL = blob => { const id = "blob:" + ui.estado.blobs.size; ui.estado.blobs.set(id, blob); return id; };
    URL.revokeObjectURL = () => {};
    globalThis.setTimeout = (callback, ms, ...args) => ms === 60000 ? (callback(), 0) : timer(callback, ms, ...args);
    try {
        const esperadas = Array.from({ length: total }, (_, i) => ({ largura: 200 + i, giro: 0 }));
        ui.selecionar("pdf-arquivo", [await arquivo("sucessivas.pdf", esperadas.map(p => p.largura))]);
        await ui.pronto();
        for (let i = 0; i < 4; i++) await ui.clicar("pdf-proxima");
        let serializacoes = 0;
        ClasseWorkspace.prototype.obterBytes = function () { ++serializacoes; return obterBytes.call(this); };
        const selecionada = esperadas[4];
        let origem = 5;
        for (const destino of [50, 10, 75, 50]) {
            await ui.clicar("pdf-girar-direita");
            selecionada.giro = (selecionada.giro + 90) % 360;
            const canvas = ui.buscar("pdf-visualizacao").querySelector("canvas");
            ui.buscar("pdf-mover-destino").value = String(destino);
            await ui.clicar("pdf-mover-aplicar");
            esperadas.splice(destino - 1, 0, ...esperadas.splice(origem - 1, 1));
            origem = destino;
            assert.equal(ui.buscar("pdf-pagina-atual").textContent, `${destino} / ${total}`);
            assert.equal(ui.buscar("pdf-visualizacao").querySelector("canvas"), canvas, "Mover a mesma página não rasteriza de novo");
        }
        await ui.clicar("pdf-girar-esquerda");
        selecionada.giro = 270;
        assert.equal(ui.estado.tasks.length, 1, "Nenhum getDocument adicional para girar ou mover");
        assert.equal(serializacoes, 4, "Uma cópia segura por movimento; zero saves para girar/visualizar");
        ui.buscar("pdf-selecao-paginas").value = "50";
        await ui.clicar("pdf-extrair");
        const extraido = await PDFDocument.load(await ui.estado.downloads[0].blob.arrayBuffer());
        assert.equal(extraido.getPage(0).getWidth(), selecionada.largura);
        assert.equal(extraido.getPage(0).getRotation().angle, 270);
        assert.equal(ui.buscar("pdf-alteracoes-pendentes").hidden, false);
        ui.buscar("pdf-selecao-paginas").value = "1";
        await ui.clicar("pdf-excluir");
        esperadas.shift();
        await ui.clicar("pdf-girar-direita");
        esperadas[49].giro = 90;
        await ui.adicionar([await arquivo("acrescimo.pdf", [900])]);
        esperadas.push({ largura: 900, giro: 0 });
        assert.equal(ui.estado.tasks.length, 3, "Exclusão e união mantêm a alternativa completa");
        await ui.clicar("pdf-salvar-copia");
        const exportado = await PDFDocument.load(await ui.estado.downloads[1].blob.arrayBuffer());
        assert.deepEqual(exportado.getPages().map(p => [p.getWidth(), p.getRotation().angle]), esperadas.map(p => [p.largura, p.giro]));
        assert.equal(ui.buscar("pdf-alteracoes-pendentes").hidden, true);
    } finally {
        ClasseWorkspace.prototype.obterBytes = obterBytes;
        URL.createObjectURL = criarURL; URL.revokeObjectURL = revogarURL; globalThis.setTimeout = timer;
        await ui.destruir();
        delete globalThis.document; delete globalThis.window;
    }
});

test("render incremental falha uma vez e recupera pelo carregamento completo", async () => {
    const ui = iniciar();
    try {
        ui.selecionar("pdf-arquivo", [await arquivo("base.pdf", [200, 300])]);
        await ui.pronto();
        ui.estado.failRenderNext = true;
        await ui.clicar("pdf-girar-direita");
        assert.equal(ui.estado.tasks.length, 2);
        assert.equal(ui.previews.at(-1).giro, 90);
        assert.ok(ui.buscar("pdf-visualizacao").querySelector("canvas").width > 0);
        assert.equal(ui.buscar("pdf-alteracoes-pendentes").hidden, false);
    } finally {
        await ui.destruir(); delete globalThis.document; delete globalThis.window;
    }
});

test("falhas persistentes preservam giro/união anteriores; falha de recarga após excluir não exibe canvas antigo", async () => {
    const ui = iniciar();
    const logErro = console.error;
    const erros = [];
    console.error = erro => erros.push(erro.message);
    try {
        ui.selecionar("pdf-arquivo", [await arquivo("base.pdf", [200, 300, 400])]);
        await ui.pronto();
        const canvas = ui.buscar("pdf-visualizacao").querySelector("canvas");
        ui.estado.failRenderAll = true;
        await ui.clicar("pdf-girar-direita");
        assert.equal(ui.buscar("pdf-visualizacao").querySelector("canvas"), canvas);
        assert.equal(ui.buscar("pdf-alteracoes-pendentes").hidden, true, "Falha persistente restaura a revisão da rotação");
        await ui.adicionar([await arquivo("outra.pdf", [500])]);
        assert.equal(ui.buscar("pdf-total").textContent, 3);
        assert.equal(ui.buscar("pdf-visualizacao").querySelector("canvas"), canvas);
        ui.estado.failRenderAll = false;
        await ui.clicar("pdf-girar-direita");
        assert.equal(ui.previews.at(-1).giro, 90);
        ui.estado.failNext = true;
        ui.buscar("pdf-selecao-paginas").value = "1";
        await ui.clicar("pdf-excluir");
        assert.equal(ui.buscar("pdf-total").textContent, 2);
        assert.equal(ui.buscar("pdf-pagina-atual").textContent, "1 / 2");
        assert.equal(ui.buscar("pdf-visualizacao").querySelector("canvas"), undefined);
        await ui.clicar("pdf-proxima");
        assert.equal(ui.previews.at(-1).numero, 3);
        assert.equal(ui.previews.at(-1).giro, 0);
        assert.equal(ui.buscar("pdf-arquivo").disabled, false);
        assert.equal(erros.length, 3);
    } finally {
        console.error = logErro;
        await ui.destruir(); delete globalThis.document; delete globalThis.window;
    }
});

test("reutilizar canvas invalida navegação antiga ainda aguardando getPage", async () => {
    const container = new Elemento("div");
    globalThis.document = { createElement: tag => new Elemento(tag) };
    let liberar;
    const espera = new Promise(resolve => { liberar = resolve; });
    const pagina = { rotate: 0, getViewport: ({ scale }) => ({ width: 200 * scale, height: 300 * scale }),
        render: () => ({ promise: Promise.resolve(), cancel() {} }) };
    const preview = new PDFPreview(container);
    try {
        await preview.adotar({ tarefa: { destroy: async () => {} }, documento: {
            getPage: async numero => { if (numero === 2) await espera; return pagina; }
        } });
        await preview.mostrar(1);
        const canvas = container.querySelector("canvas");
        const antiga = preview.mostrar(2);
        await preview.mostrar(1);
        liberar();
        assert.equal(await antiga, false);
        assert.equal(container.querySelector("canvas"), canvas);
    } finally {
        liberar(); await preview.destruir(); delete globalThis.document;
    }
});

test("redimensionar: prévia limpa, cancelar preserva edições e aplicar reutiliza o PDF.js preparado", async () => {
    const ui = iniciar();
    try {
        ui.selecionar("pdf-arquivo", [await arquivo("resize.pdf", [200, 300])]); await ui.pronto();
        ui.buscar("pdf-ferramenta-redimensionar").click();
        ui.buscar("pdf-redim-largura").value = "100,5"; ui.buscar("pdf-redim-altura").value = "150,25";
        ui.buscar("pdf-redim-prever").click();
        await aguardar(() => !ui.buscar("pdf-redim-aplicar").disabled);
        assert.equal(ui.buscar("pdf-alteracoes-pendentes").hidden, true);
        assert.equal(ui.buscar("pdf-girar-direita").disabled, true);
        assert.match(ui.buscar("pdf-redim-resultado").textContent, /100,5/);
        ui.buscar("pdf-redim-cancelar").click(); await ui.pronto();
        assert.equal(ui.buscar("pdf-alteracoes-pendentes").hidden, true);
        await ui.clicar("pdf-girar-direita");
        ui.buscar("pdf-redim-prever").click(); await aguardar(() => !ui.buscar("pdf-redim-aplicar").disabled);
        ui.buscar("pdf-redim-cancelar").click(); await ui.pronto();
        assert.equal(ui.previews.at(-1).giro, 90);
        assert.equal(ui.buscar("pdf-alteracoes-pendentes").hidden, false);
        ui.buscar("pdf-redim-prever").click(); await aguardar(() => !ui.buscar("pdf-redim-aplicar").disabled);
        const carregamentos = ui.estado.tasks.length;
        ui.buscar("pdf-redim-aplicar").click(); await ui.pronto();
        assert.equal(ui.estado.tasks.length, carregamentos, "Aplicar não recarrega bytes já carregados na prévia");
        assert.equal(ui.previews.at(-1).giro, 0, "A rotação foi incorporada ao conteúdo, sem duplicação");
        assert.match(ui.buscar("pdf-lista-paginas").children[0].children[0].children[0].textContent, /100,5 × 150,25/);
        await ui.clicar("pdf-girar-direita"); await ui.clicar("pdf-mover-baixo");
        assert.equal(ui.previews.at(-1).giro, 90);
        assert.equal(ui.buscar("pdf-pagina-atual").textContent, "2 / 2");
        assert.equal(ui.estado.downloads.length, 0);
    } finally { await ui.destruir(); delete globalThis.document; delete globalThis.window; }
});

test("modo de impressão: diagnóstico, recusa, cancelamento, aplicação e invalidação de consentimento", async () => {
    const confirmar = globalThis.confirm, ui = iniciar(), mensagens = [];
    let aceitar = false;
    globalThis.confirm = mensagem => { mensagens.push(mensagem); return aceitar; };
    try {
        const pdf = await PDFDocument.create(); pdf.addPage([960, 540]); marcarEstrutura(pdf);
        const bytes = await pdf.save();
        ui.selecionar("pdf-arquivo", [{ name: "marcado.pdf", arrayBuffer: async () => bytes.slice().buffer }]); await ui.pronto();
        ui.buscar("pdf-redim-prever").click();
        await aguardar(() => !ui.buscar("pdf-redim-impressao").hidden);
        assert.match(ui.buscar("pdf-redim-resultado").textContent, /acessibilidade/);
        ui.buscar("pdf-redim-impressao").click(); await proximoTurno();
        assert.equal(ui.buscar("pdf-redim-aplicar").disabled, true);
        assert.equal(ui.buscar("pdf-alteracoes-pendentes").hidden, true);
        assert.match(mensagens[0], /INTEIRO/);
        aceitar = true;
        ui.buscar("pdf-redim-impressao").click(); await aguardar(() => !ui.buscar("pdf-redim-aplicar").disabled);
        assert.equal(ui.buscar("pdf-alteracoes-pendentes").hidden, true);
        ui.buscar("pdf-redim-cancelar").click(); await ui.pronto();
        assert.equal(ui.buscar("pdf-alteracoes-pendentes").hidden, true);
        ui.buscar("pdf-redim-prever").click(); await aguardar(() => !ui.buscar("pdf-redim-impressao").hidden);
        await ui.clicar("pdf-girar-direita");
        assert.equal(ui.buscar("pdf-redim-impressao").hidden, true, "Edição invalida diagnóstico/consentimento anterior");
        ui.buscar("pdf-redim-prever").click(); await aguardar(() => !ui.buscar("pdf-redim-impressao").hidden);
        ui.buscar("pdf-redim-impressao").click(); await aguardar(() => !ui.buscar("pdf-redim-aplicar").disabled);
        ui.buscar("pdf-redim-largura").value = "148"; ui.buscar("pdf-redim-largura").emitir("input"); await ui.pronto();
        assert.equal(ui.buscar("pdf-redim-impressao").hidden, true);
        ui.buscar("pdf-redim-prever").click(); await aguardar(() => !ui.buscar("pdf-redim-impressao").hidden);
        ui.buscar("pdf-redim-impressao").click(); await aguardar(() => !ui.buscar("pdf-redim-aplicar").disabled);
        ui.buscar("pdf-redim-aplicar").click(); await ui.pronto();
        assert.equal(ui.buscar("pdf-alteracoes-pendentes").hidden, false);
        assert.equal(ui.downloads.length, 0);
        assert.equal(ui.buscar("pdf-redim-impressao").hidden, true);
        assert.equal(mensagens.length, 4, "Cada preparação exige novo consentimento");
    } finally { globalThis.confirm = confirmar; await ui.destruir(); delete globalThis.document; }
});

test("cancelar durante comparação de ActualText fecha tarefas sem apagar documento ou aprovar descarte", async () => {
    const ui = iniciar();
    try {
        const bytes = await criarFixtureImpressao({ actualText: textoImpressao });
        ui.selecionar("pdf-arquivo", [{ name: "actualtext.pdf", arrayBuffer: async () => bytes.slice().buffer }]); await ui.pronto();
        ui.estado.renderStarted = false; ui.estado.renderGate = new Promise(() => {});
        ui.buscar("pdf-redim-prever").click(); await aguardar(() => ui.estado.renderStarted);
        assert.equal(ui.buscar("pdf-redim-cancelar").disabled, false);
        ui.buscar("pdf-redim-cancelar").click(); await ui.pronto();
        ui.estado.renderGate = null;
        assert.equal(ui.buscar("pdf-total").textContent, 1);
        assert.equal(ui.buscar("pdf-alteracoes-pendentes").hidden, true);
        assert.equal(ui.buscar("pdf-redim-impressao").hidden, true);
        assert.equal(ui.buscar("pdf-redim-aplicar").disabled, true);
        assert.ok(ui.app.querySelector("canvas"));
        assert.ok(ui.tasks.slice(1).every(t => t.destruicoes === 1));
        assert.equal(ui.tasks[0].destruicoes, 0);
        await ui.clicar("pdf-girar-direita");
        assert.equal(ui.buscar("pdf-alteracoes-pendentes").hidden, false);
    } finally { await ui.destruir(); delete globalThis.document; }
});

test("redimensionar: intervalo fora da página atual e invalidação de prévia em andamento", async () => {
    const ui = iniciar();
    try {
        ui.selecionar("pdf-arquivo", [await arquivo("intervalo.pdf", [200, 250, 300])]); await ui.pronto();
        ui.buscar("pdf-redim-alvo").value = "intervalo"; ui.buscar("pdf-redim-alvo").emitir("change");
        ui.buscar("pdf-redim-intervalo").value = "2-3";
        ui.buscar("pdf-redim-largura").value = "100"; ui.buscar("pdf-redim-altura").value = "150";
        ui.buscar("pdf-redim-prever").click(); await aguardar(() => !ui.buscar("pdf-redim-aplicar").disabled);
        assert.equal(ui.buscar("pdf-pagina-atual").textContent, "Prévia: 2 / 3");
        ui.buscar("pdf-redim-largura").value = "110"; ui.buscar("pdf-redim-largura").emitir("input"); await ui.pronto();
        assert.equal(ui.buscar("pdf-redim-aplicar").disabled, true);
        assert.equal(ui.buscar("pdf-pagina-atual").textContent, "1 / 3");
        let liberar;
        ui.estado.renderGate = new Promise(resolve => { liberar = resolve; }); ui.estado.renderStarted = false;
        ui.buscar("pdf-redim-prever").click(); await aguardar(() => ui.estado.renderStarted);
        ui.buscar("pdf-redim-largura").value = "120"; ui.buscar("pdf-redim-largura").emitir("input");
        liberar(); ui.estado.renderGate = null; await ui.pronto();
        assert.equal(ui.buscar("pdf-redim-aplicar").disabled, true);
        assert.equal(ui.buscar("pdf-alteracoes-pendentes").hidden, true);
        ui.buscar("pdf-redim-prever").click(); await aguardar(() => !ui.buscar("pdf-redim-aplicar").disabled);
        ui.buscar("pdf-redim-aplicar").click(); await ui.pronto();
        assert.equal(ui.buscar("pdf-pagina-atual").textContent, "1 / 3");
        assert.equal(ui.buscar("pdf-lista-paginas").children.length, 2);
        assert.match(ui.buscar("pdf-lista-paginas").children[1].children[0].children[0].textContent, /120 × 150.*2 página/);
        assert.match(ui.buscar("pdf-lista-paginas").children[1].children[0].children[1].textContent, /2–3/);
    } finally { await ui.destruir(); delete globalThis.document; delete globalThis.window; }
});

test("redimensionar: recursos não suportados desbloqueiam o editor e não marcam pendência", async () => {
    const ui = iniciar();
    const { PDFName } = await import("../public/vendor/pdf-lib/pdf-lib.js");
    const pdf = await PDFDocument.create();
    const p = pdf.addPage([200, 300]);
    p.node.addAnnot(pdf.context.register(pdf.context.obj({ Type: "Annot", Subtype: "Widget", Rect: [0, 0, 20, 20] })));
    const dados = await pdf.save();
    try {
        ui.selecionar("pdf-arquivo", [{ name: "campo.pdf", arrayBuffer: async () => dados.slice().buffer }]); await ui.pronto();
        ui.buscar("pdf-redim-prever").click();
        await aguardar(() => ui.buscar("pdf-redim-resultado").textContent.includes("Widget"));
        await ui.pronto();
        assert.equal(ui.buscar("pdf-alteracoes-pendentes").hidden, true);
        assert.equal(ui.buscar("pdf-redim-aplicar").disabled, true);
        assert.equal(ui.buscar("pdf-proxima").disabled, true);
    } finally { await ui.destruir(); delete globalThis.document; delete globalThis.window; }
});

test("redimensionar: trocar ferramenta cancela prévia; tamanho igual permite cancelar sem pendência", async () => {
    const ui = iniciar();
    try {
        ui.selecionar("pdf-arquivo", [await arquivo("troca.pdf", [200])]); await ui.pronto();
        ui.buscar("pdf-redim-prever").click(); await aguardar(() => !ui.buscar("pdf-redim-aplicar").disabled);
        ui.buscar("pdf-ferramenta-rotacionar").click(); await ui.pronto();
        assert.equal(ui.buscar("pdf-redim-aplicar").disabled, true);
        assert.equal(ui.buscar("pdf-alteracoes-pendentes").hidden, true);
        ui.buscar("pdf-redim-largura").value = String(200 * 25.4 / 72);
        ui.buscar("pdf-redim-altura").value = String(300 * 25.4 / 72);
        ui.buscar("pdf-redim-prever").click();
        await aguardar(() => ui.buscar("pdf-redim-resultado").textContent.includes("nada a aplicar"));
        assert.equal(ui.buscar("pdf-redim-aplicar").disabled, true);
        assert.equal(ui.buscar("pdf-redim-cancelar").disabled, false);
        ui.buscar("pdf-redim-cancelar").click(); await ui.pronto();
        assert.equal(ui.buscar("pdf-alteracoes-pendentes").hidden, true);
    } finally { await ui.destruir(); delete globalThis.document; delete globalThis.window; }
});

test("cancelar preparação PDF.js não depende de promessa de carregamento concluir e destrói uma vez", async () => {
    const ui = iniciar(), preview = new PDFPreview(new Elemento("div"));
    let destruicoes = 0;
    const carregar = ui.estado.load;
    ui.estado.load = () => ({ promise: new Promise(() => {}), destroy: async () => { ++destruicoes; } });
    try {
        const controle = new AbortController();
        const promessa = preview.preparar(new Uint8Array(), { signal: controle.signal });
        const rejeicao = assert.rejects(promessa, { name: "AbortError" });
        await aguardar(() => preview.tarefas.size === 1);
        controle.abort(); await rejeicao; await preview.destruir();
        assert.equal(destruicoes, 1);
    } finally {
        ui.estado.load = carregar;
        await preview.destruir(); await ui.destruir(); delete globalThis.document; delete globalThis.window;
    }
});

test("falha ao aplicar/restaurar redimensionamento não deixa a prévia descartada como versão definitiva", async () => {
    const ui = iniciar();
    try {
        ui.selecionar("pdf-arquivo", [await arquivo("falha-resize.pdf", [200, 300])]); await ui.pronto();
        ui.buscar("pdf-redim-prever").click(); await aguardar(() => !ui.buscar("pdf-redim-aplicar").disabled);
        ui.estado.failRenderAll = true;
        ui.buscar("pdf-redim-aplicar").click(); await ui.pronto();
        assert.equal(ui.buscar("pdf-alteracoes-pendentes").hidden, true);
        assert.equal(ui.buscar("pdf-redim-aplicar").disabled, true);
        assert.equal(ui.buscar("pdf-visualizacao").querySelector("canvas"), undefined);
        assert.match(ui.buscar("pdf-redim-resultado").textContent, /Falha de render/);
        ui.estado.failRenderAll = false;
        await ui.clicar("pdf-proxima");
        assert.equal(ui.previews.at(-1).numero, 2);
        assert.equal(ui.buscar("pdf-pagina-atual").textContent, "2 / 2");
    } finally { await ui.destruir(); delete globalThis.document; delete globalThis.window; }
});
