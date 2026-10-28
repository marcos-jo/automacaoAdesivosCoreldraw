import { Elemento, JanelaTeste } from "./helpers/pdf-dom.mjs";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { setImmediate as proximoTurno } from "node:timers/promises";
import test from "node:test";
import { PDFDocument } from "../public/vendor/pdf-lib/pdf-lib.js";
import * as pdfjs from "../public/vendor/pdfjs/pdf.js";
import { PDFDocumentoVisual } from "../public/js/tools/pdf-visual-document.js";

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
const workspace = await carregar("../public/js/tools/pdf-workspace.js", { "/vendor/pdf-lib/pdf-lib.js": biblioteca });
const editor = await carregar("../public/js/tools/pdf-editor.js", { "/vendor/pdf-lib/pdf-lib.js": biblioteca });
const viewer = modulo(`
    export const GlobalWorkerOptions = { set workerSrc(_) {} };
    export function getDocument(options) { return globalThis.__pdfIntegration.load(options); }
`);
const preview = await carregar("../public/js/tools/pdf-preview.js", { "/vendor/pdfjs/pdf.js": viewer });
const controlador = await carregar("../public/js/tools/pdfs.js", {
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
    const clicar = async id => { buscar(id).click(); await pronto(); await proximoTurno(); };
    const selecionar = (id, arquivos) => { buscar(id).files = arquivos; buscar(id).emitir("change"); };
    const adicionar = async arquivos => {
        buscar("pdf-uniao-abrir").click();
        selecionar("pdf-uniao-arquivos", arquivos);
        await clicar("pdf-uniao-gerar");
    };
    const confirmarSaida = () => destruirAtual.confirmarSaida();
    return { ...estado, estado, app, buscar, pronto, clicar, selecionar, adicionar, destruir, remontar, confirmarSaida };
}

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
        assert.deepEqual(erros, ["Download indisponível"]);
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
