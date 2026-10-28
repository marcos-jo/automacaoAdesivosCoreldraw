import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { setImmediate as proximoTurno } from "node:timers/promises";
import test from "node:test";
import { PDFDocument } from "../public/vendor/pdf-lib/pdf-lib.js";
import * as pdfjs from "../public/vendor/pdfjs/pdf.js";

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
const controlador = await carregar("../public/js/tools/pdfs.js", {
    "./pdf-workspace.js": workspace,
    "./pdf-editor.js": editor,
    "./pdf-merge.js": url("../public/js/tools/pdf-merge.js"),
    "/vendor/pdfjs/pdf.js": viewer
});
const { renderPDFs } = await import(controlador);
pdfjs.GlobalWorkerOptions.workerSrc = url("../public/vendor/pdfjs/pdf.worker.js");

// DOM mínimo para acionar os módulos reais. A rasterização do canvas é substituída;
// leitura, dimensões e extração usam PDF.js/pdf-lib locais com bytes PDF reais.
class Elemento {
    constructor(tag) {
        this.tag = tag;
        this.children = [];
        this.listeners = new Map();
        this.dataset = {};
        this.value = "";
        this.textContent = "";
        this.classList = { toggle: (classe, ativo) => { this[classe] = ativo; } };
    }
    append(...filhos) { for (const filho of filhos) { filho.parent = this; this.children.push(filho); } }
    appendChild(filho) { this.append(filho); }
    replaceChildren(...filhos) { this.children = []; this.append(...filhos); }
    before(filho) { filho.parent = this.parent; this.parent.children.splice(this.parent.children.indexOf(this), 0, filho); }
    remove() { this.parent.children = this.parent.children.filter(filho => filho !== this); }
    set innerHTML(html) {
        this.replaceChildren();
        const pilha = [this];
        for (const [, fecha, tag, atributos] of html.matchAll(/<(\/)?([\w-]+)([^>]*)>/g)) {
            if (fecha) { pilha.pop(); continue; }
            const filho = new Elemento(tag);
            for (const [, nome, valor] of atributos.matchAll(/([\w-]+)="([^"]*)"/g)) filho[nome] = valor;
            filho.disabled = /\bdisabled\b/.test(atributos);
            filho.hidden = /\bhidden\b/.test(atributos);
            pilha.at(-1).append(filho);
            if (!["input", "br", "hr"].includes(tag)) pilha.push(filho);
        }
    }
    matches(seletor) {
        if (seletor === "[id]") return !!this.id;
        if (seletor.startsWith("#")) return this.id === seletor.slice(1);
        return this.tag === "button" && (seletor === "button" || this.dataset.acao !== undefined);
    }
    querySelectorAll(seletor) {
        return this.children.flatMap(filho => [
            ...(filho.matches(seletor) ? [filho] : []), ...filho.querySelectorAll(seletor)
        ]);
    }
    querySelector(seletor) { return this.querySelectorAll(seletor)[0]; }
    closest(seletor) { return this.matches(seletor) ? this : this.parent?.closest(seletor); }
    addEventListener(tipo, handler) {
        if (!this.listeners.has(tipo)) this.listeners.set(tipo, []);
        this.listeners.get(tipo).push(handler);
    }
    emitir(tipo, alvo = this) {
        for (let atual = this; atual; atual = atual.parent) {
            for (const handler of atual.listeners.get(tipo) || []) handler({ target: alvo, preventDefault() {} });
        }
    }
    click() {
        if (this.disabled) return;
        if (this.tag === "a") {
            globalThis.__pdfIntegration.downloads.push({ nome: this.download, blob: this.hrefBlob });
        } else this.emitir("click");
    }
    set href(valor) { this.hrefBlob = globalThis.__pdfIntegration.blobs.get(valor); }
    showModal() { this.open = true; }
    close() { this.open = false; }
    getContext() { return {}; }
}

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
                        pagina.render = ({ viewport }) => {
                            this.previews.push({ numero, giro: viewport.rotation });
                            this.renderStarted = true;
                            return { promise: this.renderGate || Promise.resolve() };
                        };
                        return pagina;
                    };
                    return doc;
                })
            };
        }
    };
    renderPDFs();
    const buscar = id => app.querySelector("#" + id);
    const pronto = () => aguardar(() => !buscar("pdf-arquivo").disabled);
    const clicar = async id => { buscar(id).click(); await pronto(); await proximoTurno(); };
    const selecionar = (id, arquivos) => { buscar(id).files = arquivos; buscar(id).emitir("change"); };
    const adicionar = async arquivos => {
        buscar("pdf-uniao-abrir").click();
        selecionar("pdf-uniao-arquivos", arquivos);
        await clicar("pdf-uniao-gerar");
    };
    return { ...estado, estado, app, buscar, pronto, clicar, selecionar, adicionar };
}

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
    globalThis.setTimeout = callback => { callback(); return 0; };
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
        ui.app.innerHTML = "";
        renderPDFs();
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
        delete globalThis.document;
    }
});
