import {testarAssinaturasBrowser} from './helpers/pdf-signature-browser.mjs';
// Validação opcional com Chrome/Edge instalado, sem dependências adicionais.
// Executar: node tests/pdf-signature-policy-browser.mjs (BROWSER_PATH pode indicar o executável).
import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { access, mkdtemp, readFile, readdir, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { setTimeout as pausa } from "node:timers/promises";
import { PDFDocument } from "../public/vendor/pdf-lib/pdf-lib.js";
const raiz = fileURLToPath(new URL("../", import.meta.url));
const temporario = await mkdtemp(path.join(tmpdir(), "4printshops-browser-"));
const caminhos = [process.env.BROWSER_PATH,
    "C:/Program Files/Google/Chrome/Application/chrome.exe",
    "C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe",
    "/usr/bin/google-chrome", "/usr/bin/chromium"].filter(Boolean);
let executavel;
for (const candidato of caminhos) {
    try { await access(candidato); executavel = candidato; break; } catch {}
}
assert.ok(executavel, "Chrome/Edge não encontrado. Informe BROWSER_PATH.");
let servidor;
let navegador;
let socket;
let id = 0;
const pedidos = new Map();
const erros = [];
let dialogoNativo;
let arrasteNativo;
async function aguardar(acao, descricao, tempo = 20000) {
    const limite = Date.now() + tempo;
    while (Date.now() < limite) {
        const valor = await acao();
        if (valor) return valor;
        await pausa(50);
    }
    throw new Error("Tempo esgotado: " + descricao);
}
function cdp(method, params = {}) {
    const atual = ++id;
    return new Promise((resolve, reject) => {
        const timer = setTimeout(() => {
            pedidos.delete(atual);
            reject(new Error("Comando do navegador não respondeu: " + method));
        }, 30000);
        pedidos.set(atual, {
            resolve: valor => { clearTimeout(timer); resolve(valor); },
            reject: erro => { clearTimeout(timer); reject(erro); }
        });
        socket.send(JSON.stringify({ id: atual, method, params }));
    });
}
async function avaliar(expression) {
    const resposta = await cdp("Runtime.evaluate", { expression, awaitPromise: true, returnByValue: true });
    if (resposta.exceptionDetails) throw new Error(resposta.exceptionDetails.exception?.description || resposta.exceptionDetails.text);
    return resposta.result.value;
}
const esperarTela = (expressao, descricao) => aguardar(() => avaliar(expressao), descricao);
const clicar = async seletor => {
    if (seletor === '#pdf-uniao-gerar') await esperarTela(`!document.querySelector('#pdf-uniao-gerar').disabled`, 'Análise da combinação');
    await avaliar(`document.querySelector(${JSON.stringify(seletor)}).click()`);
    // Cenários que já solicitavam um arquivo agora confirmam a janela real.
    if (seletor === '#pdf-salvar-copia') {
        await esperarTela(`document.querySelector('#app-exportacao-dialogo').open`, 'Janela de exportação');
        await avaliar(`document.querySelector('#app-exportacao-confirmar').click()`);
    }
};
async function clicarUsuario(seletor) {
    const ponto = await avaliar(`(()=>{const r=document.querySelector(${JSON.stringify(seletor)}).getBoundingClientRect();return {x:r.x+r.width/2,y:r.y+r.height/2};})()`);
    await cdp("Input.dispatchMouseEvent", { type: "mousePressed", ...ponto, button: "left", clickCount: 1 });
    await cdp("Input.dispatchMouseEvent", { type: "mouseReleased", ...ponto, button: "left", clickCount: 1 });
}
const pronto = () => esperarTela(`!document.querySelector('#pdf-arquivo').disabled && !!document.querySelector('#pdf-visualizacao canvas')`, "PDF pronto");
async function imagem(nome) {
    const { data } = await cdp("Page.captureScreenshot", { format: "png", captureBeyondViewport: true });
    await writeFile(path.join(temporario, nome + ".png"), Buffer.from(data, "base64"));
}
async function baixar(nome, paginas) {
    const arquivo = path.join(temporario, nome);
    await aguardar(async () => {
        try { return (await readFile(arquivo)).length > 0; } catch { return false; }
    }, "Download " + nome);
    const pdf = await PDFDocument.load(await readFile(arquivo));
    assert.equal(pdf.getPageCount(), paginas);
    return pdf;
}

const {createServer}=await import('node:http');
servidor=createServer(async(req,res)=>{try{const relative=req.url==='/'?'index.html':decodeURIComponent(req.url.split('?')[0].slice(1));const file=path.resolve(raiz,'public',relative);if(!file.startsWith(path.join(raiz,'public')+path.sep))throw Error('path');res.setHeader('Content-Type',file.endsWith('.js')?'text/javascript':file.endsWith('.css')?'text/css':file.endsWith('.html')?'text/html':'application/octet-stream');res.end(await readFile(file));}catch{res.statusCode=404;res.end();}});
await new Promise(resolve=>servidor.listen(0,'127.0.0.1',resolve));const origem='http://127.0.0.1:'+servidor.address().port;
try {
    const perfil = path.join(temporario, "perfil");
    navegador = spawn(executavel, ["--headless=new", "--disable-gpu", "--no-first-run",
        "--no-default-browser-check", "--remote-debugging-port=0", "--user-data-dir=" + perfil,
        "about:blank"], { stdio: "ignore", windowsHide: true });
    const porta = await aguardar(async () => {
        try { return (await readFile(path.join(perfil, "DevToolsActivePort"), "utf8")).split("\n")[0]; } catch { return false; }
    }, "Chrome headless");
    const alvo = await (await fetch(`http://127.0.0.1:${porta}/json/new?about:blank`, { method: "PUT" })).json();
    socket = new WebSocket(alvo.webSocketDebuggerUrl);
    await new Promise((resolve, reject) => { socket.addEventListener("open", resolve, { once: true }); socket.addEventListener("error", reject, { once: true }); });
    socket.addEventListener("message", evento => {
        const mensagem = JSON.parse(evento.data);
        if (mensagem.id) {
            const pedido = pedidos.get(mensagem.id);
            if (!pedido) return;
            pedidos.delete(mensagem.id);
            if (mensagem.error) pedido.reject(new Error(mensagem.error.message));
            else pedido.resolve(mensagem.result);
        }
        if (mensagem.method === "Runtime.exceptionThrown") erros.push(mensagem.params.exceptionDetails.text);
        if (mensagem.method === "Page.javascriptDialogOpening") dialogoNativo = mensagem.params;
        if (mensagem.method === "Input.dragIntercepted") arrasteNativo = mensagem.params.data;
        if (mensagem.method === "Runtime.consoleAPICalled" && mensagem.params.type === "error") {
            erros.push(mensagem.params.args.map(arg => arg.description || arg.value).join(" "));
        }
    });
    await cdp("Runtime.enable");
    await cdp("Page.enable");
    await cdp("Browser.setDownloadBehavior", { behavior: "allow", downloadPath: temporario });
    await cdp("Emulation.setDeviceMetricsOverride", { width: 1440, height: 1000, deviceScaleFactor: 1, mobile: false });
    await cdp("Page.navigate", { url: origem });
    await esperarTela(`!!document.querySelector('[data-tool="pdfs"]')`, "Menu inicial");
    await clicarUsuario('[data-tool="pdfs"]');
    await esperarTela(`!!document.querySelector('#pdf-ferramenta')`, "Editor");
    await testarAssinaturasBrowser({avaliar,pronto,esperarTela,imagem});
    assert.deepEqual(erros.filter(e=>!/assinatura digital|PDF_ASSINATURA_PROTEGIDA/.test(e)),[]);
} finally {
    if (socket?.readyState === WebSocket.OPEN) {
        socket.send(JSON.stringify({ id: ++id, method: "Browser.close" }));
        socket.close();
    }
    if (navegador && navegador.exitCode === null) navegador.kill();
    await servidor?.close();
}
