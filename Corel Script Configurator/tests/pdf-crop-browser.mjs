// Validação opcional com Chrome/Edge instalado, sem dependências adicionais.
// Executar: node tests/pdf-crop-browser.mjs (BROWSER_PATH pode indicar o executável).
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
    const resultado=await avaliar(`(async()=>{
        const {PDFDocument,degrees}=await import('/vendor/pdf-lib/pdf-lib.js');
        const {PDFWorkspace}=await import('/js/tools/pdf-workspace.js');
        const {PDFPreview}=await import('/js/tools/pdf-preview.js');
        const {abrirRecorte,caixaDasMargens}=await import('/js/tools/pdf-crop.js');
        const {PDFDocumentoVisual}=await import('/js/tools/pdf-visual-document.js');
        const pt=n=>n*72/25.4, perto=(a,b)=>{if(Math.abs(a-b)>.001)throw Error(a+' != '+b);};
        const resultados=[];
        for(const rotacao of [0,90,180,270])for(const scale of [.5,1,2]){
            const pdf=await PDFDocument.create(),p=pdf.addPage([pt(200),pt(100)]);p.setMediaBox(37,-19,pt(200),pt(100));p.setRotation(degrees(rotacao));p.drawText('Vetor',{x:80,y:30});
            const bytes=await pdf.save(),ws=new PDFWorkspace();await ws.abrir({name:'sintetico.pdf',arrayBuffer:async()=>bytes.slice().buffer});
            const container=document.createElement('div');container.style.width='900px';document.body.append(container);
            const preview=new PDFPreview(container);preview.calcularEscala=()=>scale;
            const candidate=await preview.preparar(bytes);await preview.adotar(candidate);await preview.mostrar(1);
            const viewport=preview.exibicao.viewport,box=caixaDasMargens(viewport,{esquerda:10,direita:10,superior:10,inferior:10});
            perto(box.x,37+pt(10));perto(box.y,-19+pt(10));perto(box.width,pt(180));perto(box.height,pt(80));
            const ready=async()=>{for(let i=0;i<400;i++){const b=document.querySelector('.pdf-recorte-dialog [data-aplicar]');if(b&&!b.disabled)return;await new Promise(r=>setTimeout(r,10));}throw Error('prévia não pronta');};
            const before=ws.revisao;
            let sessao=abrirRecorte({workspace:ws,previa:preview,numero:1,raiz:document.body,aplicar:async(prep,op)=>ws.aplicarRecorte(prep,op)});
            await ready();
            let dialog=document.querySelector('.pdf-recorte-dialog');
            for(const input of dialog.querySelectorAll('.pdf-recorte-campos input'))if(input.name!=='zoom'){input.value=10;input.dispatchEvent(new Event('input'));}
            await ready();
            const esperado=rotacao%180?'80.00 × 180.00':'180.00 × 80.00';if(!dialog.querySelector('[role=status]').textContent.includes(esperado))throw Error('medidas da tela');
            dialog.querySelector('[data-cancelar]').click();await sessao;if(ws.revisao!==before)throw Error('cancelamento mudou documento');
            if (rotacao===0 && scale===.5) {
                sessao=abrirRecorte({workspace:ws,previa:preview,numero:1,raiz:document.body,aplicar:async(prep,op)=>ws.aplicarRecorte(prep,op)});await ready();dialog=document.querySelector('.pdf-recorte-dialog');
                ws.girarPagina(1,180);const revisao=ws.revisao;dialog.querySelector('[data-aplicar]').click();await new Promise(r=>setTimeout(r,50));
                if(!dialog.open || !dialog.querySelector('[role=status]').textContent.includes('vencida') || ws.revisao!==revisao)throw Error('prévia obsoleta aplicada');
                dialog.querySelector('[data-cancelar]').click();await sessao;ws.girarPagina(1,180);
            }
            sessao=abrirRecorte({workspace:ws,previa:preview,numero:1,raiz:document.body,aplicar:async(prep,op)=>ws.aplicarRecorte(prep,op)});await ready();dialog=document.querySelector('.pdf-recorte-dialog');
            for(const input of dialog.querySelectorAll('.pdf-recorte-campos input'))if(input.name!=='zoom'){input.value=10;input.dispatchEvent(new Event('input'));}await ready();
            dialog.querySelector('[data-aplicar]').click();await sessao;
            const reaberto=await PDFDocument.load(await ws.obterBytes());perto(reaberto.getPage(0).getWidth(),pt(180));perto(reaberto.getPage(0).getHeight(),pt(80));
            const visual=await PDFDocumentoVisual.criar((await preview.preparar(await ws.obterBytes())).documento,ws);const dims=visual.obterDimensoes()[0];perto(dims.largura,rotacao%180?80:180);perto(dims.altura,rotacao%180?180:80);
            await preview.destruir();container.remove();resultados.push({rotacao,scale});
        }
        return resultados;
    })()`);
    assert.equal(resultado.length,12);
    console.log('PASSOU: Chrome, 12 combinações de viewport PDF.js real, quatro rotações, origem deslocada, três zooms, margens/tela, aplicar/cancelar e PDF reaberto.');
    await avaliar(`(async()=>{
        const {PDFDocument}=await import('/vendor/pdf-lib/pdf-lib.js');const pdf=await PDFDocument.create();const p=pdf.addPage([200*72/25.4,100*72/25.4]);p.drawText('Teste do botão Recortar',{x:70,y:70});
        const dt=new DataTransfer();dt.items.add(new File([await pdf.save()],'recorte-ui.pdf',{type:'application/pdf'}));const input=document.querySelector('#pdf-arquivo');input.files=dt.files;input.dispatchEvent(new Event('change',{bubbles:true}));
    })()`);
    await pronto();
    await clicarUsuario('#pdf-recortar');
    await esperarTela(`!!document.querySelector('.pdf-recorte-dialog[open] [data-aplicar]:not(:disabled)')`,'Recorte principal');
    await avaliar(`(()=>{const d=document.querySelector('.pdf-recorte-dialog');d.querySelector('[name=zoom]').value=50;d.querySelector('[name=zoom]').dispatchEvent(new Event('input'));})()`);
    const rect=await avaliar(`document.querySelector('.pdf-recorte-dialog canvas').getBoundingClientRect().toJSON()`);
    await cdp('Input.dispatchMouseEvent',{type:'mousePressed',x:rect.x+rect.width*.1,y:rect.y+rect.height*.1,button:'left',clickCount:1});
    await cdp('Input.dispatchMouseEvent',{type:'mouseMoved',x:rect.x+rect.width*.8,y:rect.y+rect.height*.8,buttons:1});
    await cdp('Input.dispatchMouseEvent',{type:'mouseReleased',x:rect.x+rect.width*.8,y:rect.y+rect.height*.8,button:'left',clickCount:1});
    const m=await avaliar(`Object.fromEntries([...document.querySelectorAll('.pdf-recorte-campos input')].map(i=>[i.name,Number(i.value)]))`);
    for(const [key,value]of Object.entries({esquerda:20,direita:40,superior:10,inferior:20}))assert.ok(Math.abs(m[key]-value)<.1, key+': '+m[key]);
    await avaliar(`document.querySelector('.pdf-recorte-dialog [data-cancelar]').click()`);
    await pronto();await clicarUsuario('#pdf-recortar');await esperarTela(`!!document.querySelector('.pdf-recorte-dialog[open] [data-aplicar]:not(:disabled)')`,'Recorte reaberto');
    await avaliar(`(()=>{for(const input of document.querySelectorAll('.pdf-recorte-campos input'))if(input.name!=='zoom'){input.value=10;input.dispatchEvent(new Event('input'));}})()`);
    await esperarTela(`!!document.querySelector('.pdf-recorte-dialog [data-aplicar]:not(:disabled)')`,'Prévia de 10 mm');
    await imagem('recorte-180x80');
    await avaliar(`document.querySelector('.pdf-recorte-dialog [data-aplicar]').click()`);await pronto();
    await esperarTela(`document.querySelector('#pdf-status').textContent.includes('Página recortada')`,'Aplicação principal');
    await clicar('#pdf-salvar-copia');const salvo=await baixar('recorte-ui_editado.pdf',1);assert.ok(Math.abs(salvo.getPage(0).getWidth()*25.4/72-180)<.001);assert.ok(Math.abs(salvo.getPage(0).getHeight()*25.4/72-80)<.001);
    console.log('PASSOU: botão integrado, arraste real a 50%, margens sincronizadas, cancelar, aplicar, exportar e reabrir 180×80 mm. Captura: '+temporario);

    assert.deepEqual(erros, [], 'Erros JavaScript no navegador');
} finally {
    if (socket?.readyState === WebSocket.OPEN) {
        socket.send(JSON.stringify({ id: ++id, method: "Browser.close" }));
        socket.close();
    }
    if (navegador && navegador.exitCode === null) navegador.kill();
    await servidor?.close();
}
