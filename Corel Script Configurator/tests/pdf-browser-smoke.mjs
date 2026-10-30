// Validação opcional com Chrome/Edge instalado, sem dependências adicionais.
// Executar: node tests/pdf-browser-smoke.mjs (BROWSER_PATH pode indicar o executável).
import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { access, mkdtemp, readFile, readdir, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { setTimeout as pausa } from "node:timers/promises";
import { PDFDocument, PDFName } from "../public/vendor/pdf-lib/pdf-lib.js";
import { criarFixtureImpressao, marcarEstrutura, textoImpressao } from "./helpers/pdf-print-fixture.mjs";
import { impressaoUrl } from "./helpers/pdf-resize-modules.mjs";
import { testarActualTextBrowser } from "./helpers/pdf-actualtext-browser.mjs";
import { testarAssinaturasBrowser } from "./helpers/pdf-signature-browser.mjs";
import { testarUXBrowser } from "./helpers/pdf-ux-browser.mjs";
import { testarAbasBrowser } from "./helpers/app-tabs-browser.mjs";
const { diagnosticarImpressao, prepararParaImpressao } = await import(impressaoUrl);

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
async function contarUnload() {
    const resultado = await cdp("Runtime.evaluate", {
        expression: "(getEventListeners(window).beforeunload || []).length",
        returnByValue: true, includeCommandLineAPI: true
    });
    return resultado.result.value;
}

async function iniciarArraste(numero) {
    arrasteNativo = null;
    const ponto = await avaliar(`(()=>{const e=document.querySelector('#pdf-miniaturas button[data-pagina="${numero}"]');const r=e.getBoundingClientRect();return {x:r.x+r.width/2,y:r.y+r.height/2};})()`);
    await cdp("Input.dispatchMouseEvent", { type: "mouseMoved", ...ponto });
    await cdp("Input.dispatchMouseEvent", { type: "mousePressed", ...ponto, button: "left", clickCount: 1 });
    await cdp("Input.dispatchMouseEvent", { type: "mouseMoved", x: ponto.x + 8, y: ponto.y + 12, button: "left", buttons: 1 });
    await pausa(30);
    await cdp("Input.dispatchMouseEvent", { type: "mouseMoved", x: ponto.x + 10, y: ponto.y + 90, button: "left", buttons: 1 });
    try {
        await aguardar(() => arrasteNativo, "Arraste nativo da miniatura");
    } catch (erro) {
        console.error("Diagnóstico do arraste:", await avaliar(`(()=>{const e=document.querySelector('#pdf-miniaturas button[data-pagina="${numero}"]');return {draggable:e?.draggable,disabled:e?.disabled,retangulo:e?.getBoundingClientRect().toJSON(),alvo:document.elementFromPoint(${ponto.x},${ponto.y})?.outerHTML,status:document.querySelector('#pdf-status')?.textContent};})()`), erros);
        await imagem("falha-arraste");
        console.error("Capturas: " + temporario);
        throw erro;
    }
}

async function soltarApos(numero) {
    await avaliar(`document.querySelector('#pdf-miniaturas button[data-pagina="${numero}"]').scrollIntoView({block:'nearest'})`);
    const ponto = await avaliar(`(()=>{const e=document.querySelector('#pdf-miniaturas button[data-pagina="${numero}"]');const r=e.getBoundingClientRect();return {x:r.x+r.width/2,y:r.y+r.height*0.75};})()`);
    await cdp("Input.dispatchDragEvent", { type: "dragEnter", ...ponto, data: arrasteNativo });
    await cdp("Input.dispatchDragEvent", { type: "dragOver", ...ponto, data: arrasteNativo });
    assert.equal(await avaliar(`!!document.querySelector('.pdf-miniatura-insercao:not([hidden])')`), true);
    await imagem("05-arraste");
    await cdp("Input.dispatchDragEvent", { type: "drop", ...ponto, data: arrasteNativo });
    await pausa(30);
    await cdp("Input.dispatchMouseEvent", { type: "mouseReleased", ...ponto, button: "left", clickCount: 1 });
    await pronto();
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

async function conferirPixelsExportados(nome, numero) {
    const base64 = (await readFile(path.join(temporario, nome))).toString("base64");
    const diferenca = await avaliar(`(async()=>{
        const {getDocument}=await import('/vendor/pdfjs/pdf.js');
        const tarefa=getDocument({data:Uint8Array.from(atob(${JSON.stringify(base64)}),c=>c.charCodeAt(0))});
        try{
            const doc=await tarefa.promise,pagina=await doc.getPage(${numero});
            const atual=document.querySelector('#pdf-visualizacao canvas');
            const esperado=document.createElement('canvas');esperado.width=atual.width;esperado.height=atual.height;
            const base=pagina.getViewport({scale:1});
            await pagina.render({canvas:esperado,canvasContext:esperado.getContext('2d'),viewport:pagina.getViewport({scale:window.viewportPrincipalTeste?.scale??atual.width/base.width})}).promise;
            const a=atual.getContext('2d').getImageData(0,0,atual.width,atual.height).data;
            const b=esperado.getContext('2d').getImageData(0,0,esperado.width,esperado.height).data;
            let diferentes=0;for(let i=0;i<a.length;i++)if(Math.abs(a[i]-b[i])>2)diferentes++;
            esperado.width=esperado.height=0;
            return diferentes/a.length;
        }finally{await tarefa.destroy();}
    })()`);
    assert.ok(diferenca < .001, `Prévia difere do PDF exportado: ${diferenca}`);
}

try {
    let existente;
    try { existente = await (await fetch("http://localhost:3000")).text(); } catch {}
    if (!existente) {
        servidor = spawn(process.execPath, ["app.js"], { cwd: raiz, stdio: "ignore", windowsHide: true });
        await aguardar(async () => {
            try { return (await fetch("http://localhost:3000")).ok; } catch { return false; }
        }, "Servidor local");
    } else assert.ok(existente.includes("4PrintShops"), "A porta 3000 está ocupada por outro aplicativo");

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
    await cdp("Page.navigate", { url: "http://localhost:3000" });
    await esperarTela(`!!document.querySelector('[data-tool="pdfs"]')`, "Menu inicial");
    await clicarUsuario('[data-tool="pdfs"]');
    await esperarTela(`!!document.querySelector('#pdf-ferramenta')`, "Editor");
    if (process.env.APP_TESTE_ABAS_APENAS || !process.env.PDF_TESTE_ASSINATURAS_APENAS && !process.env.PDF_TESTE_UX_APENAS) await testarAbasBrowser({ avaliar, clicar, pronto, esperarTela, imagem, baixar, cdp, contarUnload,
        navegarECancelar: async () => {
            dialogoNativo=null; const tarefa=cdp('Page.navigate',{url:'about:blank'});
            await aguardar(()=>dialogoNativo,'beforeunload com PDF inativo'); assert.equal(dialogoNativo.type,'beforeunload');
            await cdp('Page.handleJavaScriptDialog',{accept:false}); await tarefa;
        }
    });
    if (!process.env.APP_TESTE_ABAS_APENAS && (!process.env.PDF_TESTE_ASSINATURAS_APENAS || process.env.PDF_TESTE_UX_ETAPA)) await testarUXBrowser({ avaliar, clicar, pronto, esperarTela, cdp, imagem, baixar, etapa: Number(process.env.PDF_TESTE_UX_ETAPA || 6),
        confirmar: async (seletor, aceitar) => {
            dialogoNativo = null; const tarefa = clicar(seletor); await aguardar(() => dialogoNativo, 'Confirmação das páginas');
            assert.match(dialogoNativo.message, /Excluir .*página/); await cdp('Page.handleJavaScriptDialog', { accept:aceitar }); await tarefa;
        }
    });
    if (!process.env.APP_TESTE_ABAS_APENAS && !process.env.PDF_TESTE_ASSINATURAS_APENAS && !process.env.PDF_TESTE_UX_APENAS) {
    dialogoNativo = null;
    // O teste de persistência recarrega a aba; restabelecer ativação real do usuário.
    await clicarUsuario('[data-tool="pdfs"]');
    const downloadsAntesRegressao = (await readdir(temporario)).filter(nome => nome.endsWith(".pdf")).length;
    await avaliar(`(async () => {
        const {PDFDocument, rgb} = await import('/vendor/pdf-lib/pdf-lib.js');
        window.criarPDFTeste = async (nome, quantidade, largura) => {
            const pdf = await PDFDocument.create();
            for (let i = 0; i < quantidade; i++) {
                const p = pdf.addPage([largura, 600]);
                p.drawRectangle({x:25,y:420,width:largura-50,height:100,color:rgb(i%2 ? 0.2 : 0.1,0.4,0.7)});
                p.drawText('Pagina '+(i+1), {x:40,y:550,size:24});
            }
            return new File([await pdf.save()],nome,{type:'application/pdf'});
        };
        window.pdfA = await criarPDFTeste('a.pdf',3,400);
        window.pdfB = await criarPDFTeste('b.pdf',2,500);
        window.selecionarPDFTeste = (id, arquivos) => {
            const dados = new DataTransfer(); arquivos.forEach(a => dados.items.add(a));
            const entrada = document.getElementById(id); entrada.files = dados.files;
            entrada.dispatchEvent(new Event('change',{bubbles:true}));
        };
    })()`);

    await clicar("#pdf-ferramenta-unir");
    await clicar("#pdf-uniao-abrir");
    await avaliar(`selecionarPDFTeste('pdf-uniao-arquivos',[pdfA,pdfB]); document.querySelector('#pdf-uniao-nome').value='smoke_unido'`);
    await clicar("#pdf-uniao-gerar");
    await pronto();
    await esperarTela(`document.querySelectorAll('#pdf-miniaturas canvas').length === 5`, "Miniaturas reais");
    assert.equal(await avaliar(`document.querySelector('#pdf-total').textContent`), "5");
    await imagem("01-editor");
    const largura = await avaliar(`document.querySelector('#pdf-visualizacao').clientWidth`);
    await clicar("#pdf-painel-alternar");
    assert.ok(await avaliar(`document.querySelector('#pdf-visualizacao').clientWidth > ${largura}`));
    await clicar("#pdf-painel-alternar");
    await clicar('#pdf-miniaturas button[data-pagina="3"]');
    await esperarTela(`document.querySelector('#pdf-pagina-atual').textContent === '3 / 5'`, "Clique na miniatura");
    await clicar("#pdf-ferramenta-rotacionar");
    await clicar("#pdf-girar-direita");
    await pronto();
    await esperarTela(`(()=>{const c=document.querySelector('#pdf-miniaturas button[data-pagina="3"] canvas'); return c && c.width>c.height;})()`, "Miniatura girada");
    await clicar("#pdf-ferramenta-organizar");
    await avaliar(`document.querySelector('#pdf-selecao-paginas').value='2'`);
    await clicar("#pdf-excluir");
    await pronto();
    assert.equal(await avaliar(`document.querySelector('#pdf-total').textContent`), "4");
    await clicar("#pdf-ferramenta-unir");
    await clicar("#pdf-uniao-abrir");
    await avaliar(`selecionarPDFTeste('pdf-uniao-arquivos',[pdfB])`);
    await clicar("#pdf-uniao-gerar");
    await pronto();
    assert.equal((await readdir(temporario)).filter(nome => nome.endsWith(".pdf")).length, downloadsAntesRegressao, "União não pode baixar arquivos");
    await clicar("#pdf-informacoes");
    assert.equal(await avaliar(`document.querySelector('#pdf-detalhes').hidden`), false);
    await clicar("#pdf-informacoes");
    await clicar("#pdf-salvar-copia");
    const exportado = await baixar("smoke_unido_editado.pdf", 6);
    assert.equal(exportado.getPage(1).getRotation().angle, 90);
    assert.deepEqual(exportado.getPages().map(pagina => pagina.getWidth()), [400,400,500,500,500,500]);
    assert.equal(await avaliar(`document.querySelector('#pdf-alteracoes-pendentes').hidden`), true);
    assert.equal(await contarUnload(), 0);
    await clicar("#pdf-ferramenta-organizar");
    await avaliar(`document.querySelector('#pdf-selecao-paginas').value='1-2'`);
    await clicar("#pdf-extrair");
    await baixar("smoke_unido_extraidas.pdf", 2);
    await pronto();
    await imagem("02-edicoes");

    await avaliar(`(async()=>{window.pdfGrande=await criarPDFTeste('350-paginas.pdf',350,400); selecionarPDFTeste('pdf-arquivo',[pdfGrande]);})()`);
    await pronto();
    assert.equal(await avaliar(`document.querySelector('#pdf-total').textContent`), "350");
    assert.ok(await avaliar(`document.querySelectorAll('#pdf-miniaturas canvas').length <= 16`));
    await avaliar(`const painel=document.querySelector('#pdf-miniaturas');painel.scrollTop=painel.scrollHeight`);
    await esperarTela(`!!document.querySelector('#pdf-miniaturas button[data-pagina="350"] canvas')`, "Última miniatura");
    await clicar('#pdf-miniaturas button[data-pagina="350"]');
    await esperarTela(`document.querySelector('#pdf-pagina-atual').textContent === '350 / 350'`, "Página 350");
    assert.equal(await avaliar(`document.querySelector('#pdf-miniaturas button[data-pagina="350"]').getAttribute('aria-current')`), "true");
    await imagem("03-350-paginas");
    await avaliar(`selecionarPDFTeste('pdf-arquivo',[pdfB])`);
    await pronto();
    await esperarTela(`document.querySelectorAll('#pdf-miniaturas canvas').length === 2`, "Substituição de PDF");

    // Somente visualizar não pede confirmação e clicar no módulo ativo não o recria.
    await clicar('[data-close-tool="pdfs"]');
    await esperarTela(`!document.querySelector('#pdf-ferramenta')`, "Saída sem alterações");
    await clicar('[data-tool="pdfs"]');
    await esperarTela(`!!document.querySelector('#pdf-ferramenta')`, "Volta ao PDF");
    await avaliar(`selecionarPDFTeste('pdf-arquivo',[pdfB])`);
    await pronto();
    await avaliar(`window.canvasOriginal=document.querySelector('#pdf-visualizacao canvas')`);
    await clicar('[data-tool="pdfs"]');
    assert.equal(await avaliar(`window.canvasOriginal === document.querySelector('#pdf-visualizacao canvas')`), true);
    await clicar("#pdf-girar-direita");
    await pronto();
    assert.equal(await contarUnload(), 1);
    assert.equal(await avaliar(`document.querySelector('#pdf-alteracoes-pendentes').hidden`), false);

    // Navegação real da aba aciona beforeunload; cancelar mantém o editor vivo.
    await avaliar(`window.canvasProtegido=document.querySelector('#pdf-visualizacao canvas');window.geometriaProtegida={w:canvasProtegido.width,h:canvasProtegido.height,area:document.querySelector('#pdf-visualizacao').getBoundingClientRect().toJSON()}`);
    const tentativaNativa = cdp("Page.navigate", { url: "about:blank" });
    await aguardar(() => dialogoNativo, "beforeunload nativo");
    assert.equal(dialogoNativo.type, "beforeunload");
    await cdp("Page.handleJavaScriptDialog", { accept: false });
    await tentativaNativa;
    assert.equal(await avaliar(`window.canvasProtegido === document.querySelector('#pdf-visualizacao canvas')`), true, await avaliar(`JSON.stringify({antes:geometriaProtegida,agora:{w:document.querySelector('#pdf-visualizacao canvas').width,h:document.querySelector('#pdf-visualizacao canvas').height,area:document.querySelector('#pdf-visualizacao').getBoundingClientRect().toJSON()}})`));

    // Importação cancelada não muda o documento; aprovação substitui por um PDF limpo.
    await avaliar(`selecionarPDFTeste('pdf-arquivo',[pdfA])`);
    await esperarTela(`document.querySelector('#pdf-confirmar-descarte').open`, "Confirmação de importação");
    await imagem("04-confirmacao");
    await clicar("#pdf-continuar-editando");
    await pronto();
    assert.equal(await avaliar(`document.querySelector('#pdf-nome').textContent`), "b.pdf");
    assert.equal(await avaliar(`window.canvasProtegido === document.querySelector('#pdf-visualizacao canvas')`), true);
    await avaliar(`selecionarPDFTeste('pdf-arquivo',[pdfA])`);
    await esperarTela(`document.querySelector('#pdf-confirmar-descarte').open`, "Aprovação de importação");
    await clicar("#pdf-descartar-sair");
    await pronto();
    assert.equal(await avaliar(`document.querySelector('#pdf-nome').textContent`), "a.pdf");
    assert.equal(await contarUnload(), 0);
    await clicar("#pdf-girar-direita");
    await pronto();
    await clicar("#pdf-salvar-copia");
    await baixar("a_editado.pdf", 3);
    assert.equal(await avaliar(`document.querySelector('#pdf-alteracoes-pendentes').hidden`), true);
    await clicar("#pdf-girar-esquerda");
    await pronto();
    assert.equal(await avaliar(`document.querySelector('#pdf-alteracoes-pendentes').hidden`), false);

    for (const destino of ["inicio", "qrcodes"]) {
        await avaliar(`window.canvasProtegido=document.querySelector('#pdf-visualizacao canvas')`);
        await clicar(`[data-tool="${destino}"]`);
        await esperarTela(`document.querySelector('#app-shell').dataset.activeTool==='${destino}'`, "Alternar sem descarte para " + destino);
        assert.equal(await avaliar(`document.querySelector('#pdf-confirmar-descarte').open`),false);
        assert.equal(await contarUnload(),1);
        await clicar('[data-close-tool="pdfs"]');
        await esperarTela(`document.querySelector('#pdf-confirmar-descarte').open`, "Aviso ao fechar aba");
        assert.equal(await avaliar(`document.querySelector('#app-shell').dataset.activeTool`), "pdfs");
        await clicar("#pdf-continuar-editando");
        await esperarTela(`!document.querySelector('#pdf-confirmar-descarte').open`, "Cancelar saída");
        assert.equal(await avaliar(`window.canvasProtegido === document.querySelector('#pdf-visualizacao canvas')`), true);
        assert.equal(await contarUnload(), 1);
        await clicar(`[data-tool="${destino}"]`);
        await esperarTela(`document.querySelector('#app-shell').dataset.activeTool==='${destino}'`, "Alternância preservada");
        await clicar('[data-close-tool="pdfs"]');
        await esperarTela(`document.querySelector('#pdf-confirmar-descarte').open`, "Confirmar saída");
        await clicar("#pdf-descartar-sair");
        await esperarTela(`!document.querySelector('#pdf-ferramenta')`, "Saída confirmada");
        assert.equal(await avaliar(`!!document.querySelector('[data-close-tool="pdfs"]')`), false);
        assert.equal(await contarUnload(), 0);
        await clicar('[data-tool="pdfs"]');
        await esperarTela(`!!document.querySelector('#pdf-ferramenta')`, "Nova instância");
        await avaliar(`selecionarPDFTeste('pdf-arquivo',[pdfB])`);
        await pronto();
        if (destino !== "qrcodes") { await clicar("#pdf-girar-direita"); await pronto(); }
    }
    for (let i = 0; i < 3; i++) {
        await clicar('[data-close-tool="pdfs"]');
        await esperarTela(`!document.querySelector('#pdf-ferramenta')`, "Encerramento");
        await clicar('[data-tool="pdfs"]');
        await esperarTela(`!!document.querySelector('#pdf-ferramenta')`, "Reentrada");
    }
    assert.equal(await avaliar(`document.querySelectorAll('#pdf-ferramenta').length`), 1);

    // Ordem por controles de teclado e arraste nativo, conservando a página selecionada.
    await cdp("Emulation.setDeviceMetricsOverride", { width: 1440, height: 1250, deviceScaleFactor: 1, mobile: false });
    await avaliar(`(async()=>{
        const {PDFDocument}=await import('/vendor/pdf-lib/pdf-lib.js');
        const pdf=await PDFDocument.create();
        for(let i=0;i<4;i++){const p=pdf.addPage([400+i*10,600]);p.drawText('Ordem '+(i+1),{x:30,y:540,size:28});}
        selecionarPDFTeste('pdf-arquivo',[new File([await pdf.save()],'reordem.pdf',{type:'application/pdf'})]);
    })()`);
    await pronto();
    await esperarTela(`!!document.querySelector('#pdf-miniaturas button[data-pagina="2"] canvas')`, "Miniaturas para reorganizar");
    await clicar('#pdf-miniaturas button[data-pagina="2"]');
    await esperarTela(`document.querySelector('#pdf-pagina-atual').textContent === '2 / 4'`, "Seleção antes de mover");
    await clicar("#pdf-girar-direita");
    await pronto();
    await clicarUsuario("#pdf-mover-baixo");
    await pronto();
    assert.equal(await avaliar(`document.querySelector('#pdf-pagina-atual').textContent`), "3 / 4");
    assert.equal(await avaliar(`document.activeElement.id`), "pdf-mover-baixo");
    await avaliar(`document.querySelector('#pdf-miniaturas').scrollTop=0`);
    await esperarTela(`!!document.querySelector('#pdf-miniaturas button[data-pagina="1"] canvas')`, "Origem do arraste");
    await cdp("Input.setInterceptDrags", { enabled: true });
    await iniciarArraste(1);
    await soltarApos(4);
    assert.equal(await avaliar(`document.querySelector('#pdf-pagina-atual').textContent`), "2 / 4");
    assert.equal(await avaliar(`document.querySelector('#pdf-miniaturas button[aria-current="true"]').dataset.pagina`), "2");
    assert.equal(await avaliar(`document.querySelector('#pdf-alteracoes-pendentes').hidden`), false);
    await clicar("#pdf-salvar-copia");
    const reordenado = await baixar("reordem_editado.pdf", 4);
    assert.deepEqual(reordenado.getPages().map(p => p.getWidth()), [420, 410, 430, 400]);
    assert.equal(reordenado.getPage(1).getRotation().angle, 90);
    // Soltar logo após a própria página não cria outra revisão pendente.
    await iniciarArraste(2);
    await soltarApos(2);
    assert.equal(await avaliar(`document.querySelector('#pdf-alteracoes-pendentes').hidden`), true);

    await avaliar(`selecionarPDFTeste('pdf-arquivo',[pdfGrande])`);
    await pronto();
    const referenciasGrandes = await avaliar(`(async()=>{const {PDFDocument}=await import('/vendor/pdf-lib/pdf-lib.js');const pdf=await PDFDocument.load(await pdfGrande.arrayBuffer());return pdf.getPages().map(p=>p.ref.toString());})()`);
    await esperarTela(`!!document.querySelector('#pdf-miniaturas button[data-pagina="1"] canvas')`, "Origem em 350 páginas");
    await iniciarArraste(1);
    await avaliar(`const painelOrdem=document.querySelector('#pdf-miniaturas');painelOrdem.scrollTop=painelOrdem.scrollHeight`);
    await esperarTela(`!!document.querySelector('#pdf-miniaturas button[data-pagina="350"] canvas')`, "Destino distante");
    assert.equal(await avaliar(`!!document.querySelector('#pdf-miniaturas button[data-pagina="1"]')`), true);
    assert.ok(await avaliar(`document.querySelectorAll('#pdf-miniaturas canvas').length <= 16`));
    await soltarApos(350);
    assert.equal(await avaliar(`document.querySelector('#pdf-pagina-atual').textContent`), "350 / 350");
    assert.ok(await avaliar(`document.querySelectorAll('#pdf-miniaturas canvas').length <= 16`));
    await clicar("#pdf-salvar-copia");
    const grandeReordenado = await baixar("350-paginas_editado.pdf", 350);
    assert.deepEqual(grandeReordenado.getPages().map(p => p.ref.toString()), [...referenciasGrandes.slice(1), referenciasGrandes[0]]);

    // Casos reais de uso: roda do mouse, destino fora da janela e erros/atrasos.
    await cdp("Emulation.setDeviceMetricsOverride", { width: 1440, height: 1000, deviceScaleFactor: 1, mobile: false });
    await avaliar(`(async()=>{
        const pdfjs=await import('/vendor/pdfjs/pdf.js');
        const tarefa=pdfjs.getDocument({data:new Uint8Array(await pdfB.arrayBuffer())});
        const doc=await tarefa.promise;
        const pagina=await doc.getPage(1);
        const proto=Object.getPrototypeOf(pagina);
        const render=proto.render;
        window.controleMiniaturas={modo:'normal',atraso:0,falhas:0,iniciadas:0};
        proto.render=function(opcoes){
            if(opcoes.canvas.className==='pdf-pagina-canvas')window.viewportPrincipalTeste=opcoes.viewport;
            const controle=window.controleMiniaturas;
            if(opcoes.canvas.className!=='pdf-miniatura-canvas')return render.call(this,opcoes);
            controle.iniciadas++;
            if(controle.modo==='erro' && this.pageNumber===2 && controle.falhas++===0){
                return {promise:Promise.reject(new Error('Falha de miniatura injetada no teste')),cancel(){}};
            }
            if(controle.modo==='pendente' && this.pageNumber===1){
                return {promise:new Promise(()=>{}),cancel(){}};
            }
            const tarefaReal=render.call(this,opcoes);
            if(!controle.atraso)return tarefaReal;
            return {promise:tarefaReal.promise.then(()=>new Promise(resolve=>setTimeout(resolve,controle.atraso))),cancel(){tarefaReal.cancel();}};
        };
        await tarefa.destroy();
        window.pdf100=await criarPDFTeste('100-posicoes.pdf',100,400);
        window.pdf350=await criarPDFTeste('350-posicoes.pdf',350,400);
    })()`);
    const estabilizarMiniaturas = () => esperarTela(`Array.from(document.querySelectorAll('#pdf-miniaturas button')).every(b=>b.dataset.estado==='pronto'||b.dataset.estado==='erro')`, "Fila de miniaturas finalizada");
    const moverDireto = async destino => {
        await avaliar(`document.querySelector('#pdf-mover-destino').value=${JSON.stringify(String(destino))}`);
        await clicar("#pdf-mover-aplicar");
        await pronto();
    };
    for (const total of [100,350]) {
        await avaliar(`selecionarPDFTeste('pdf-arquivo',[pdf${total}])`);
        await pronto();
        const referencias = await avaliar(`(async()=>{const {PDFDocument}=await import('/vendor/pdf-lib/pdf-lib.js');const doc=await PDFDocument.load(await pdf${total}.arrayBuffer());return doc.getPages().map(p=>p.ref.toString());})()`);
        await esperarTela(`!!document.querySelector('#pdf-miniaturas button[data-pagina="5"]')`, "Página 5 disponível");
        await clicar('#pdf-miniaturas button[data-pagina="5"]');
        await esperarTela(`document.querySelector('#pdf-pagina-atual').textContent==='5 / ${total}'`, "Página 5 selecionada");
        const sequencia = referencias.slice();
        const moverSequencia = (origem,destino) => sequencia.splice(destino-1,0,...sequencia.splice(origem-1,1));
        await avaliar(`window.canvasAntesDeMover=document.querySelector('#pdf-visualizacao canvas');window.geoMovimento={w:canvasAntesDeMover.width,h:canvasAntesDeMover.height,area:document.querySelector('#pdf-visualizacao').getBoundingClientRect().toJSON()}`);
        await moverDireto(50);
        moverSequencia(5,50);
        assert.equal(await avaliar(`document.querySelector('#pdf-pagina-atual').textContent`), `50 / ${total}`);
        assert.equal(await avaliar(`window.canvasAntesDeMover===document.querySelector('#pdf-visualizacao canvas')`),true,"Reordenar conserva o canvas da página selecionada: " + await avaliar(`JSON.stringify({antes:geoMovimento,w:document.querySelector('#pdf-visualizacao canvas').width,h:document.querySelector('#pdf-visualizacao canvas').height,area:document.querySelector('#pdf-visualizacao').getBoundingClientRect().toJSON()})`));
        assert.equal(await avaliar(`(()=>{const p=document.querySelector('#pdf-miniaturas').getBoundingClientRect();const b=document.querySelector('#pdf-miniaturas button[data-pagina="50"]')?.getBoundingClientRect();return !!b&&b.top>=p.top&&b.bottom<=p.bottom;})()`),true,"Destino selecionado visível antes da rolagem manual");

        const scrolls = await avaliar(`(()=>{const p=document.querySelector('#pdf-miniaturas');const r=p.getBoundingClientRect();return {top:p.scrollTop,corpo:window.scrollY,x:r.x+r.width/2,y:Math.min(r.bottom,innerHeight)-40};})()`);
        await cdp("Input.dispatchMouseEvent", { type: "mouseMoved", x: scrolls.x, y: scrolls.y });
        await cdp("Input.dispatchMouseEvent", { type: "mouseWheel", x: scrolls.x, y: scrolls.y, deltaX: 0, deltaY: 420 });
        await esperarTela(`document.querySelector('#pdf-miniaturas').scrollTop>${scrolls.top}`, "Roda do mouse na lista");
        assert.equal(await avaliar("window.scrollY"), scrolls.corpo);
        assert.equal(await avaliar(`document.querySelector('#pdf-pagina-atual').textContent`), `50 / ${total}`);

        // Operações subsequentes começam com a fila deliberadamente atrasada.
        await avaliar(`controleMiniaturas.atraso=150`);
        for (const destino of [20,75,12,50]) {
            const atual = Number((await avaliar(`document.querySelector('#pdf-pagina-atual').textContent`)).split(' / ')[0]);
            await moverDireto(destino);
            moverSequencia(atual,destino);
        }
        await avaliar(`controleMiniaturas.atraso=0`);
        await estabilizarMiniaturas();
        assert.ok(await avaliar(`document.querySelectorAll('#pdf-miniaturas canvas').length<=16`));
        await clicar("#pdf-girar-direita");
        await pronto();
        const girada = sequencia[49];
        await clicar("#pdf-ferramenta-organizar");
        await avaliar(`document.querySelector('#pdf-selecao-paginas').value='50'`);
        await clicar("#pdf-extrair");
        await pronto();
        const extraidoAtual=await baixar(`${total}-posicoes_extraidas.pdf`,1);
        assert.equal(extraidoAtual.getPage(0).getRotation().angle,90);
        await conferirPixelsExportados(`${total}-posicoes_extraidas.pdf`,1);
        await avaliar(`document.querySelector('#pdf-selecao-paginas').value='2'`);
        await clicar("#pdf-excluir");
        await pronto();
        sequencia.splice(1,1);
        await clicar("#pdf-ferramenta-unir");
        await clicar("#pdf-uniao-abrir");
        await avaliar(`selecionarPDFTeste('pdf-uniao-arquivos',[pdfB])`);
        await clicar("#pdf-uniao-gerar");
        await pronto();
        await estabilizarMiniaturas();
        await clicar("#pdf-salvar-copia");
        const saida = await baixar(`${total}-posicoes_editado.pdf`, total+1);
        assert.deepEqual(saida.getPages().slice(0,total-1).map(p=>p.ref.toString()),sequencia);
        assert.equal(saida.getPage(sequencia.indexOf(girada)).getRotation().angle,90);
        assert.deepEqual(saida.getPages().slice(-2).map(p=>p.getWidth()),[500,500]);
        const paginaAtual=Number((await avaliar(`document.querySelector('#pdf-pagina-atual').textContent`)).split(' / ')[0]);
        await conferirPixelsExportados(`${total}-posicoes_editado.pdf`,paginaAtual);
    }

    // Autoscroll sem alterar scrollTop por JavaScript durante o gesto.
    await avaliar(`selecionarPDFTeste('pdf-arquivo',[pdf100])`);
    await pronto();
    await estabilizarMiniaturas();
    await iniciarArraste(2);
    const borda = await avaliar(`(()=>{const r=document.querySelector('#pdf-miniaturas').getBoundingClientRect();return {x:r.x+r.width/2,y:Math.min(r.bottom,innerHeight)-4};})()`);
    await cdp("Input.dispatchDragEvent", { type: "dragEnter", ...borda, data: arrasteNativo });
    await cdp("Input.dispatchDragEvent", { type: "dragOver", ...borda, data: arrasteNativo });
    await esperarTela(`document.querySelector('#pdf-miniaturas').scrollTop>700`, "Autoscroll inferior nativo");
    const bordaSuperior = await avaliar(`(()=>{const r=document.querySelector('#pdf-miniaturas').getBoundingClientRect();return {x:r.x+r.width/2,y:Math.max(r.top,0)+4};})()`);
    const antes = await avaliar(`document.querySelector('#pdf-miniaturas').scrollTop`);
    await cdp("Input.dispatchMouseEvent", { type: "mouseMoved", ...bordaSuperior, button: "left", buttons: 1 });
    await cdp("Input.dispatchDragEvent", { type: "dragEnter", ...bordaSuperior, data: arrasteNativo });
    await cdp("Input.dispatchDragEvent", { type: "dragOver", ...bordaSuperior, data: arrasteNativo });
    await esperarTela(`document.querySelector('#pdf-miniaturas').scrollTop<${antes}-200`, "Autoscroll superior nativo");
    // No arraste interceptado pelo CDP, atualizar o destino fora da borda antes
    // do cancelamento também encerra a rolagem nativa do Chrome.
    const centro = await avaliar(`(()=>{const r=document.querySelector('#pdf-miniaturas').getBoundingClientRect();return {x:r.x+r.width/2,y:r.y+r.height/2};})()`);
    await cdp("Input.dispatchMouseEvent", { type: "mouseMoved", ...centro, button: "left", buttons: 1 });
    await cdp("Input.dispatchDragEvent", { type: "dragEnter", ...centro, data: arrasteNativo });
    await cdp("Input.dispatchDragEvent", { type: "dragOver", ...centro, data: arrasteNativo });
    await cdp("Input.dispatchDragEvent", { type: "dragCancel", ...centro, data: arrasteNativo });
    await cdp("Input.dispatchMouseEvent", { type: "mouseReleased", ...centro, button: "left", clickCount: 1 });
    await cdp("Input.dispatchMouseEvent", { type: "mouseMoved", x: 800, y: 300 });
    await esperarTela(`!document.querySelector('.pdf-miniatura-arrastando') && !document.querySelector('.pdf-miniatura-insercao')`, "Cancelamento do arraste nativo");

    // Falha isolada e retry sem interromper páginas vizinhas.
    await avaliar(`controleMiniaturas.modo='erro';controleMiniaturas.falhas=0;selecionarPDFTeste('pdf-arquivo',[pdf100])`);
    await pronto();
    await esperarTela(`document.querySelector('#pdf-miniaturas button[data-pagina="2"]')?.dataset.estado==='erro'`, "Erro visível da miniatura");
    await esperarTela(`document.querySelector('#pdf-miniaturas button[data-pagina="3"]')?.dataset.estado==='pronto'`, "Fila prossegue após falha");
    await clicar('#pdf-miniaturas button[data-pagina="2"]');
    await esperarTela(`document.querySelector('#pdf-miniaturas button[data-pagina="2"]')?.dataset.estado==='pronto'`, "Retry de miniatura");
    // Render sem conclusão/cancelamento também não retém uma reorganização.
    await avaliar(`controleMiniaturas.modo='pendente';selecionarPDFTeste('pdf-arquivo',[new File([pdf100],'pendente.pdf',{type:'application/pdf'})])`);
    await pronto();
    await esperarTela(`document.querySelector('#pdf-miniaturas button[data-pagina="1"]')?.dataset.estado==='renderizando'`, "Render pendente injetado");
    await avaliar(`controleMiniaturas.modo='normal'`);
    await moverDireto(50);
    await estabilizarMiniaturas();
    assert.equal(await avaliar(`document.querySelector('#pdf-pagina-atual').textContent`),"50 / 100");
    assert.equal(await avaliar(`(()=>{const painel=document.querySelector('#pdf-miniaturas').getBoundingClientRect();const item=document.querySelector('#pdf-miniaturas button[data-pagina="50"]')?.getBoundingClientRect();return !!item&&item.top>=painel.top&&item.bottom<=painel.bottom;})()`), true, "A página movida deve permanecer visível na lista");
    await imagem("06-mover-posicao");
    await clicar("#pdf-salvar-copia");
    await baixar("pendente_editado.pdf",100);

    // Redimensionamento: conteúdo real, rotação, três modos e pixels da prévia.
    await clicar('#pdf-ferramenta-redimensionar');
    await avaliar(`(async()=>{
        const {PDFDocument,degrees,rgb}=await import('/vendor/pdf-lib/pdf-lib.js');
        const imagem=document.createElement('canvas');imagem.width=imagem.height=16;
        const ctx=imagem.getContext('2d');ctx.fillStyle='#ff00ff';ctx.fillRect(0,0,16,16);ctx.fillStyle='#ffff00';ctx.fillRect(0,0,8,8);
        window.criarResizeTeste=async rotacao=>{
            const pdf=await PDFDocument.create();const p=pdf.addPage([450,650]);
            p.setCropBox(25,35,400,600);p.setRotation(degrees(rotacao));
            p.drawText('Canto superior',{x:50,y:570,size:20});
            p.drawRectangle({x:45,y:55,width:35,height:35,color:rgb(1,0,0)});
            p.drawRectangle({x:350,y:55,width:35,height:35,color:rgb(0,1,0)});
            p.drawRectangle({x:350,y:520,width:35,height:35,color:rgb(0,0,1)});
            p.drawImage(await pdf.embedPng(imagem.toDataURL()),{x:150,y:200,width:100,height:80});
            return new File([await pdf.save()],'resize-'+rotacao+'.pdf',{type:'application/pdf'});
        };
        window.centroides=canvas=>{
            const dados=canvas.getContext('2d').getImageData(0,0,canvas.width,canvas.height).data;
            const cores=[[0,0,0],[0,0,0],[0,0,0]];
            for(let y=0;y<canvas.height;y++)for(let x=0;x<canvas.width;x++){
                const i=(y*canvas.width+x)*4;let c=-1;
                if(dados[i]>240&&dados[i+1]<15&&dados[i+2]<15)c=0;
                if(dados[i]<15&&dados[i+1]>240&&dados[i+2]<15)c=1;
                if(dados[i]<15&&dados[i+1]<15&&dados[i+2]>240)c=2;
                if(c>=0){cores[c][0]+=x;cores[c][1]+=y;cores[c][2]++;}
            }
            return {width:canvas.width,height:canvas.height,cores:cores.map(([x,y,n])=>[x/n,y/n,n])};
        };
    })()`);
    const configurarResize = async (modo='ajustar',alvo='atual',largura='148',altura='210') => avaliar(`
        document.querySelector('#pdf-redim-largura').value=${JSON.stringify(largura)};
        document.querySelector('#pdf-redim-altura').value=${JSON.stringify(altura)};
        document.querySelector('#pdf-redim-largura').dispatchEvent(new Event('input',{bubbles:true}));
        document.querySelector('#pdf-redim-altura').dispatchEvent(new Event('input',{bubbles:true}));
        document.querySelector('#pdf-redim-modo').value=${JSON.stringify(modo)};
        document.querySelector('#pdf-redim-alvo').value=${JSON.stringify(alvo)};
        document.querySelector('#pdf-redim-alvo').dispatchEvent(new Event('change',{bubbles:true}));
    `);
    const resizePronto = () => esperarTela(`!document.querySelector('#pdf-redim-aplicar').disabled`, 'Prévia real de redimensionamento');
    for(const rotacao of [0,90,180,270])for(const modo of ['ajustar','centralizar','preencher']){
        await avaliar(`(async()=>selecionarPDFTeste('pdf-arquivo',[new File([await criarResizeTeste(${rotacao})],'resize-${rotacao}-${modo}.pdf',{type:'application/pdf'})]))()`);await pronto();
        await avaliar(`window.centrosOriginais=centroides(document.querySelector('#pdf-visualizacao canvas'))`);
        await configurarResize(modo);
        await clicar('#pdf-redim-prever');await resizePronto();
        assert.equal(await avaliar(`document.querySelector('#pdf-alteracoes-pendentes').hidden`),true);
        assert.equal(await avaliar(`document.querySelector('#pdf-girar-direita').disabled`),true);
        if(modo==='ajustar'){
            const erroCentros=await avaliar(`(()=>{
                const a=centrosOriginais,b=centroides(document.querySelector('#pdf-visualizacao canvas'));
                const s=Math.min(b.width/a.width,b.height/a.height),dx=(b.width-a.width*s)/2,dy=(b.height-a.height*s)/2;
                return Math.max(...a.cores.flatMap((p,i)=>[Math.abs(b.cores[i][0]-(p[0]*s+dx)),Math.abs(b.cores[i][1]-(p[1]*s+dy))]));
            })()`);
            assert.ok(erroCentros<2,'Rotação/origem/centralização não preservadas: '+erroCentros);
        }
        if(rotacao===0&&modo==='ajustar'){
            await clicar('#pdf-redim-cancelar');await pronto();
            assert.equal(await avaliar(`document.querySelector('#pdf-alteracoes-pendentes').hidden`),true);
            await clicar('#pdf-redim-prever');await resizePronto();
        }
        await avaliar(`(()=>{const c=document.querySelector('#pdf-visualizacao canvas');window.pixelsResize={width:c.width,height:c.height,dados:c.getContext('2d').getImageData(0,0,c.width,c.height).data};})()`);
        await clicar('#pdf-redim-aplicar');await pronto();
        assert.equal(await avaliar(`(()=>{const c=document.querySelector('#pdf-visualizacao canvas'),p=pixelsResize;if(c.width!==p.width||c.height!==p.height)return false;return c.getContext('2d').getImageData(0,0,c.width,c.height).data.every((v,i)=>v===p.dados[i]);})()`),true,'Prévia e aplicação devem coincidir');
        if(rotacao===90&&modo==='ajustar')await imagem('07-redimensionar');
        await clicar('#pdf-salvar-copia');await baixar(`resize-${rotacao}-${modo}_editado.pdf`,1);
        await conferirPixelsExportados(`resize-${rotacao}-${modo}_editado.pdf`,1);
        const pdfRedimensionado=await PDFDocument.load(await readFile(path.join(temporario,`resize-${rotacao}-${modo}_editado.pdf`)));
        assert.ok(Math.abs(pdfRedimensionado.getPage(0).getWidth()-148*72/25.4)<1e-6);
        assert.equal(pdfRedimensionado.getPage(0).getRotation().angle,0);
    }
    await avaliar(`(async()=>{
        const {PDFRedimensionarSessao}=await import('/js/tools/pdf-resize-session.js');
        const prever=PDFRedimensionarSessao.prototype.prever;window.medicoesResize=[];
        PDFRedimensionarSessao.prototype.prever=async function(...args){const t=performance.now();let concluida=false;try{const resultado=await prever.apply(this,args);concluida=true;return resultado;}finally{medicoesResize.push({paginas:args[0].paginas.length,ms:performance.now()-t,cancelada:!concluida});}};
    })()`);
    for(const total of [100,350]){
        await avaliar(`(async()=>selecionarPDFTeste('pdf-arquivo',[await criarPDFTeste('resize-grande-${total}.pdf',${total},400)]))()`);await pronto();
        await configurarResize('ajustar','todas');
        if(total===350){
            await clicar('#pdf-redim-prever');
            await avaliar(`document.querySelector('#pdf-redim-largura').value='150';document.querySelector('#pdf-redim-largura').dispatchEvent(new Event('input',{bubbles:true}))`);
            await pronto();
            assert.equal(await avaliar(`document.querySelector('#pdf-alteracoes-pendentes').hidden`),true);
            await configurarResize('ajustar','todas');
        }
        await avaliar(`window.quadrosResize=0;window.timerResize=setInterval(()=>quadrosResize++,5)`);
        await clicar('#pdf-redim-prever');await resizePronto();
        assert.ok(await avaliar('quadrosResize>0'),'Interface deve ceder tempo aos eventos');
        await avaliar(`clearInterval(timerResize)`);
        await clicar('#pdf-redim-aplicar');await pronto();
        await clicar('#pdf-girar-direita');await pronto();
        await moverDireto(50);
        await clicar('#pdf-ferramenta-organizar');await avaliar(`document.querySelector('#pdf-selecao-paginas').value='2'`);
        await clicar('#pdf-excluir');await pronto();
        await clicar('#pdf-ferramenta-unir');await clicar('#pdf-uniao-abrir');
        await avaliar(`selecionarPDFTeste('pdf-uniao-arquivos',[pdfB])`);await clicar('#pdf-uniao-gerar');await pronto();
        await clicar('#pdf-salvar-copia');const resultado=await baixar(`resize-grande-${total}_editado.pdf`,total+1);
        assert.ok(resultado.getPages().slice(0,total-1).every(p=>Math.abs(p.getWidth()-148*72/25.4)<1e-6));
        assert.equal(resultado.getPage(48).getRotation().angle,90);
        assert.ok(await avaliar(`document.querySelectorAll('#pdf-miniaturas canvas').length<=16`));
    }
    console.log('Prévia de redimensionamento (inclui cópia/save/carregamento/render):',await avaliar('medicoesResize'));

    // Compatibilidade: comparação independente da preparação (SEM geometria).
    await avaliar(`window.compararImpressao=async(a,b)=>{
        const {getDocument}=await import('/vendor/pdfjs/pdf.js');
        const render=async(base64)=>{
            const t=getDocument({data:Uint8Array.from(atob(base64),c=>c.charCodeAt(0))});
            try{
                const p=await(await t.promise).getPage(1),v=p.getViewport({scale:1}),c=document.createElement('canvas');
                c.width=Math.ceil(v.width);c.height=Math.ceil(v.height);
                await p.render({canvas:c,canvasContext:c.getContext('2d'),viewport:v}).promise;
                const pixels=c.getContext('2d').getImageData(0,0,c.width,c.height).data;
                const texto=(await p.getTextContent()).items.map(i=>i.str).join(' ');
                const resultado={width:c.width,height:c.height,pixels,texto};c.width=c.height=0;return resultado;
            }finally{await t.destroy();}
        };
        const x=await render(a),y=await render(b);
        let diferentes=0;for(let i=0;i<x.pixels.length;i++)if(x.pixels[i]!==y.pixels[i])diferentes++;
        return {dimensoes:x.width===y.width&&x.height===y.height,diferentes,total:x.pixels.length,textoAntes:x.texto,textoDepois:y.texto};
    }`);
    for(const opcoes of [{},{navegacao:true,links:true},{navegacao:true,links:true,rotacao:90}]){
        const bytes=await criarFixtureImpressao(opcoes),copia=await PDFDocument.load(bytes);
        prepararParaImpressao(copia,diagnosticarImpressao(copia,[1]));
        const preparado=await copia.save();
        const comparacao=await avaliar(`compararImpressao(${JSON.stringify(Buffer.from(bytes).toString('base64'))},${JSON.stringify(Buffer.from(preparado).toString('base64'))})`);
        assert.equal(comparacao.dimensoes,true);assert.equal(comparacao.diferentes,0,'Preparação mudou pixels antes da geometria');
        assert.equal(comparacao.textoAntes,comparacao.textoDepois);
        assert.ok(comparacao.textoAntes.includes(textoImpressao));
        console.log('Preparação impressão:',JSON.stringify(opcoes),comparacao.total,'componentes RGBA idênticos; texto com símbolos preservado');
    }
    const confirmarImpressao=async (aceitar,marcadores=true)=>{
        dialogoNativo=null;const pendente=clicar('#pdf-redim-impressao');
        await aguardar(()=>dialogoNativo,'Consentimento modo impressão');
        assert.equal(dialogoNativo.type,'confirm');assert.match(dialogoNativo.message,/INTEIRO/);
        assert.match(dialogoNativo.message,/acessibilidade/);if(marcadores)assert.match(dialogoNativo.message,/Marcadores/);
        await cdp('Page.handleJavaScriptDialog',{accept:aceitar});await pendente;
    };
    for(const rotacao of [0,90])for(const modo of ['ajustar','centralizar','preencher']){
        const bytes=await criarFixtureImpressao({navegacao:true,links:true,rotacao});
        const nome=`impressao-${rotacao}-${modo}`;
        await avaliar(`selecionarPDFTeste('pdf-arquivo',[new File([Uint8Array.from(atob(${JSON.stringify(Buffer.from(bytes).toString('base64'))}),c=>c.charCodeAt(0))],${JSON.stringify(nome+'.pdf')},{type:'application/pdf'})])`);
        await pronto();await clicar('#pdf-ferramenta-redimensionar');await configurarResize(modo,'atual','210','297');
        await clicar('#pdf-redim-prever');await esperarTela(`!document.querySelector('#pdf-redim-impressao').hidden`,'Diagnóstico específico');
        if(rotacao===0&&modo==='ajustar'){
            await confirmarImpressao(false);await pronto();
            assert.equal(await avaliar(`document.querySelector('#pdf-alteracoes-pendentes').hidden`),true);
            await confirmarImpressao(true);await resizePronto();
            await clicar('#pdf-redim-cancelar');await pronto();
            assert.equal(await avaliar(`document.querySelector('#pdf-alteracoes-pendentes').hidden`),true);
            await clicar('#pdf-redim-prever');await esperarTela(`!document.querySelector('#pdf-redim-impressao').hidden`,'Consentimento após Cancelar');
        }
        await confirmarImpressao(true);await resizePronto();
        await avaliar(`(()=>{const c=document.querySelector('#pdf-visualizacao canvas');window.pixelsImpressao={width:c.width,height:c.height,dados:c.getContext('2d').getImageData(0,0,c.width,c.height).data};})()`);
        await clicar('#pdf-redim-aplicar');await pronto();
        assert.equal(await avaliar(`(()=>{const c=document.querySelector('#pdf-visualizacao canvas'),p=pixelsImpressao;return c.width===p.width&&c.height===p.height&&c.getContext('2d').getImageData(0,0,c.width,c.height).data.every((v,i)=>v===p.dados[i]);})()`),true);
        assert.equal(await avaliar(`document.querySelector('#pdf-alteracoes-pendentes').hidden`),false);
        await clicar('#pdf-salvar-copia');const exportado=await baixar(nome+'_editado.pdf',1);
        assert.equal(exportado.catalog.has(PDFName.of('StructTreeRoot')),false);
        await conferirPixelsExportados(nome+'_editado.pdf',1);
        assert.equal(await avaliar(`document.querySelector('#pdf-alteracoes-pendentes').hidden`),true);
    }
    // Aparência própria não recebe opção de descarte, mesmo com estrutura marcada.
    const bloqueado=await PDFDocument.create();const paginaBloqueada=bloqueado.addPage([960,540]);marcarEstrutura(bloqueado);
    paginaBloqueada.node.addAnnot(bloqueado.context.register(bloqueado.context.obj({Type:'Annot',Subtype:'Widget',Rect:[0,0,100,30]})));
    const bytesBloqueados=await bloqueado.save();
    await avaliar(`selecionarPDFTeste('pdf-arquivo',[new File([Uint8Array.from(atob(${JSON.stringify(Buffer.from(bytesBloqueados).toString('base64'))}),c=>c.charCodeAt(0))],'impressao-bloqueada.pdf',{type:'application/pdf'})])`);
    await pronto();await configurarResize();await clicar('#pdf-redim-prever');
    await esperarTela(`document.querySelector('#pdf-redim-resultado').textContent.includes('Widget')`,'Bloqueio visual específico');
    assert.equal(await avaliar(`document.querySelector('#pdf-redim-impressao').hidden`),true);
    assert.equal(await avaliar(`document.querySelector('#pdf-alteracoes-pendentes').hidden`),true);
    await pronto();await imagem('08-impressao-bloqueada');

    await testarActualTextBrowser({avaliar,clicar,pronto,esperarTela,configurarResize,resizePronto,confirmarImpressao,baixar,conferirPixelsExportados,imagem,temporario});

    // Biblioteca persistente no perfil de teste; confirmação nativa para excluir.
    await clicar('#pdf-ferramenta-redimensionar');await configurarResize('ajustar','atual','123,45','234,56');
    await avaliar(`document.querySelector('#pdf-formato-nome').value='Formato 4'`);await clicar('#pdf-formato-salvar');
    await esperarTela(`document.querySelector('#pdf-formatos-status').textContent.includes('salvo')`,'Salvar formato');
    await cdp('Page.reload');await esperarTela(`!!document.querySelector('[data-tool="pdfs"]')`,'Recarregar biblioteca');
    await clicar('[data-tool="pdfs"]');await esperarTela(`!!document.querySelector('#pdf-redim-formato')`,'Editor com biblioteca');
    const formatoId=await avaliar(`[...document.querySelector('#pdf-redim-formato').options].find(o=>o.textContent.startsWith('Formato 4')).value`);
    await avaliar(`document.querySelector('#pdf-redim-formato').value=${JSON.stringify(formatoId)};document.querySelector('#pdf-redim-formato').dispatchEvent(new Event('change',{bubbles:true}))`);
    assert.equal(await avaliar(`document.querySelector('#pdf-redim-largura').value`),'123,45');
    await avaliar(`document.querySelector('#pdf-formato-nome').value='Convite personalizado';document.querySelector('#pdf-redim-largura').value='130,5'`);
    await clicar('#pdf-formato-atualizar');
    assert.equal(await avaliar(`JSON.parse(localStorage.getItem('4printshops.pdf.formatos.v1')).formatos[0].nome`),'Convite personalizado');
    for(const aceitar of [false,true]){
        dialogoNativo=null;const excluir=clicar('#pdf-formato-excluir');
        await aguardar(()=>dialogoNativo,'Confirmação de exclusão de formato');assert.equal(dialogoNativo.type,'confirm');
        await cdp('Page.handleJavaScriptDialog',{accept:aceitar});await excluir;
        assert.equal(await avaliar(`JSON.parse(localStorage.getItem('4printshops.pdf.formatos.v1')).formatos.length`),aceitar?0:1);
    }
    }
    const confirmarCopia = async aceitar => {
        dialogoNativo = null; const tarefa = clicar('#pdf-impressao-aplicar');
        await aguardar(() => dialogoNativo, 'Confirmação da cópia de impressão');
        assert.match(dialogoNativo.message, /não preservará a validade da assinatura digital/);
        await cdp('Page.handleJavaScriptDialog', { accept: aceitar }); await tarefa;
    };
    if (!process.env.APP_TESTE_ABAS_APENAS && !process.env.PDF_TESTE_UX_APENAS) await testarAssinaturasBrowser({ avaliar, clicar, pronto, aguardar, esperarTela, confirmar: confirmarCopia, baixar, imagem, temporario });
    assert.deepEqual(erros, [], "Erros JavaScript no navegador");
    console.log(process.env.APP_TESTE_ABAS_APENAS ? "PASSOU: Chrome headless dedicado às abas e menus." : process.env.PDF_TESTE_UX_APENAS ? "PASSOU: Chrome headless dedicado à etapa UX solicitada." : process.env.PDF_TESTE_ASSINATURAS_APENAS ? "PASSOU: Chrome headless, execução dedicada às assinaturas (demais regressões não executadas neste modo)." : "PASSOU: Chrome headless, regressões de PDFs, redimensionamento nos três modos/quatro rotações, 100/350 páginas, biblioteca persistente, ActualText/PowerPoint, assinaturas e cópias para impressão.");
    console.log("Capturas e PDFs sintéticos: " + temporario);
} catch (erro) {
    if (socket?.readyState === WebSocket.OPEN) {
        try {
            console.error("Estado no navegador:", await avaliar(`({status:document.querySelector('#pdf-status')?.textContent,pagina:document.querySelector('#pdf-pagina-atual')?.textContent,scroll:document.querySelector('#pdf-miniaturas')?.scrollTop,miniaturas:[...document.querySelectorAll('#pdf-miniaturas button')].map(b=>[b.dataset.pagina,b.dataset.estado]),erros:document.querySelectorAll('.pdf-miniatura-erro').length})`));
            await imagem("falha");
            console.error("Capturas: " + temporario);
        } catch {}
    }
    throw erro;
} finally {
    if (socket?.readyState === WebSocket.OPEN) {
        socket.send(JSON.stringify({ id: ++id, method: "Browser.close" }));
        socket.close();
    }
    if (navegador && navegador.exitCode === null) navegador.kill();
    if (servidor && servidor.exitCode === null) servidor.kill();
}
