// Benchmark local reproduzível. Uso: node tests/pdf-performance.mjs antes|depois
// Tempos do Chrome headless com PDFs sintéticos; não são limites de aceitação.
import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { access, mkdtemp, readFile, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { setTimeout as pausa } from "node:timers/promises";

const pasta = await mkdtemp(path.join(tmpdir(), "4printshops-performance-"));
let browser, servidor, socket, id = 0;
const pedidos = new Map();
const erros = [];
async function esperar(acao, descricao) {
    const fim = Date.now() + 30000;
    while (Date.now() < fim) {
        const valor = await acao();
        if (valor) return valor;
        await pausa(20);
    }
    throw new Error("Tempo esgotado: " + descricao);
}
function cdp(method, params = {}) {
    return new Promise((resolve, reject) => {
        const numero = ++id;
        const timer = setTimeout(() => { pedidos.delete(numero); reject(new Error(method)); }, 60000);
        pedidos.set(numero, { resolve: r => { clearTimeout(timer); resolve(r); }, reject });
        socket.send(JSON.stringify({ id: numero, method, params }));
    });
}
async function avaliar(expression) {
    const r = await cdp("Runtime.evaluate", { expression, awaitPromise: true, returnByValue: true });
    if (r.exceptionDetails) throw new Error(r.exceptionDetails.exception?.description || r.exceptionDetails.text);
    return r.result.value;
}
try {
    let executavel;
    for (const candidato of [process.env.BROWSER_PATH, "C:/Program Files/Google/Chrome/Application/chrome.exe", "/usr/bin/google-chrome"].filter(Boolean)) {
        try { await access(candidato); executavel = candidato; break; } catch {}
    }
    assert.ok(executavel, "Chrome não encontrado");
    try { assert.ok((await (await fetch("http://localhost:3000")).text()).includes("4PrintShops")); }
    catch {
        servidor = spawn(process.execPath, ["app.js"], { stdio: "ignore", windowsHide: true });
        await esperar(async () => { try { return (await fetch("http://localhost:3000")).ok; } catch { return false; } }, "servidor");
    }
    const perfil = path.join(pasta, "perfil");
    browser = spawn(executavel, ["--headless=new", "--disable-gpu", "--no-first-run", "--no-default-browser-check", "--remote-debugging-port=0", "--user-data-dir=" + perfil, "about:blank"], { stdio: "ignore", windowsHide: true });
    const porta = await esperar(async () => { try { return (await readFile(path.join(perfil, "DevToolsActivePort"), "utf8")).split("\n")[0]; } catch { return false; } }, "Chrome");
    const alvo = await (await fetch(`http://127.0.0.1:${porta}/json/new?about:blank`, { method: "PUT" })).json();
    socket = new WebSocket(alvo.webSocketDebuggerUrl);
    await new Promise(resolve => socket.addEventListener("open", resolve, { once: true }));
    socket.addEventListener("message", e => {
        const m = JSON.parse(e.data);
        if (m.id && pedidos.has(m.id)) {
            const p = pedidos.get(m.id); pedidos.delete(m.id);
            if (m.error) p.reject(new Error(m.error.message)); else p.resolve(m.result);
        }
        if (m.method === "Runtime.exceptionThrown") erros.push(m.params.exceptionDetails.text);
    });
    await cdp("Runtime.enable");
    await cdp("Emulation.setDeviceMetricsOverride", { width: 1440, height: 1000, deviceScaleFactor: 1, mobile: false });
    await cdp("Page.navigate", { url: "http://localhost:3000" });
    await esperar(() => avaliar(`!!document.querySelector('[data-tool="pdfs"]')`), "menu");
    await avaliar(`document.querySelector('[data-tool="pdfs"]').click()`);
    await esperar(() => avaliar(`!!document.querySelector('#pdf-arquivo')`), "editor");
    const resultados = await avaliar(`(async()=>{
        const {PDFDocument,rgb}=await import('/vendor/pdf-lib/pdf-lib.js');
        const {PDFPreview}=await import('/js/tools/pdf-preview.js');
        const {PDFMiniaturas}=await import('/js/tools/pdf-thumbnails.js');
        let medicao=null, inicioDimensoes=null;
        const medir=(alvo,nome,categoria)=>{
            const original=alvo[nome];
            alvo[nome]=async function(...args){
                const m=medicao,t=performance.now();
                try{return await original.apply(this,args);}
                finally{if(m){m[categoria]=(m[categoria]||0)+performance.now()-t;m[categoria+'Chamadas']=(m[categoria+'Chamadas']||0)+1;}}
            };
        };
        medir(PDFDocument.prototype,'save','serializacao');
        medir(PDFDocument,'load','copia');
        medir(PDFPreview.prototype,'mostrar','renderPrincipal');
        medir(PDFMiniaturas.prototype,'processarFila','filaMiniaturas');
        const preparar=PDFPreview.prototype.preparar;
        PDFPreview.prototype.preparar=async function(...args){
            const t=performance.now(),m=medicao;
            const r=await preparar.apply(this,args);
            if(m){m.carregamento=(m.carregamento||0)+performance.now()-t;m.carregamentoChamadas=(m.carregamentoChamadas||0)+1;inicioDimensoes=performance.now();}
            return r;
        };
        const limpar=PDFMiniaturas.prototype.limpar;
        if(PDFPreview.prototype.previsualizar){
            const previsualizar=PDFPreview.prototype.previsualizar;
            PDFPreview.prototype.previsualizar=async function(...args){
                if(medicao&&inicioDimensoes!==null){medicao.dimensoes=(medicao.dimensoes||0)+performance.now()-inicioDimensoes;inicioDimensoes=null;}
                return previsualizar.apply(this,args);
            };
        }
        PDFMiniaturas.prototype.limpar=async function(...args){
            if(medicao&&inicioDimensoes!==null){medicao.dimensoes=(medicao.dimensoes||0)+performance.now()-inicioDimensoes;inicioDimensoes=null;}
            const t=performance.now(),m=medicao;
            try{return await limpar.apply(this,args);}finally{if(m)m.limpeza=(m.limpeza||0)+performance.now()-t;}
        };
        const todosProntos=()=>[...document.querySelectorAll('#pdf-miniaturas button')].every(b=>b.dataset.estado==='pronto'||b.dataset.estado==='erro');
        const esperar=async condicao=>{const fim=performance.now()+30000;while(!condicao()){if(performance.now()>fim)throw new Error('Benchmark não estabilizou');await new Promise(r=>setTimeout(r,2));}};
        const acionar=async acao=>{
            await new Promise((resolve,reject)=>{
                const entrada=document.querySelector('#pdf-arquivo');
                const observador=new MutationObserver(()=>{if(!entrada.disabled){observador.disconnect();resolve();}});
                observador.observe(entrada,{attributes:true,attributeFilter:['disabled']});
                try{acao();if(!entrada.disabled){observador.disconnect();resolve();}}catch(e){observador.disconnect();reject(e);}
            });
        };
        const importar=async arquivo=>{
            const dados=new DataTransfer();dados.items.add(arquivo);
            await acionar(()=>{const e=document.querySelector('#pdf-arquivo');e.files=dados.files;e.dispatchEvent(new Event('change',{bubbles:true}));if(document.querySelector('#pdf-confirmar-descarte').open)document.querySelector('#pdf-descartar-sair').click();});
            await esperar(todosProntos);
        };
        const saida=[];
        for(const total of [20,100,350]){
            const pdf=await PDFDocument.create();
            for(let i=0;i<total;i++){const p=pdf.addPage([400,600]);p.drawText('Pagina '+(i+1),{x:30,y:550,size:24});p.drawRectangle({x:20,y:400,width:350,height:100,color:rgb(.1,.4,.7)});}
            const arquivo=new File([await pdf.save()],total+'-benchmark.pdf',{type:'application/pdf'});
            for(const operacao of ['giro','movimento','exclusao']){
                await importar(arquivo);
                for(let i=0;i<4;i++){document.querySelector('#pdf-proxima').click();await esperar(()=>document.querySelector('#pdf-pagina-atual').textContent.startsWith((i+2)+' /'));}
                for(let amostra=0;amostra<3;amostra++){
                    await esperar(todosProntos);
                    const m=medicao={paginas:total,operacao,amostra};
                    const t=performance.now();
                    await acionar(()=>{
                        if(operacao==='giro')document.querySelector('#pdf-girar-direita').click();
                        if(operacao==='movimento'){document.querySelector('#pdf-mover-destino').value=String(amostra%2?5:Math.min(50,total));document.querySelector('#pdf-mover-aplicar').click();}
                        if(operacao==='exclusao'){document.querySelector('#pdf-selecao-paginas').value=String(total-amostra);document.querySelector('#pdf-excluir').click();}
                    });
                    m.total=performance.now()-t;
                    await esperar(todosProntos);
                    m.estabilizacao=performance.now()-t;
                    m.interface=Math.max(0,m.total-['serializacao','copia','carregamento','dimensoes','renderPrincipal','limpeza'].reduce((s,k)=>s+(m[k]||0),0));
                    medicao=null;saida.push(m);
                }
            }
        }
        return saida;
    })()`);
    assert.deepEqual(erros, []);
    const medianas = [];
    for (const paginas of [20, 100, 350]) for (const operacao of ["giro", "movimento", "exclusao"]) {
        const grupo = resultados.filter(r => r.paginas === paginas && r.operacao === operacao);
        const linha = { paginas, operacao };
        for (const chave of ["total", "estabilizacao", "serializacao", "copia", "carregamento", "dimensoes", "renderPrincipal", "interface", "filaMiniaturas", "serializacaoChamadas", "carregamentoChamadas"]) {
            linha[chave] = +grupo.map(r => r[chave] || 0).sort((a, b) => a - b)[1].toFixed(2);
        }
        medianas.push(linha);
    }
    const arquivo = path.join(pasta, "medicoes.json");
    await writeFile(arquivo, JSON.stringify({ etiqueta: process.argv[2] || "atual", navegador: await cdp("Browser.getVersion"), medianas, amostras: resultados }, null, 2));
    console.log(JSON.stringify(medianas, null, 2));
    console.log("Medições: " + arquivo);
} finally {
    if (socket?.readyState === WebSocket.OPEN) { socket.send(JSON.stringify({ id: ++id, method: "Browser.close" })); socket.close(); }
    if (browser?.exitCode === null) browser.kill();
    if (servidor?.exitCode === null) servidor.kill();
}
