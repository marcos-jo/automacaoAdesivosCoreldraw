import fs from 'node:fs/promises';
import path from 'node:path';
import { spawn } from 'node:child_process';
import { createServer } from 'node:http';
import assert from 'node:assert/strict';
import { fileURLToPath } from 'node:url';
import { criarPDFGeracoes } from './helpers/pdf-xref-fixture.mjs';
import { PDFDocument } from '../public/vendor/pdf-lib/pdf-lib.js';
const root = fileURLToPath(new URL('../', import.meta.url)), evidence = await fs.mkdtemp(path.join(root, '.desktop-test/xref-chrome-'));
const comum=await PDFDocument.create();comum.addPage([200,100]).drawText('Comum');
const casos=[{id:'sintetico-xref',bytes:criarPDFGeracoes(),restrito:true}];
if(process.env.PDF_XREF_REAL_PATH)casos.push({id:'real-autorizado',bytes:new Uint8Array(await fs.readFile(process.env.PDF_XREF_REAL_PATH)),restrito:true});
casos.push({id:'comum',bytes:await comum.save(),restrito:false});
const errors = [], network = []; let chrome, socket, serial = 0; const pending = new Map();
const wait = async (fn, label) => { const until = Date.now() + 25000; while (Date.now() < until) { const v = await fn(); if (v) return v; await new Promise(r => setTimeout(r, 80)); } throw Error('Timeout ' + label); };
const cdp = (method, params = {}) => new Promise((resolve, reject) => { const id = ++serial, timer = setTimeout(() => { pending.delete(id); reject(Error('CDP ' + method)); }, 30000); pending.set(id, { resolve: v => { clearTimeout(timer); resolve(v); }, reject: e => { clearTimeout(timer); reject(e); } }); socket.send(JSON.stringify({ id, method, params })); });
const evaluate = async expression => { const r = await cdp('Runtime.evaluate', { expression, awaitPromise: true, returnByValue: true }); if (r.exceptionDetails) throw Error(r.exceptionDetails.exception?.description || r.exceptionDetails.text); return r.result.value; };
const server = createServer(async (req, res) => {
    try {
        if (/^\/case-copy\/\d+$/.test(req.url)) { const c=casos[Number(req.url.split('/').at(-1))];if(!c)throw Error('case');res.setHeader('Content-Type','application/pdf');res.end(c.bytes);return; }
        const name = decodeURIComponent(new URL(req.url, 'http://localhost').pathname), f = path.resolve(root, 'public', name === '/' ? 'index.html' : '.' + name);
        if (!f.startsWith(path.join(root, 'public') + path.sep)) throw Error('path');
        res.setHeader('Content-Type', f.endsWith('.js') ? 'text/javascript' : f.endsWith('.css') ? 'text/css' : f.endsWith('.html') ? 'text/html' : 'application/octet-stream'); res.end(await fs.readFile(f));
    } catch { res.statusCode = 404; res.end(); }
});
await new Promise(r => server.listen(0, '127.0.0.1', r)); const origin = 'http://127.0.0.1:' + server.address().port;
try {
    const profile = path.join(evidence, 'chrome-profile'); await fs.mkdir(profile, { recursive: true });
    chrome = spawn('C:/Program Files (x86)/Google/Chrome/Application/chrome.exe', ['--headless=new', '--disable-gpu', '--disable-background-networking', '--disable-component-update', '--no-first-run', '--no-default-browser-check', '--remote-debugging-port=0', '--user-data-dir=' + profile, 'about:blank'], { windowsHide: true, stdio: 'ignore' });
    const port = await wait(async () => { try { return (await fs.readFile(path.join(profile, 'DevToolsActivePort'), 'utf8')).split('\n')[0]; } catch { return false; } }, 'Chrome');
    const tab = await (await fetch(`http://127.0.0.1:${port}/json/new?about:blank`, { method: 'PUT' })).json();
    socket = new WebSocket(tab.webSocketDebuggerUrl); await new Promise((resolve, reject) => { socket.addEventListener('open', resolve, { once: true }); socket.addEventListener('error', reject, { once: true }); });
    socket.addEventListener('message', e => { const m = JSON.parse(e.data); if (m.id) { const p = pending.get(m.id); if (!p) return; pending.delete(m.id); m.error ? p.reject(Error(m.error.message)) : p.resolve(m.result); }
        if (m.method === 'Runtime.consoleAPICalled' && ['error', 'warn'].includes(m.params.type)) errors.push(m.params.args.map(a => a.description || a.value).join(' '));
        if (m.method === 'Fetch.requestPaused') { const url = m.params.request.url, allow = url.startsWith(origin + '/') || url.startsWith('data:') || url.startsWith('blob:'); network.push({ url, allow }); cdp(allow ? 'Fetch.continueRequest' : 'Fetch.failRequest', allow ? { requestId: m.params.requestId } : { requestId: m.params.requestId, errorReason: 'BlockedByClient' }).catch(() => {}); }
    });
    await cdp('Runtime.enable'); await cdp('Page.enable'); await cdp('Fetch.enable', { patterns: [{ urlPattern: 'http://*' }, { urlPattern: 'https://*' }] });
    await cdp('Page.navigate', { url: origin }); await wait(() => evaluate(`!!document.querySelector('[data-tool="pdfs"]')`), 'menu');
    await evaluate(`document.querySelector('[data-tool="pdfs"]').click()`); await wait(() => evaluate(`!!document.querySelector('#pdf-arquivo')`), 'editor');
    await evaluate(`(async()=>{const {PDFWorkspace}=await import('/js/tools/pdf-workspace.js');const abrir=PDFWorkspace.prototype.abrir;PDFWorkspace.prototype.abrir=async function(...args){const r=await abrir.apply(this,args);window.xrefWorkspace=this;return r;};window.xrefDownloads=0;const click=HTMLAnchorElement.prototype.click;HTMLAnchorElement.prototype.click=function(){if(this.download)xrefDownloads++;else click.call(this);};})()`);
    const resultados=[];
    for(let i=0;i<casos.length;i++){
        const caso=casos[i];
        await evaluate(`(async()=>{const b=await(await fetch('/case-copy/${i}')).arrayBuffer();window.xrefReceived=new Uint8Array(b).slice();const dt=new DataTransfer();dt.items.add(new File([b],'caso.pdf',{type:'application/pdf'}));const input=document.querySelector('#pdf-arquivo');input.files=dt.files;input.dispatchEvent(new Event('change',{bubbles:true}));})()`);
        await wait(()=>evaluate(`!document.querySelector('#pdf-arquivo').disabled && !!document.querySelector('#pdf-visualizacao canvas')`),'PDF visualizado');
        const antes=await evaluate(`({status:document.querySelector('#pdf-status').textContent,revisao:xrefWorkspace.revisao,total:xrefWorkspace.getTotalPaginas(),pdf:xrefWorkspace.pdf.getPages().map(p=>p.node.toString()).join('\\n')})`);
        if(caso.restrito){
            assert.match(antes.status,/aberto para visualização/);
            const bytesIguais=await evaluate(`(async()=>{let n=0;const pdf=xrefWorkspace.pdf,save=pdf.save;pdf.save=function(...args){n++;return save.apply(this,args);};try{const b=await xrefWorkspace.obterBytesVisualizacao();return n===0&&b.length===xrefReceived.length&&b.every((v,i)=>v===xrefReceived[i]);}finally{pdf.save=save;}})()`);assert.equal(bytesIguais,true);
            for(const seletor of ['#pdf-girar-direita','#pdf-recortar','#pdf-salvar-copia']){
                await evaluate(`document.querySelector(${JSON.stringify(seletor)}).click()`);
                await wait(()=>evaluate(`/Operação bloqueada/.test(document.querySelector('#pdf-status').textContent)||!!document.querySelector('#app-exportacao-dialogo')?.open`),'bloqueio');
                if(await evaluate(`!!document.querySelector('#app-exportacao-dialogo')?.open`)){await evaluate(`document.querySelector('#app-exportacao-confirmar').click()`);await wait(()=>evaluate(`/Operação bloqueada/.test(document.querySelector('#pdf-status').textContent)`),'exportação recusada');}
                assert.equal(await evaluate(`xrefWorkspace.revisao`),antes.revisao);
                assert.equal(await evaluate(`xrefWorkspace.pdf.getPages().map(p=>p.node.toString()).join('\\n')`),antes.pdf);
                assert.equal(await evaluate(`xrefDownloads`),0);
            }
            const worker=await evaluate(`(async()=>{const bytes=xrefReceived.slice(),w=new Worker('/js/tools/pdf-import-worker.js',{type:'module'});try{return await new Promise((resolve,reject)=>{const t=setTimeout(()=>reject(Error('worker timeout')),15000);w.onmessage=e=>{clearTimeout(t);resolve(e.data);};w.onerror=e=>{clearTimeout(t);reject(Error(e.message));};w.postMessage({acao:'preparar',nome:'xref.pdf',bytes},[bytes.buffer]);});}finally{w.terminate();}})()`);
            assert.equal(worker.erro?.codigo,'PDF_XREF_INCOMPATIVEL');
            const screen=await cdp('Page.captureScreenshot',{format:'png'});await fs.writeFile(path.join(evidence,caso.id+'.png'),Buffer.from(screen.data,'base64'),{flag:'wx'});
            resultados.push({id:caso.id,total:antes.total,visualizado:true,bytesOriginaisIguais:true,edicaoExportacaoBloqueadas:true,downloads:0,workerBloqueado:true});
        }else{assert.equal(antes.total,1);await evaluate(`document.querySelector('#pdf-girar-direita').click()`);await wait(()=>evaluate(`document.querySelector('#pdf-status').textContent.includes('Página girada')`),'giro comum');const valido=await evaluate(`(async()=>{const bytes=await xrefWorkspace.obterBytes();const pdfjs=await import('/vendor/pdfjs/pdf.js');const task=pdfjs.getDocument({data:bytes});try{const d=await task.promise;await d.getPage(1);return d.numPages===1;}finally{await task.destroy();}})()`);assert.equal(valido,true);resultados.push({id:caso.id,visualizado:true,giroExportacaoReabertura:true});}
    }
    const report={resultados,console:errors,network:{externalBlocked:network.filter(r=>!r.allow).length},ambiente:'Chrome isolado, não Electron instalado'};
    await fs.writeFile(path.join(evidence,'resultado.json'),JSON.stringify(report,null,2),{flag:'wx'});console.log(JSON.stringify(report,null,2));console.log('Evidências: '+evidence);
} finally { if (socket?.readyState === WebSocket.OPEN) { try { await cdp('Browser.close'); } catch {} socket.close(); } if (chrome && chrome.exitCode === null) chrome.kill(); server.close(); }
