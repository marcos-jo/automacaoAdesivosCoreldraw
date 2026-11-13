import fs from 'node:fs/promises';
import path from 'node:path';
import { spawn } from 'node:child_process';
import { createServer } from 'node:http';
import assert from 'node:assert/strict';
import { fileURLToPath } from 'node:url';
import { criarPDFAssinado } from './helpers/pdf-signature-fixture.mjs';
import { PDFDocument } from '../public/vendor/pdf-lib/pdf-lib.js';
const root = fileURLToPath(new URL('../', import.meta.url)), evidence = await fs.mkdtemp(path.join(root, '.desktop-test/mensagens-chrome-'));
const comum=await PDFDocument.create();comum.addPage([200,100]).drawText('Comum');
const casos=[{id:'comum',tipo:'comum',bytes:await comum.save()},{id:'assinado-sintetico',tipo:'assinatura',bytes:await criarPDFAssinado()}];
if(process.env.PDF_MESSAGES_REAL_LIBRARY){
 const base=path.resolve(process.env.PDF_MESSAGES_REAL_LIBRARY), encontrados=new Map();
 async function localizar(dir){for(const e of await fs.readdir(dir,{withFileTypes:true})){const f=path.join(dir,e.name);if(e.isDirectory())await localizar(f);else if(/^CASO-(003|005|006|011|012|013|029)\.pdf$/.test(e.name))encontrados.set(e.name,f);}}
 await localizar(base);
 for(const [id,tipo,codigo]of [['CASO-011','uniao','PDF_RESTRITO'],['CASO-012','uniao','PDF_RESTRITO'],['CASO-013','uniao','PDF_CAMADAS_INCOMPATIVEIS'],['CASO-003','anotacao','anotacao-montagem'],['CASO-005','recorte'],['CASO-006','recorte'],['CASO-029','xref','PDF_XREF_INCOMPATIVEL']]){
  assert.ok(encontrados.has(id+'.pdf'),'Cópia autorizada ausente: '+id);casos.push({id,tipo,codigo,bytes:new Uint8Array(await fs.readFile(encontrados.get(id+'.pdf')))});
 }
}
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
    await evaluate("Object.defineProperty(navigator,'clipboard',{configurable:true,value:{writeText:async code=>{window.codigoCopiado=code;}}})");
    for(let i=0;i<casos.length;i++){
        const caso=casos[i];
        if(caso.tipo==='uniao'){
            const estado=await evaluate('xrefWorkspace.revisao');
            await evaluate("document.querySelector('#pdf-uniao-abrir').click();document.querySelector('#pdf-uniao-limpar').click()");
            await evaluate("(async()=>{const b=await(await fetch('/case-copy/"+i+"')).arrayBuffer();const dt=new DataTransfer();dt.items.add(new File([b],'caso.pdf',{type:'application/pdf'}));const el=document.querySelector('#pdf-uniao-arquivos');el.files=dt.files;el.dispatchEvent(new Event('change',{bubbles:true}));})()");
            await wait(()=>evaluate("!!document.querySelector('#pdf-uniao-preparacao li[data-falha=\"true\"] button[aria-label^=\"Remover\"]:not([hidden])')"),'recusa de união');
            const r=await evaluate("(()=>{const host=document.querySelector('.pdf-import-state'),d=host.querySelector('details');return {mensagem:host.firstChild.textContent,detalhes:d.querySelector('pre').textContent,recolhido:!d.open};})()");
            assert.match(r.mensagem,caso.tipo==='uniao'&&caso.id==='CASO-013'?/camadas incompatíveis/:/não pode ser unido/);assert.ok(r.detalhes.includes(caso.codigo));assert.equal(r.recolhido,true);
            await evaluate("document.querySelector('.pdf-import-state summary').click();document.querySelector('.pdf-import-state button').click()");
            await wait(()=>evaluate('codigoCopiado==='+JSON.stringify(caso.codigo)),'cópia do código');
            await evaluate("document.querySelector('.pdf-import-state summary').click()");assert.equal(await evaluate('xrefWorkspace.revisao'),estado);
            await evaluate("document.querySelector('#pdf-uniao-preparacao button[aria-label^=\"Remover\"]').click();document.querySelector('#pdf-uniao-fechar').click()");
            resultados.push({id:caso.id,...r,copiaSomenteCodigo:true,estadoPreservado:true});continue;
        }
        if(caso.tipo==='anotacao'){
            const r=await evaluate("(async()=>{const {JanelaMontagem}=await import('/js/tools/pdf-imposition-ui.js');const bytes=new Uint8Array(await(await fetch('/case-copy/"+i+"')).arrayBuffer()),root=document.createElement('div');document.body.append(root);const ui=new JanelaMontagem(root,{snapshot:async()=>bytes.slice()});try{await ui.open();const ate=Date.now()+20000;while(ui.get('compatibilidade').hidden){if(Date.now()>ate)throw Error('timeout diagnóstico');await new Promise(r=>setTimeout(r,50));}return{mensagem:ui.get('compat-resumo').textContent,detalhes:ui.get('compat-detalhes').textContent,semResultado:!ui.result};}finally{await ui.destroy();root.remove();}})()");
            assert.equal(r.mensagem,'Montagem bloqueada por elementos do PDF. Veja os detalhes.');assert.ok(r.detalhes.includes(caso.codigo));assert.equal(r.semResultado,true);resultados.push({id:caso.id,...r});continue;
        }
        await evaluate("(async()=>{const b=await(await fetch('/case-copy/"+i+"')).arrayBuffer();const dt=new DataTransfer();dt.items.add(new File([b],'caso.pdf',{type:'application/pdf'}));const el=document.querySelector('#pdf-arquivo');el.files=dt.files;el.dispatchEvent(new Event('change',{bubbles:true}));})()");
        await wait(()=>evaluate("!document.querySelector('#pdf-arquivo').disabled && !!document.querySelector('#pdf-visualizacao canvas')"),'abertura');
        if(caso.tipo==='assinatura'){
            assert.equal(await evaluate("document.querySelector('#pdf-assinaturas-titulo').textContent"),'PDF assinado. Alterações bloqueadas.');
            await evaluate("document.querySelector('#pdf-assinaturas-resolver').click()");
            assert.equal(await evaluate("document.querySelector('#pdf-assinaturas-expandido').hidden"),false);
            assert.equal(await evaluate("document.querySelector('#pdf-impressao-dialogo').open"),false);
            resultados.push({id:caso.id,restricoesSemPreparo:true});
        }else if(caso.tipo==='recorte'){
            const estado=await evaluate('xrefWorkspace.revisao');
            await evaluate("document.querySelector('#pdf-recortar').click()");
            await wait(()=>evaluate("!!document.querySelector('.pdf-recorte-aviso:not([hidden])')"),'confirmação de sangria');
            assert.match(await evaluate("document.querySelector('.pdf-recorte-aviso').textContent"),/O recorte removerá os ajustes de corte e sangria/);
            await evaluate("document.querySelector('.pdf-recorte-dialog [data-aplicar]').click()");
            assert.equal(await evaluate('xrefWorkspace.revisao'),estado);
            await evaluate("document.querySelector('.pdf-recorte-dialog [data-cancelar]').click()");
            assert.equal(await evaluate('xrefWorkspace.revisao'),estado);resultados.push({id:caso.id,aceiteObrigatorio:true,cancelamentoPreservaEstado:true});
        }else if(caso.tipo==='xref'){
            const estado=await evaluate('xrefWorkspace.revisao');
            assert.match(await evaluate("document.querySelector('#pdf-status').firstChild.textContent"),/Este PDF permite apenas visualização/);
            await evaluate("document.querySelector('#pdf-status summary').click();document.querySelector('#pdf-status button').click()");
            await wait(()=>evaluate('codigoCopiado==='+JSON.stringify(caso.codigo)),'código XRef');
            await evaluate("document.querySelector('#pdf-status summary').click()");assert.equal(await evaluate('xrefWorkspace.revisao'),estado);
            resultados.push({id:caso.id,detalhesECopia:true,estadoPreservado:true});
        }else resultados.push({id:caso.id,aberturaSemCliquesAdicionais:true});
    }
    assert.equal(await evaluate('xrefDownloads'),0);
    const screen=await cdp('Page.captureScreenshot',{format:'png'});await fs.writeFile(path.join(evidence,'mensagem-xref.png'),Buffer.from(screen.data,'base64'),{flag:'wx'});
    const report={resultados,console:errors,network:{externalBlocked:network.filter(r=>!r.allow).length},ambiente:'Chrome isolado, não Electron instalado'};
    await fs.writeFile(path.join(evidence,'resultado.json'),JSON.stringify(report,null,2),{flag:'wx'});console.log(JSON.stringify(report,null,2));console.log('Evidências: '+evidence);
} finally { if (socket?.readyState === WebSocket.OPEN) { try { await cdp('Browser.close'); } catch {} socket.close(); } if (chrome && chrome.exitCode === null) chrome.kill(); server.close(); }
