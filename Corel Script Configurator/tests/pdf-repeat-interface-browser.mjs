import fs from 'node:fs/promises';
import path from 'node:path';
import { spawn } from 'node:child_process';
import { createServer } from 'node:http';
import assert from 'node:assert/strict';
import { fileURLToPath } from 'node:url';
import { PDFDocument } from '../public/vendor/pdf-lib/pdf-lib.js';
const root = fileURLToPath(new URL('../', import.meta.url)), evidence = await fs.mkdtemp(path.join(root, '.desktop-test/repeat-interface-'));
const phase=process.env.REPEAT_UI_PHASE==='antes'?'antes':'depois';
const comum=await PDFDocument.create();for(const label of ['FRENTE','VERSO']){const p=comum.addPage([90*72/25.4,50*72/25.4]);p.drawText(label+' 90 x 50 mm',{x:12,y:70,size:12});}comum.setCreationDate(new Date('2020-01-01'));comum.setModificationDate(new Date('2020-01-01'));
const casos=[{bytes:await comum.save()}];
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
    chrome = spawn(process.env.BROWSER_PATH || 'C:/Program Files (x86)/Google/Chrome/Application/chrome.exe', ['--headless=new', '--disable-gpu', '--disable-background-networking', '--disable-component-update', '--no-first-run', '--no-default-browser-check', '--remote-debugging-port=0', '--user-data-dir=' + profile, 'about:blank'], { windowsHide: true, stdio: 'ignore' });
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
    await cdp('Emulation.setDeviceMetricsOverride',{width:1366,height:900,deviceScaleFactor:1,mobile:false});
    await evaluate("(async()=>{const {JanelaMontagem}=await import('/js/tools/pdf-imposition-ui.js');const bytes=new Uint8Array(await(await fetch('/case-copy/0')).arrayBuffer()),root=document.createElement('div');document.querySelector('#app').append(root);window.repeatSource=bytes.slice();window.repeatUI=new JanelaMontagem(root,{snapshot:async()=>bytes.slice()});window.confirm=()=>true;window.downloads=[];const create=URL.createObjectURL.bind(URL);window.blobs=new Map();URL.createObjectURL=b=>{const u=create(b);blobs.set(u,b);return u;};const click=HTMLAnchorElement.prototype.click;HTMLAnchorElement.prototype.click=function(){if(this.download)downloads.push({name:this.download,blob:blobs.get(this.href)});else click.call(this);};await repeatUI.open();})()");
    const ready=()=>wait(()=>evaluate('!!repeatUI.result&&!repeatUI.busy&&!repeatUI.dirty'),'prévia vigente');
    const edit=async(id,value)=>{await evaluate("(()=>{const f=repeatUI.get("+JSON.stringify(id)+");f.value="+JSON.stringify(value)+";f.dispatchEvent(new Event('input',{bubbles:true}));})()");await ready();};
    const state=()=>evaluate("(()=>{const u=repeatUI,p=u.result.plan;return{config:u.config(),plan:p,basic:[...u.dialog.querySelectorAll('form input,form select')].filter(e=>e.checkVisibility()).map(e=>e.dataset.id),summary:u.get('medidas-resumo').textContent,status:u.get('status').textContent,source:u.get('origem').textContent,detailsOpen:u.get('repeat-tecnicos')?.open??false,advanced:u.get('avancadas').open,canvas:u.dialog.querySelector('canvas').width,save:u.get('salvar').disabled};})()");
    const screenshot=async name=>{const r=await cdp('Page.captureScreenshot',{format:'png'});await fs.writeFile(path.join(evidence,phase+'-'+name+'.png'),Buffer.from(r.data,'base64'),{flag:'wx'});};
    await ready();const inicial=await state();await screenshot('1366x900');
    await cdp('Emulation.setDeviceMetricsOverride',{width:1024,height:768,deviceScaleFactor:1,mobile:false});await screenshot('1024x768');
    await cdp('Emulation.setDeviceMetricsOverride',{width:720,height:640,deviceScaleFactor:1,mobile:false});await screenshot('720x640');
    assert.equal(await evaluate("repeatUI.dialog.querySelector('footer').getBoundingClientRect().bottom<=innerHeight"),true);
    await cdp('Emulation.setDeviceMetricsOverride',{width:1366,height:900,deviceScaleFactor:1,mobile:false});
    await edit('modo','nup');const nup=await state();await screenshot('nup');await edit('modo','repeat');
    const resultados={phase,inicial,nup};
    if(phase==='depois'){
        assert.deepEqual(inicial.basic,['modo','formato','orientacao','rotacao','colunas','linhas','viradaDuplex']);assert.equal(inicial.detailsOpen,false);assert.equal(inicial.advanced,false);assert.match(inicial.summary,/1 arte por folha.*Escala 100%/s);
        assert.equal(await evaluate("repeatUI.dialog.querySelectorAll('[data-id=rotacao]').length"),1);
        // Parte A: regressão independente do catálogo de formatos salvos.
        await edit('formato','custom');await edit('largura','297');await edit('altura','420');await edit('rotacao','90');
        for(const [orientacao,w,h]of [['paisagem',420,297],['retrato',297,420],['paisagem',420,297]]){
            await edit('orientacao',orientacao);const s=await state();assert.deepEqual([s.config.largura,s.config.altura,s.config.rotacao],[w,h,90]);
            assert.deepEqual([s.plan.folha.largura,s.plan.folha.altura],[w,h]);assert.ok(await evaluate("repeatUI.dialog.querySelector('canvas').width>0"));
            const medidas=await evaluate("(async()=>{const {PDFDocument}=await import('/vendor/pdf-lib/pdf-lib.js');const p=await PDFDocument.load(repeatUI.result.bytes);return p.getPages().map(p=>[p.getWidth()*25.4/72,p.getHeight()*25.4/72]);})()");
            for(const [a,b]of medidas){assert.ok(Math.abs(a-w)<.001);assert.ok(Math.abs(b-h)<.001);}
        }
        await edit('largura','318');await edit('altura','318');await edit('orientacao','paisagem');assert.deepEqual((await state()).plan.folha,{largura:318,altura:318});
        await edit('formato','A3');await edit('orientacao','retrato');await edit('rotacao','0');
        await edit('orientacao','paisagem');const paisagem=await state();assert.equal(paisagem.config.rotacao,0);assert.deepEqual([paisagem.config.largura,paisagem.config.altura],[420,297]);
        await edit('rotacao','90');const girada=await state();assert.deepEqual([girada.config.largura,girada.config.altura],[420,297]);assert.equal(girada.config.rotacao,90);assert.ok(girada.plan.folhas[0].every(p=>p.rotacao===90));
        await edit('colunas','2');await edit('linhas','3');const grade=await state();assert.equal(grade.plan.capacidade,6);assert.match(grade.summary,/6 artes por folha/);
        await edit('linhas','2');await edit('viradaDuplex','longa');const duplex=await state();assert.equal(duplex.plan.viradaDuplex,'longa');assert.equal(duplex.plan.totalFolhas%2,0);assert.match(duplex.summary,/por face/);await edit('viradaDuplex','independente');
        const planoAntesDetalhes=await evaluate('repeatUI.result.plan');
        await evaluate("repeatUI.get('repeat-tecnicos').querySelector('summary').focus()");
        await cdp('Input.dispatchKeyEvent',{type:'rawKeyDown',key:' ',code:'Space',windowsVirtualKeyCode:32});await cdp('Input.dispatchKeyEvent',{type:'keyUp',key:' ',code:'Space',windowsVirtualKeyCode:32});
        assert.equal(await evaluate("repeatUI.get('repeat-tecnicos').open"),true);const semMudanca=await state();assert.deepEqual(semMudanca.plan,planoAntesDetalhes);await screenshot('detalhes');
        await evaluate("repeatUI.get('repeat-tecnicos').querySelector('summary').click()");assert.equal(await evaluate("repeatUI.get('repeat-tecnicos').open"),false);
        await edit('colunas','1');await edit('linhas','1');await edit('rotacao','0');await edit('orientacao','retrato');
        await evaluate("repeatUI.get('proxima').click()");await ready();assert.equal(await evaluate('repeatUI.page'),2);assert.equal(await evaluate("repeatUI.get('salvar').disabled"),false);await evaluate("repeatUI.get('anterior').click()");await ready();assert.equal(await evaluate('repeatUI.page'),1);await evaluate("repeatUI.get('guias').checked=false;repeatUI.get('guias').dispatchEvent(new Event('input',{bubbles:true}))");await ready();assert.equal(await evaluate("repeatUI.get('salvar').disabled"),false);await evaluate('repeatUI.save()');const exportacao=await evaluate("(async()=>{if(downloads.length!==1)throw Error('Download esperado');const bytes=new Uint8Array(await downloads[0].blob.arrayBuffer()),{PDFDocument}=await import('/vendor/pdf-lib/pdf-lib.js'),pdf=await PDFDocument.load(bytes);return{paginas:pdf.getPageCount(),dimensoes:pdf.getPages().map(p=>[p.getWidth(),p.getHeight()])};})()");assert.equal(exportacao.paginas,2);
        const reabertura=await evaluate("(async()=>{const bytes=new Uint8Array(await downloads[0].blob.arrayBuffer()),pdfjs=await import('/vendor/pdfjs/pdf.js'),task=pdfjs.getDocument({data:bytes});try{const pdf=await task.promise;for(let n=1;n<=pdf.numPages;n++)await(await pdf.getPage(n)).getOperatorList();return pdf.numPages;}finally{await task.destroy();}})()");assert.equal(reabertura,2);
        assert.equal(await evaluate('repeatUI.originalBytes.every((v,i)=>v===repeatSource[i])'),true);
        await evaluate("repeatUI.get('avancadas').open=true");await edit('pecaLargura','500');const fora=await state();assert.equal(await evaluate("repeatUI.get('excesso').checkVisibility()"),true);assert.equal(fora.save,true);await screenshot('alerta-corte');
        await edit('pecaLargura','');await edit('colunas','3');await edit('linhas','2');await screenshot('grade');
        resultados.verificacoes={orientacaoIndependente:true,giroIndependente:true,grade:true,duplex:true,teclado:true,exportacao,excessoExigeAceite:true,controleUnico:true};
    }
    await fs.writeFile(path.join(evidence,'estado.json'),JSON.stringify(resultados,null,2),{flag:'wx'});
    const report={resultados,console:errors,network:{externalBlocked:network.filter(r=>!r.allow).length},ambiente:'Chromium isolado, não Electron instalado'};
    await fs.writeFile(path.join(evidence,'resultado.json'),JSON.stringify(report,null,2),{flag:'wx'});console.log(JSON.stringify(report,null,2));console.log('Evidências: '+evidence);
} finally { if (socket?.readyState === WebSocket.OPEN) { try { await cdp('Browser.close'); } catch {} socket.close(); } if (chrome && chrome.exitCode === null) chrome.kill(); server.close(); }
