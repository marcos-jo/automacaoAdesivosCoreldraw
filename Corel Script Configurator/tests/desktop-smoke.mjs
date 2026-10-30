import assert from 'node:assert/strict';
import {spawn} from 'node:child_process';
import {mkdtemp,readFile,writeFile,unlink} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {setTimeout as pause} from 'node:timers/promises';
import {PDFDocument,PDFName} from '../public/vendor/pdf-lib/pdf-lib.js';
import {criarPDFAssinado} from './helpers/pdf-signature-fixture.mjs';

const root=fileURLToPath(new URL('../',import.meta.url)), folder=await mkdtemp(path.join(tmpdir(),'4printshops-desktop-'));
let child,socket,id=0,serial=0,options={},debug='',jsDialog; const requests=new Map(),errors=[];
const wait=async(fn,label,limit=30000)=>{const end=Date.now()+limit;while(Date.now()<end){const r=await fn();if(r)return r;await pause(50);}throw new Error('Tempo esgotado: '+label);};
const command=async(data)=>{options={...options,...data,serial:++serial};await writeFile(path.join(folder,'command.json'),JSON.stringify(options));await wait(async()=>{if(child.exitCode!==null)return true;try{return JSON.parse(await readFile(path.join(folder,'status.json'),'utf8')).serial===serial;}catch{return false;}},'Comando desktop');};
const cdp=(method,params={})=>new Promise((resolve,reject)=>{const n=++id,t=setTimeout(()=>{requests.delete(n);reject(new Error('CDP: '+method));},30000);requests.set(n,{resolve:v=>{clearTimeout(t);resolve(v);},reject});socket.send(JSON.stringify({id:n,method,params}));});
const evaluate=async(expression)=>{const r=await cdp('Runtime.evaluate',{expression,awaitPromise:true,returnByValue:true,userGesture:true});if(r.exceptionDetails)throw new Error(r.exceptionDetails.exception?.description||r.exceptionDetails.text);return r.result.value;};
const click=async s=>evaluate(`document.querySelector(${JSON.stringify(s)}).click()`);
const ready=()=>wait(()=>evaluate(`!!document.querySelector('#pdf-arquivo')&&!document.querySelector('#pdf-arquivo').disabled`),'PDF pronto');
async function launch(){
    await unlink(path.join(folder,'profile','DevToolsActivePort')).catch(()=>{});
    const env={...process.env,PRINTSHOPS_TEST_FOLDER:folder}; delete env.ELECTRON_RUN_AS_NODE;
    child=spawn(path.join(root,'node_modules/electron/dist/electron.exe'),['tests/helpers/desktop-entry.mjs','--remote-debugging-port=0'],{cwd:root,env,windowsHide:false,stdio:['ignore','pipe','pipe']});
    child.stderr.on('data',d=>debug+=d);child.stdout.on('data',d=>debug+=d);
    await writeFile(path.join(folder,'command.json'),JSON.stringify({serial:0}));
    const port=await wait(async()=>{try{return (await readFile(path.join(folder,'profile','DevToolsActivePort'),'utf8')).split('\n')[0];}catch{return false;}},'Electron iniciou');
    const target=await wait(async()=>{try{return (await(await fetch(`http://127.0.0.1:${port}/json`)).json()).find(t=>t.type==='page'&&t.url.startsWith('app://4printshops'));}catch{return false;}},'Página desktop');
    socket=new WebSocket(target.webSocketDebuggerUrl);await new Promise(r=>socket.addEventListener('open',r,{once:true}));
    socket.addEventListener('message',e=>{const m=JSON.parse(e.data);if(m.id){requests.get(m.id)?.resolve(m.result);requests.delete(m.id);}if(m.method==='Runtime.exceptionThrown')errors.push(m.params.exceptionDetails.text);if(m.method==='Page.javascriptDialogOpening')jsDialog=m.params;});
    await cdp('Runtime.enable');await cdp('Page.enable');
    await wait(()=>evaluate(`!!document.querySelector('[data-tool="pdfs"]')`),'Interface pronta');
}
try {
    await launch();
    assert.equal(await evaluate(`typeof require`),'undefined');assert.equal(await evaluate(`typeof process`),'undefined');
    const status=await wait(async()=>{try{return JSON.parse(await readFile(path.join(folder,'status.json'),'utf8'));}catch{return false;}},'Estado do processo principal');
    assert.equal(status.preferences.sandbox,true);assert.equal(status.preferences.contextIsolation,true);assert.equal(status.preferences.nodeIntegration,false);
    await click('[data-tool="pdfs"]');await ready();
    await evaluate(`(async()=>{const{PDFDocument}=await import('/vendor/pdf-lib/pdf-lib.js'),p=await PDFDocument.create();for(let i=0;i<20;i++){const a=p.addPage([600,900]);a.drawText('Pagina '+(i+1),{x:30,y:700});}window.bytesDesktop=await p.save();const d=new DataTransfer();d.items.add(new File([bytesDesktop],'desktop.pdf'));const e=document.querySelector('#pdf-arquivo');e.files=d.files;e.dispatchEvent(new Event('change',{bubbles:true}));})()`);
    await ready();await wait(()=>evaluate(`!!document.querySelector('#pdf-visualizacao canvas')`),'PDF.js com worker no protocolo');
    await click('#pdf-girar-direita');await ready();
    assert.equal(await evaluate(`document.querySelector('#pdf-alteracoes-pendentes').hidden`),false);
    await command({action:'close',answer:0});await pause(200);assert.equal(child.exitCode,null,'Cancelar mantém o app aberto');
    await click('[data-tool="inicio"]');await click('[data-tool="qrcodes"]');await click('[data-tool="pdfs"]');
    await command({action:'save-options',save:{canceled:true}});
    await click('#pdf-salvar-copia');await click('#app-exportacao-confirmar');await ready();
    assert.equal(await evaluate(`document.querySelector('#pdf-alteracoes-pendentes').hidden`),false,'Cancelar diálogo nativo não quita');await click('#app-exportacao-cancelar');
    const destination=path.join(folder,'escolhida','resultado.pdf');
    // Pasta inexistente: falha de escrita preserva o workspace e permite repetir.
    await command({save:{canceled:false,filePath:destination}});await click('#pdf-salvar-copia');await click('#app-exportacao-confirmar');await ready();
    assert.equal(await evaluate(`document.querySelector('#pdf-alteracoes-pendentes').hidden`),false);
    await command({save:{canceled:false,filePath:path.join(folder,'resultado.pdf')}});await click('#app-exportacao-confirmar');await ready();
    const pdf=await PDFDocument.load(await readFile(path.join(folder,'resultado.pdf')));assert.equal(pdf.getPageCount(),20);assert.equal(pdf.getPage(0).getRotation().angle,90);
    assert.equal(await evaluate(`document.querySelector('#pdf-alteracoes-pendentes').hidden`),true);
    const windowState=async()=>{try{return JSON.parse(await readFile(path.join(folder,'status.json'),'utf8'));}catch{return {};}};
    await command({action:'size'});await command({action:'maximize'});
    await wait(async()=>(await windowState()).maximized,'Janela maximizada');
    await command({action:'restore'});await wait(async()=>{const s=await windowState();return !s.maximized&&!s.minimized;},'Janela restaurada');
    await command({action:'minimize'});
    await wait(async()=>(await windowState()).minimized,'Janela minimizada');
    await command({action:'restore'});
    // Conferência pontual de impressão no Electron: uma página rasterizada e
    // outra ainda com texto/fontes, sem repetir toda a suíte longa de CMS.
    const signed=await criarPDFAssinado({total:2,paginas:[2]});
    await evaluate(`(()=>{const b=Uint8Array.from(atob('${Buffer.from(signed).toString('base64')}'),c=>c.charCodeAt(0)),d=new DataTransfer();d.items.add(new File([b],'assinatura.pdf'));const e=document.querySelector('#pdf-arquivo');e.files=d.files;e.dispatchEvent(new Event('change',{bubbles:true}));})()`);await ready();
    await click('#pdf-assinaturas-resolver');await click('#pdf-impressao-prever');
    await wait(()=>evaluate(`document.querySelector('#pdf-impressao-progresso').textContent.includes('Prévia pronta')`),'Preparo seletivo no Electron');
    await evaluate(`document.querySelector('#pdf-impressao-conferi').checked=true;document.querySelector('#pdf-impressao-conferi').dispatchEvent(new Event('change'))`);
    jsDialog=null;const applying=click('#pdf-impressao-aplicar');await wait(()=>jsDialog,'Consentimento do preparo');await cdp('Page.handleJavaScriptDialog',{accept:true});await applying;await ready();
    assert.equal(await evaluate(`document.querySelector('#pdf-assinaturas-titulo').textContent`),'Cópia para impressão preparada');
    await command({save:{canceled:false,filePath:path.join(folder,'impressao.pdf')}});await click('#pdf-salvar-copia');await click('#app-exportacao-confirmar');await ready();
    const printBytes=await readFile(path.join(folder,'impressao.pdf')),printPDF=await PDFDocument.load(printBytes);
    assert.equal(printPDF.catalog.has(PDFName.of('AcroForm')),false);assert.ok(printPDF.getPage(0).node.Resources().has(PDFName.of('Font')));
    const texts=await evaluate(`(async()=>{const{getDocument}=await import('/vendor/pdfjs/pdf.js'),t=getDocument({data:Uint8Array.from(atob('${printBytes.toString('base64')}'),c=>c.charCodeAt(0))});try{const p=await t.promise;return await Promise.all([1,2].map(async n=>(await(await p.getPage(n)).getTextContent()).items.map(i=>i.str||'').join('')));}finally{await t.destroy()}})()`);
    assert.match(texts[0],/Documento sintetico pagina 1/);assert.equal(texts[1],'');
    console.log('PASSOU: Electron preparou/exportou cópia seletiva sem AcroForm; página não rasterizada conserva texto/fontes.');
    await evaluate(`localStorage.setItem('4printshops.desktop.teste','persistiu')`);
    const shot=await cdp('Page.captureScreenshot');await writeFile(path.join(folder,'desktop.png'),Buffer.from(shot.data,'base64'));
    await command({action:'close'}).catch(()=>{});await wait(()=>child.exitCode!==null,'Encerramento limpo');
    await launch();assert.equal(await evaluate(`localStorage.getItem('4printshops.desktop.teste')`),'persistiu');
    assert.equal(await evaluate(`location.origin`),'app://4printshops');
    await evaluate(`(()=>{const a=document.createElement('a');a.href='https://example.com';document.body.append(a);a.click();a.remove()})()`);await pause(100);
    assert.equal(await evaluate(`location.href`),'app://4printshops/index.html');
    assert.equal(await evaluate(`window.open('https://example.com')===null`),true);
    assert.deepEqual(errors,[]);
    await command({action:'close'}).catch(()=>{});await wait(()=>child.exitCode!==null,'Segundo encerramento limpo');
    console.log('PASSOU: Electron seguro, PDF.js/miniaturas, giro, abas, salvamento/cancelamento/falha nativos, proteção de fechamento, persistência/origem estável e encerramento.');
    console.log('Artefatos: '+folder);
} catch(error){console.error(debug.slice(-6000));console.error('Artefatos: '+folder);throw error;}
finally {socket?.close();if(child&&child.exitCode===null)child.kill();}
