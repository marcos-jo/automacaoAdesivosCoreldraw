import assert from 'node:assert/strict';
import {spawn} from 'node:child_process';
import {mkdtemp,readFile,writeFile,unlink} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {setTimeout as pause} from 'node:timers/promises';
import {PDFDocument,PDFName} from '../public/vendor/pdf-lib/pdf-lib.js';
import {criarPDFPreparadoSintetico} from './helpers/pdf-print-prepared-fixture.mjs';
import {testarAssinaturasBrowser} from './helpers/pdf-signature-browser.mjs';
import {ConsoleAssinaturas} from './helpers/pdf-signature-console.mjs';

const root=fileURLToPath(new URL('../',import.meta.url)), folder=await mkdtemp(path.join(tmpdir(),'4printshops-desktop-'));
let child,socket,id=0,serial=0,options={},debug=''; const requests=new Map(),errors=[],dialogs=[];
const consoleAssinaturas=new ConsoleAssinaturas();
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
    socket.addEventListener('message',e=>{const m=JSON.parse(e.data);if(m.id){requests.get(m.id)?.resolve(m.result);requests.delete(m.id);}consoleAssinaturas.registrar(m.method,m.params);if(m.method==='Runtime.exceptionThrown')errors.push(m.params.exceptionDetails.text);if(m.method==='Page.javascriptDialogOpening')dialogs.push(m.params);});
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
    const dialogsAntesDoFechamento=dialogs.length;
    await command({action:'close',answer:0});await pause(200);assert.equal(child.exitCode,null,'Cancelar mantém o app aberto');
    assert.equal(dialogs.length,dialogsAntesDoFechamento+1,'Exatamente um beforeunload na tentativa deliberada de fechar o PDF alterado');
    const fechamentoCancelado=dialogs.at(-1);
    assert.equal(fechamentoCancelado.type,'beforeunload');
    assert.equal(fechamentoCancelado.url,'app://4printshops/index.html');
    assert.equal(fechamentoCancelado.message,'');
    assert.equal(fechamentoCancelado.hasBrowserHandler,true);
    await click('[data-tool="inicio"]');await click('[data-tool="qrcodes"]');
    await command({action:'probe-work'});assert.deepEqual(JSON.parse(await readFile(path.join(folder,'status.json'),'utf8')).workState,{pending:true,busy:false},'Consulta de todas as abas encontra PDF oculto alterado');
    await click('[data-tool="pdfs"]');
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
    // PDF-008: assinaturas aplicadas não podem ser reconstruídas para criar a
    // fixture. Reutiliza recusas e snapshots estritos já exercitados no Edge.
    await testarAssinaturasBrowser({avaliar:evaluate,pronto:ready,
        esperarTela:(expressao,rotulo)=>wait(()=>evaluate(expressao),rotulo),
        imagem:async nome=>{const shot=await cdp('Page.captureScreenshot');await writeFile(path.join(folder,nome+'.png'),Buffer.from(shot.data,'base64'));},
        recusaConsole:(cenario,acao)=>consoleAssinaturas.recusaGiro(cenario,acao,{cdp,origem:'app://4printshops'})});
    assert.equal(await evaluate(`document.querySelector('#pdf-impressao-dialogo').open`),false);
    // Entrada legítima já preparada, sem remover assinatura ou forjar estado
    // interno. A página vetorial mantém texto; a rasterizada não o inventa.
    const prepared=await criarPDFPreparadoSintetico();
    await evaluate(`(()=>{const b=Uint8Array.from(atob('${Buffer.from(prepared).toString('base64')}'),c=>c.charCodeAt(0)),d=new DataTransfer();d.items.add(new File([b],'copia-preparada.pdf'));const e=document.querySelector('#pdf-arquivo');e.files=d.files;e.dispatchEvent(new Event('change',{bubbles:true}));})()`);
    await wait(()=>evaluate(`!document.querySelector('#pdf-arquivo').disabled||document.querySelector('#pdf-confirmar-descarte').open`),'Importação da cópia preparada');
    if(await evaluate(`document.querySelector('#pdf-confirmar-descarte').open`))await click('#pdf-descartar-sair');
    await ready();
    await wait(()=>evaluate(`document.querySelector('#pdf-nome').textContent.includes('copia-preparada.pdf')`),'Cópia preparada importada');
    assert.equal(await evaluate(`document.querySelector('#pdf-assinaturas-alerta').hidden`),true);
    await command({save:{canceled:false,filePath:path.join(folder,'impressao.pdf')}});await click('#pdf-salvar-copia');await click('#app-exportacao-confirmar');await ready();
    const printBytes=await readFile(path.join(folder,'impressao.pdf')),printPDF=await PDFDocument.load(printBytes);
    assert.equal(printPDF.catalog.has(PDFName.of('AcroForm')),false);assert.ok(printPDF.getPage(0).node.Resources().has(PDFName.of('Font')));
    const texts=await evaluate(`(async()=>{const{getDocument}=await import('/vendor/pdfjs/pdf.js'),t=getDocument({data:Uint8Array.from(atob('${printBytes.toString('base64')}'),c=>c.charCodeAt(0))});try{const p=await t.promise;return await Promise.all([1,2].map(async n=>(await(await p.getPage(n)).getTextContent()).items.map(i=>i.str||'').join('')));}finally{await t.destroy()}})()`);
    assert.match(texts[0],/Documento sintetico pagina 1/);assert.equal(texts[1],'');
    console.log('PASSOU: Electron abriu/exportou fixture já preparada sem AcroForm; página vetorial conserva texto/fontes e página rasterizada não tem texto.');
    assert.equal(await evaluate(`document.querySelectorAll('dialog[open]').length`),0,'Nenhum diálogo pendente após as recusas e a exportação legítima');
    await evaluate(`localStorage.setItem('4printshops.desktop.teste','persistiu')`);
    const shot=await cdp('Page.captureScreenshot');await writeFile(path.join(folder,'desktop.png'),Buffer.from(shot.data,'base64'));
    await command({action:'close'}).catch(()=>{});await wait(()=>child.exitCode!==null,'Encerramento limpo');
    await launch();assert.equal(await evaluate(`localStorage.getItem('4printshops.desktop.teste')`),'persistiu');
    assert.equal(await evaluate(`location.origin`),'app://4printshops');
    await evaluate(`(()=>{const a=document.createElement('a');a.href='https://example.com';document.body.append(a);a.click();a.remove()})()`);await pause(100);
    assert.equal(await evaluate(`location.href`),'app://4printshops/index.html');
    assert.equal(await evaluate(`window.open('https://example.com')===null`),true);
    assert.deepEqual(errors,[]);
    assert.deepEqual(dialogs,[fechamentoCancelado],'Somente o beforeunload validado na tentativa deliberada de fechar; diálogos adicionais reprovam');
    assert.equal(await evaluate(`document.querySelectorAll('dialog[open]').length`),0,'Nenhum diálogo HTML pendente');
    const consoleResumo=consoleAssinaturas.verificarFinal(true);
    await writeFile(path.join(folder,'console-assinaturas.json'),JSON.stringify(consoleResumo,null,2));
    await command({action:'close'}).catch(()=>{});await wait(()=>child.exitCode!==null,'Segundo encerramento limpo');
    console.log('PASSOU: Electron seguro, PDF.js/miniaturas, giro, abas, salvamento/cancelamento/falha nativos, proteção de fechamento, persistência/origem estável e encerramento.');
    console.log('Artefatos: '+folder);
} catch(error){console.error(debug.slice(-6000));console.error(await evaluate(`JSON.stringify({name:document.querySelector('#pdf-nome')?.textContent,progress:document.querySelector('#pdf-impressao-progresso')?.textContent,status:document.querySelector('#pdf-impressao-status')?.textContent,dialogs:[...document.querySelectorAll('dialog')].filter(d=>d.open).map(d=>d.id)})`).catch(()=>''));console.error('Artefatos: '+folder);throw error;}
finally {socket?.close();if(child&&child.exitCode===null)child.kill();}
