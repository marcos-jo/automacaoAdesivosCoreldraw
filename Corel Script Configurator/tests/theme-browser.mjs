import {spawn} from 'node:child_process';
import {readFile,writeFile,mkdir} from 'node:fs/promises';
import path from 'node:path';
import {setTimeout as pause} from 'node:timers/promises';
import assert from 'node:assert/strict';
const before=process.env.THEME_BEFORE==='1';const root=path.resolve('.desktop-test/theme15/'+(before?'before':'after'));await mkdir(root,{recursive:true});
const server=spawn(process.execPath,['app.js'],{windowsHide:true,stdio:'ignore'});await pause(800);
const profile=path.join(root,'profile-'+Date.now());await mkdir(profile);
const app=spawn('C:/Program Files/Google/Chrome/Application/chrome.exe',['--headless=new','--disable-gpu','--no-first-run','--remote-debugging-port=0','--user-data-dir='+profile,'http://localhost:3000'],{windowsHide:true,stdio:'ignore'});
let ws,id=0;const pending=new Map();
const wait=async fn=>{for(let i=0;i<400;i++){try{if(await fn())return;}catch{}await pause(50);}throw Error('Espera esgotada');};
const cdp=(method,params={})=>new Promise((resolve,reject)=>{const n=++id;pending.set(n,{resolve,reject});ws.send(JSON.stringify({id:n,method,params}));});
const ev=async expression=>{const r=await cdp('Runtime.evaluate',{expression,awaitPromise:true,returnByValue:true});if(r.exceptionDetails)throw Error(r.exceptionDetails.text);return r.result.value;};
const shot=async name=>{await pause(300);const r=await cdp('Page.captureScreenshot');await writeFile(path.join(root,name+'.png'),Buffer.from(r.data,'base64'));};
try{
 let port;await wait(async()=>{port=(await readFile(path.join(profile,'DevToolsActivePort'),'utf8')).split('\n')[0];return port;});let target;await wait(async()=>{target=(await(await fetch('http://127.0.0.1:'+port+'/json')).json()).find(t=>t.type==='page');return target;});ws=new WebSocket(target.webSocketDebuggerUrl);await new Promise(r=>ws.addEventListener('open',r,{once:true}));ws.addEventListener('message',e=>{const m=JSON.parse(e.data);if(m.id){m.error?pending.get(m.id)?.reject(m.error):pending.get(m.id)?.resolve(m.result);pending.delete(m.id);}});
 await cdp('Emulation.setDeviceMetricsOverride',{width:1366,height:768,deviceScaleFactor:1,mobile:false});await wait(()=>ev("!!document.querySelector('[data-tool=pdfs]')"));await ev("document.querySelector('[data-tool=pdfs]').click()");await wait(()=>ev("!!document.querySelector('#pdf-arquivo')"));await shot('vazio-claro');
 if(!before)assert.equal(await ev("document.documentElement.dataset.theme"),'light');
 await ev(`(async()=>{const {PDFDocument}=await import('/vendor/pdf-lib/pdf-lib.js');const pdf=await PDFDocument.create();pdf.addPage([595,842]).drawText('Teste de tema / pagina branca',{x:30,y:700});const bytes=await pdf.save();window.themeFixture=bytes;const d=new DataTransfer();d.items.add(new File([bytes],'tema.pdf',{type:'application/pdf'}));const e=document.querySelector('#pdf-arquivo');e.files=d.files;e.dispatchEvent(new Event('change',{bubbles:true}));})()`);await wait(()=>ev("!!document.querySelector('#pdf-visualizacao canvas')&&!document.querySelector('#pdf-arquivo').disabled"));await shot('pdf-claro');
 const height=await ev("document.querySelector('#pdf-visualizacao').getBoundingClientRect().height");
 if(!before){
  assert.equal(await ev("document.querySelector('#pdf-arquivo').closest('label').hidden"),true);
  await ev("document.querySelector('#app-menu-exibir').click();document.querySelector('#app-theme-menu summary').click();document.querySelector('[data-theme-choice=dark]').click()");assert.equal(await ev("document.documentElement.dataset.theme"),'dark');await ev("document.querySelectorAll('#app-menus details').forEach(d=>d.open=false)");await shot('pdf-escuro');
  assert.deepEqual(await ev("Array.from(document.querySelector('#pdf-visualizacao canvas').getContext('2d').getImageData(2,2,1,1).data)"),[255,255,255,255]);
  await ev("document.querySelector('#pdf-informacoes').click()");assert.equal(await ev("document.querySelector('#pdf-detalhes').hidden"),false);await ev("document.querySelector('#pdf-informacoes').click()");
  await ev("document.querySelector('#app-menu-arquivo').click();document.querySelector('[data-command=pdf-importar]').click()");
  await cdp('Page.reload');await wait(()=>ev("!!document.querySelector('[data-tool=pdfs]')"));assert.equal(await ev("document.documentElement.dataset.theme"),'dark');await shot('inicio-escuro-persistido');
  await ev("document.querySelector('[data-tool=pdfs]').click()");assert.equal(await ev("document.querySelector('#pdf-arquivo').closest('label').hidden"),false);
 }
 await writeFile(path.join(root,'result.json'),JSON.stringify({passed:true,height,before},null,2));console.log('PASSOU tema/layout', {before,height});
}finally{ws?.close();app.kill();server.kill();}
