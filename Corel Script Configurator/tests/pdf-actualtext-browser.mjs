// Isolamento do mesmo helper da suíte geral; somente sintéticos e perfil próprio.
import assert from 'node:assert/strict';
import {spawn} from 'node:child_process';
import {mkdir,readFile,writeFile} from 'node:fs/promises';
import path from 'node:path';
import {PDFDocument} from '../public/vendor/pdf-lib/pdf-lib.js';
import {testarActualTextBrowser} from './helpers/pdf-actualtext-browser.mjs';
const folder=path.resolve(process.env.ACTUALTEXT_RESULTS||'.desktop-test/actualtext-audit/chrome-'+Date.now());
await mkdir(folder,{recursive:true});
const profile=path.join(folder,'profile-'+Date.now());await mkdir(profile);
let server,browser,socket,dialog,id=0;const pending=new Map(),errors=[],confirmations=[];
const sleep=ms=>new Promise(r=>setTimeout(r,ms));
const wait=async(fn,label)=>{const deadline=Date.now()+30000;while(Date.now()<deadline){if(await fn())return;await sleep(50);}throw Error('Espera: '+label);};
const cdp=(method,params={})=>new Promise((resolve,reject)=>{const n=++id,timer=setTimeout(()=>{pending.delete(n);reject(Error(method));},30000);pending.set(n,{resolve:r=>{clearTimeout(timer);resolve(r);},reject:e=>{clearTimeout(timer);reject(e);}});socket.send(JSON.stringify({id:n,method,params}));});
const avaliar=async expression=>{const r=await cdp('Runtime.evaluate',{expression,returnByValue:true,awaitPromise:true,userGesture:true});if(r.exceptionDetails)throw Error(r.exceptionDetails.exception?.description||r.exceptionDetails.text);return r.result.value;};
const esperarTela=(expression,label)=>wait(()=>avaliar(expression),label);
const clicar=async selector=>{await avaliar(`document.querySelector(${JSON.stringify(selector)}).click()`);if(selector==='#pdf-salvar-copia'){await esperarTela('document.querySelector("#app-exportacao-dialogo").open','janela exportar');await avaliar('document.querySelector("#app-exportacao-confirmar").click()');}};
const pronto=()=>esperarTela('!document.querySelector("#pdf-arquivo").disabled&&!!document.querySelector("#pdf-visualizacao canvas")','PDF pronto');
const resizePronto=()=>esperarTela('!document.querySelector("#pdf-redim-aplicar").disabled','prévia pronta');
const configurarResize=(modo='ajustar',alvo='atual',largura='148',altura='210')=>avaliar(`(()=>{for(const[id,value]of ${JSON.stringify([['largura',largura],['altura',altura],['modo',modo],['alvo',alvo]])}){const e=document.querySelector('#pdf-redim-'+id);e.value=value;e.dispatchEvent(new Event(id==='largura'||id==='altura'?'input':'change',{bubbles:true}));}})()`);
const confirmarImpressao=async(accept,markers=true)=>{
 dialog=null;const request=clicar('#pdf-redim-impressao');await wait(()=>!!dialog,'diálogo nativo de descarte');
 assert.equal(dialog.type,'confirm');assert.match(dialog.message,/INTEIRO/);assert.match(dialog.message,/acessibilidade/);if(markers)assert.match(dialog.message,/Marcadores/);
 confirmations.push({accept,message:dialog.message});await cdp('Page.handleJavaScriptDialog',{accept});await request;
};
const baixar=async(name,pages)=>{const file=path.join(folder,name);await wait(async()=>{try{return(await readFile(file)).length>0;}catch{return false;}},'PDF exportado '+name);const pdf=await PDFDocument.load(await readFile(file));assert.equal(pdf.getPageCount(),pages);return pdf;};
const imagem=async name=>{const r=await cdp('Page.captureScreenshot');await writeFile(path.join(folder,name+'.png'),Buffer.from(r.data,'base64'));};
const conferirPixelsExportados=async(name,number)=>{
 const b64=(await readFile(path.join(folder,name))).toString('base64');
 const ratio=await avaliar(`(async()=>{const{getDocument}=await import('/vendor/pdfjs/pdf.js'),task=getDocument({data:Uint8Array.from(atob(${JSON.stringify(b64)}),c=>c.charCodeAt(0))});try{const page=await(await task.promise).getPage(${number}),actual=document.querySelector('#pdf-visualizacao canvas'),expected=document.createElement('canvas');expected.width=actual.width;expected.height=actual.height;await page.render({canvas:expected,canvasContext:expected.getContext('2d'),viewport:page.getViewport({scale:actual.viewportPrincipalTeste.scale})}).promise;const a=actual.getContext('2d').getImageData(0,0,actual.width,actual.height).data,b=expected.getContext('2d').getImageData(0,0,expected.width,expected.height).data;let changed=0;for(let i=0;i<a.length;i++)if(Math.abs(a[i]-b[i])>2)changed++;expected.width=expected.height=0;return changed/a.length;}finally{await task.destroy();}})()`);
 assert.ok(ratio<.001,'prévia corresponde ao PDF exportado: '+ratio);
};
try{
 try{await fetch('http://localhost:3000');}catch{server=spawn(process.execPath,['app.js'],{windowsHide:true,stdio:'ignore'});}
 await wait(async()=>{try{return(await fetch('http://localhost:3000')).ok;}catch{return false;}},'servidor');
 browser=spawn(process.env.BROWSER_PATH||'C:/Program Files/Google/Chrome/Application/chrome.exe',['--headless=new','--disable-gpu','--no-first-run','--remote-debugging-port=0','--user-data-dir='+profile,'http://localhost:3000'],{windowsHide:true,stdio:'ignore'});
 let target;await wait(async()=>{try{const port=(await readFile(path.join(profile,'DevToolsActivePort'),'utf8')).split('\n')[0];target=(await(await fetch('http://127.0.0.1:'+port+'/json')).json()).find(t=>t.type==='page');return !!target;}catch{return false;}},'Chrome');
 socket=new WebSocket(target.webSocketDebuggerUrl);await new Promise(r=>socket.addEventListener('open',r,{once:true}));socket.addEventListener('message',e=>{const m=JSON.parse(e.data);if(m.id){m.error?pending.get(m.id)?.reject(Error(m.error.message)):pending.get(m.id)?.resolve(m.result);pending.delete(m.id);}if(m.method==='Page.javascriptDialogOpening')dialog=m.params;if(m.method==='Runtime.exceptionThrown')errors.push(m.params.exceptionDetails);});
 await cdp('Runtime.enable');await cdp('Page.enable');await cdp('Browser.setDownloadBehavior',{behavior:'allow',downloadPath:folder});await cdp('Emulation.setDeviceMetricsOverride',{width:1440,height:1000,deviceScaleFactor:1,mobile:false});
 await esperarTela('!!document.querySelector("[data-tool=pdfs]")','interface');await clicar('[data-tool=pdfs]');await esperarTela('!!document.querySelector("#pdf-arquivo")','editor');
 // Mesmo viewport real capturado pela suíte geral; não deduzir escala por
 // dimensões arredondadas do canvas nem afrouxar a tolerância de pixels.
 await avaliar(`(async()=>{const{PDFDocument}=await import('/vendor/pdf-lib/pdf-lib.js'),{getDocument}=await import('/vendor/pdfjs/pdf.js'),p=await PDFDocument.create();p.addPage();(await import('/vendor/pdfjs/pdf.js')).GlobalWorkerOptions.workerSrc='/vendor/pdfjs/pdf.worker.js';const t=getDocument({data:await p.save()});try{const page=await(await t.promise).getPage(1),proto=Object.getPrototypeOf(page),render=proto.render;proto.render=function(options){if(options.canvas.className==='pdf-pagina-canvas')options.canvas.viewportPrincipalTeste=options.viewport;return render.call(this,options);};}finally{await t.destroy();}})()`);
 await avaliar(`window.selecionarPDFTeste=(id,files)=>{const t=new DataTransfer();files.forEach(f=>t.items.add(f));const e=document.getElementById(id);e.files=t.files;e.dispatchEvent(new Event('change',{bubbles:true}));};window.compararImpressao=async(a,b)=>{const{getDocument}=await import('/vendor/pdfjs/pdf.js');const render=async(data)=>{const t=getDocument({data:Uint8Array.from(atob(data),c=>c.charCodeAt(0))});try{const p=await(await t.promise).getPage(1),v=p.getViewport({scale:1}),c=document.createElement('canvas');c.width=Math.ceil(v.width);c.height=Math.ceil(v.height);await p.render({canvas:c,canvasContext:c.getContext('2d'),viewport:v}).promise;const r={width:c.width,height:c.height,pixels:c.getContext('2d').getImageData(0,0,c.width,c.height).data,texto:(await p.getTextContent()).items.map(i=>i.str).join(' ')};c.width=c.height=0;return r;}finally{await t.destroy();}};const x=await render(a),y=await render(b);let diferentes=0;for(let i=0;i<x.pixels.length;i++)if(x.pixels[i]!==y.pixels[i])diferentes++;return{dimensoes:x.width===y.width&&x.height===y.height,diferentes,total:x.pixels.length,textoAntes:x.texto,textoDepois:y.texto};};`);
 await testarActualTextBrowser({avaliar,clicar,pronto,esperarTela,configurarResize,resizePronto,confirmarImpressao,baixar,conferirPixelsExportados,imagem,temporario:folder});
 assert.deepEqual(errors,[]);assert.ok(confirmations.some(c=>!c.accept));assert.ok(confirmations.some(c=>c.accept));
 await writeFile(path.join(folder,'results.json'),JSON.stringify({passed:true,errors,confirmations,noPrinting:true},null,2));console.log('PASSOU: ActualText isolado, preservação, bloqueios e consentimento nativo. Sem impressão. Evidências: '+folder);
}catch(error){console.error(error);try{console.error(await avaliar('({status:document.querySelector("#pdf-status")?.textContent,resize:document.querySelector("#pdf-redim-resultado")?.textContent})'));await imagem('failure');}catch{}process.exitCode=1;}
finally{socket?.close();browser?.kill();server?.kill();}
