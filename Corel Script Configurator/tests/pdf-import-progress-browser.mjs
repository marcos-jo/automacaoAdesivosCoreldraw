// Chrome e Electron de desenvolvimento, somente sintéticos e perfis isolados.
import assert from 'node:assert/strict';import {spawn} from 'node:child_process';
import {mkdir,readFile,writeFile} from 'node:fs/promises';import path from 'node:path';
import {png,jpg} from './helpers/pdf-combine-fixtures.mjs';
import {pdfProtegido} from './helpers/pdf-import-protection-fixtures.mjs';
import {criarPDFAssinado} from './helpers/pdf-signature-fixture.mjs';
const electron=process.env.IMPORT_PROGRESS_ELECTRON==='1',folder=path.resolve(`.desktop-test/import-progress/${electron?'electron':'chrome'}-${Date.now()}`);await mkdir(folder,{recursive:true});
let app,server,socket,id=0;const pending=new Map(),errors=[],pause=ms=>new Promise(r=>setTimeout(r,ms));
const wait=async(fn,label)=>{const end=Date.now()+60000;while(Date.now()<end){if(await fn())return;await pause(50);}throw Error('Espera: '+label);};
const cdp=(method,params={})=>new Promise((resolve,reject)=>{const n=++id,t=setTimeout(()=>{pending.delete(n);reject(Error(method));},60000);pending.set(n,{resolve:v=>{clearTimeout(t);resolve(v);},reject});socket.send(JSON.stringify({id:n,method,params}));});
const ev=async expression=>{const r=await cdp('Runtime.evaluate',{expression,awaitPromise:true,returnByValue:true,userGesture:true});if(r.exceptionDetails)throw Error(r.exceptionDetails.exception?.description||r.exceptionDetails.text);return r.result.value;};
const click=s=>ev(`document.querySelector(${JSON.stringify(s)}).click()`);
const run=code=>ev(`(async()=>{${code}})()`);
const ready=()=>wait(()=>ev('!document.querySelector("#pdf-arquivo").disabled&&!document.querySelector("#pdf-uniao-arquivos").disabled'),'controles liberados');
const screenshot=async name=>{const r=await cdp('Page.captureScreenshot');await writeFile(path.join(folder,name+'.png'),Buffer.from(r.data,'base64'));};
try{
 let profile,exe,args;const env={...process.env};delete env.ELECTRON_RUN_AS_NODE;
 if(electron){env.PRINTSHOPS_TEST_FOLDER=folder;await writeFile(path.join(folder,'command.json'),'{"serial":0}');profile=path.join(folder,'profile');exe=path.resolve('node_modules/electron/dist/electron.exe');args=['tests/helpers/desktop-entry.mjs','--remote-debugging-port=0'];}
 else{try{await fetch('http://localhost:3000');}catch{server=spawn(process.execPath,['app.js'],{windowsHide:true,stdio:'ignore'});}await wait(async()=>{try{return(await fetch('http://localhost:3000')).ok;}catch{return false;}},'servidor');profile=path.join(folder,'profile');await mkdir(profile);exe='C:/Program Files/Google/Chrome/Application/chrome.exe';args=['--headless=new','--disable-gpu','--no-first-run','--remote-debugging-port=0','--user-data-dir='+profile,'http://localhost:3000'];}
 app=spawn(exe,args,{env,windowsHide:true,stdio:'ignore'});let target;
 await wait(async()=>{try{const port=(await readFile(path.join(profile,'DevToolsActivePort'),'utf8')).split('\n')[0];target=(await(await fetch('http://127.0.0.1:'+port+'/json')).json()).find(t=>t.type==='page');return !!target;}catch{return false;}},'página');
 socket=new WebSocket(target.webSocketDebuggerUrl);await new Promise(r=>socket.addEventListener('open',r,{once:true}));socket.addEventListener('message',e=>{const m=JSON.parse(e.data);if(m.id){m.error?pending.get(m.id)?.reject(Error(m.error.message)):pending.get(m.id)?.resolve(m.result);pending.delete(m.id);}if(m.method==='Runtime.exceptionThrown')errors.push(m.params.exceptionDetails.text);});
 await cdp('Runtime.enable');await cdp('Page.enable');await wait(()=>ev('!!document.querySelector("[data-tool=pdfs]")'),'início');await click('[data-tool=pdfs]');await ready();
 await run(`window.makeImportPDF=async(n=1)=>{const{PDFDocument}=await import('/vendor/pdf-lib/pdf-lib.js'),p=await PDFDocument.create();for(let i=0;i<n;i++)p.addPage([300+i/100,400]).drawText('SYNTHETIC '+i,{x:20,y:40});return new File([await p.save()],'sintetico.pdf',{type:'application/pdf'});};window.selectImport=(id,files)=>document.querySelector('#'+id).dispatchEvent(new CustomEvent('change',{bubbles:true,detail:{files}}));window.originalImport=await makeImportPDF();selectImport('pdf-arquivo',[originalImport]);`);
 await ready();await click('#pdf-girar-direita');await ready();const before=await ev('({total:document.querySelector("#pdf-total").textContent,size:document.querySelector("#pdf-tamanho").textContent,pending:document.querySelector("#pdf-alteracoes-pendentes").hidden})');assert.equal(before.pending,false);
 await click('#pdf-uniao-abrir');
 await run(`window.tickImport=0;window.tickImportTimer=setInterval(()=>tickImport++,20);window.slowImportGate=new Promise(r=>window.releaseImport=r);const a=await makeImportPDF(),p=Uint8Array.from(atob('${Buffer.from(png()).toString('base64')}'),c=>c.charCodeAt(0));window.slowImport={name:'lenta.png',size:p.length,arrayBuffer:async()=>{await slowImportGate;return p.buffer;}};window.photoImport=new File([Uint8Array.from(atob('${Buffer.from(jpg()).toString('base64')}'),c=>c.charCodeAt(0))],'foto.jpg');selectImport('pdf-uniao-arquivos',[a,slowImport,photoImport]);`);
 await wait(()=>ev('document.querySelectorAll("#pdf-uniao-preparacao li").length===3&&!!document.querySelector("#pdf-uniao-preparacao img")'),'primeira miniatura antes do lote');
 assert.match(await ev('document.querySelector("#pdf-uniao-progresso").textContent'),/1 de 3/);assert.equal(await ev('document.querySelector("#pdf-uniao-gerar").disabled'),true);await pause(160);assert.ok(await ev('tickImport>5'));await screenshot('progressivo');
 await click('#pdf-uniao-cancelar-importacao');await ev('releaseImport();clearInterval(tickImportTimer)');await pause(200);assert.equal(await ev('document.querySelector("#pdf-uniao-janela").open'),false);assert.equal(await ev('document.querySelectorAll("#pdf-uniao-preparacao img").length'),0);
 assert.deepEqual(await ev('({total:document.querySelector("#pdf-total").textContent,size:document.querySelector("#pdf-tamanho").textContent,pending:document.querySelector("#pdf-alteracoes-pendentes").hidden})'),before);
 for(const pages of [1,100,350]){await click('#pdf-uniao-abrir');await run(`selectImport('pdf-uniao-arquivos',[await makeImportPDF(${pages})]);`);await ready();assert.equal(await ev('document.querySelectorAll("#pdf-uniao-lista li").length'),pages);assert.equal(await ev('document.querySelector("#pdf-uniao-gerar").disabled'),false);await click('#pdf-uniao-limpar');await ready();}
 const withPassword=pdfProtegido('sintetica').toString('base64'),restricted=pdfProtegido().toString('base64');
 await run(`const fromBase64=(bytes,name)=>new File([Uint8Array.from(atob(bytes),c=>c.charCodeAt(0))],name,{type:'application/pdf'});selectImport('pdf-uniao-arquivos',[await makeImportPDF(),fromBase64('${withPassword}','senha.pdf'),new File(['invalid'],'falha.pdf'),fromBase64('${restricted}','restrito.pdf'),photoImport]);`);
 await wait(()=>ev('document.querySelectorAll("#pdf-uniao-preparacao button:not([hidden])").length===3'),'diagnósticos individuais');
 const states=await ev('Array.from(document.querySelectorAll(".pdf-import-state"),e=>e.textContent)');
 assert.equal(states[0],'Pronto');assert.match(states[1],/Senha de abertura/);assert.match(states[2],/PDF inválido/);assert.match(states[3],/restrições de operações/);assert.equal(states[4],'Pronto');
 assert.equal(await ev('document.querySelector("#pdf-uniao-gerar").disabled'),true);
 await screenshot('arquivos-com-falha');
 await click('#pdf-uniao-preparacao li:nth-child(2) button');await click('#pdf-uniao-preparacao li:nth-child(3) button');
 assert.equal(await ev('document.querySelector("#pdf-uniao-gerar").disabled'),true);
 await click('#pdf-uniao-preparacao li:nth-child(4) button');await ready();
 assert.equal(await ev('document.querySelectorAll("#pdf-uniao-lista li").length'),2);
 assert.deepEqual(await ev('({total:document.querySelector("#pdf-total").textContent,size:document.querySelector("#pdf-tamanho").textContent,pending:document.querySelector("#pdf-alteracoes-pendentes").hidden})'),before);
 await click('#pdf-uniao-limpar');await ready();await click('#pdf-uniao-fechar');
 await click('#pdf-uniao-abrir');
 await run(`
  window.WorkerImportOriginal=Worker;
  window.workerImportUrl=URL.createObjectURL(new Blob([
   'self.onmessage=()=>{self.postMessage("iniciou");const fim=Date.now()+10000;while(Date.now()<fim){};}'
  ],{type:'text/javascript'}));
  window.Worker=class extends WorkerImportOriginal {
   constructor(url,options){
    const importar=String(url).endsWith('/pdf-import-worker.js');
    super(importar?workerImportUrl:url,options);
    if(importar){this.importacao=true;this.addEventListener('message',event=>{event.stopImmediatePropagation();window.workerImportOcupado=true;});}
   }
   terminate(){if(this.importacao)window.workerImportEncerrado=true;super.terminate();}
  };
  selectImport('pdf-uniao-arquivos',[await makeImportPDF()]);
 `);
 await wait(()=>ev('window.workerImportOcupado===true'),'worker ocupado com tarefa sintética');
 assert.equal(await ev('window.workerImportEncerrado===true'),false);
 await click('#pdf-uniao-cancelar-importacao');
 await wait(()=>ev('window.workerImportEncerrado===true&&!document.querySelector("#pdf-uniao-janela").open'),'cancelamento encerra worker real');
 await ev('window.Worker=WorkerImportOriginal;URL.revokeObjectURL(workerImportUrl)');
 const signed=Buffer.from(await criarPDFAssinado({total:3})).toString('base64');
 const signature=await run(`
  const {PDFCombinacao}=await import('/js/tools/pdf-combine-session.js');
  const {executarImportacao}=await import('/js/tools/pdf-import-executor.js');
  const {PDFWorkspace}=await import('/js/tools/pdf-workspace.js');
  const draft=new PDFCombinacao(executarImportacao);
  await draft.add([new File([Uint8Array.from(atob('${signed}'),c=>c.charCodeAt(0))],'assinada.pdf')]);
  draft.move(0,2);
  const workspace=new PDFWorkspace();
  await workspace.adicionarArquivos([new File([await draft.build()],'lote.pdf')],'lote.pdf',{fontes:draft.origins()});
  const resumo=workspace.obterResumoAssinaturas();
  return {total:workspace.getTotalPaginas(),assinaturas:resumo.presentes,visual:resumo.paginasVisuais.length};
 `);
 assert.equal(signature.total,3);assert.equal(signature.assinaturas,1);assert.ok(signature.visual>0);
 await click('#pdf-uniao-abrir');await run(`selectImport('pdf-uniao-arquivos',[await makeImportPDF()]);`);await ready();await ev(`document.querySelector('#pdf-uniao-destino').value='novo'`);await click('#pdf-uniao-gerar');await wait(()=>ev('document.querySelector("#pdf-confirmar-descarte").open'),'edição pendente protegida');await click('#pdf-continuar-editando');assert.equal(await ev('document.querySelector("#pdf-tamanho").textContent'),before.size);
 await click('#pdf-uniao-gerar');await wait(()=>ev('document.querySelector("#pdf-confirmar-descarte").open'),'novo consentimento');await click('#pdf-descartar-sair');await ready();assert.equal(await ev('document.querySelector("#pdf-total").textContent'),'1');
 if(electron){
  await run(`const{PreparacaoConferencia}=await import('/js/tools/pdf-import-review.js');for(const name of ['begin','accept','complete']){const fn=PreparacaoConferencia.prototype[name];PreparacaoConferencia.prototype[name]=function(...args){try{const r=fn.apply(this,args);return r?.then?r.catch(e=>{window.importTestError=name+': '+e.message;throw e;}):r;}catch(e){window.importTestError=name+': '+e.message;throw e;}};}`);
  const data=await ev('(async()=>Array.from(new Uint8Array(await originalImport.arrayBuffer())))()');
  const files=['primeiro.pdf','protegido.pdf','ultimo.jpg'].map(name=>path.join(folder,name));
  await Promise.all([writeFile(files[0],new Uint8Array(data)),writeFile(files[1],pdfProtegido()),writeFile(files[2],jpg())]);
  await writeFile(path.join(folder,'command.json'),JSON.stringify({serial:1,open:{canceled:false,filePaths:files}}));await wait(async()=>{try{return JSON.parse(await readFile(path.join(folder,'status.json'),'utf8')).serial===1;}catch{return false;}},'diálogo sintético');
  await click('#pdf-arquivo');await wait(()=>ev(`(()=>{if(document.querySelector('#pdf-status').textContent.startsWith('Não foi possível importar:'))throw Error(window.importTestError||document.querySelector('#pdf-status').textContent);return !!document.querySelector('#pdf-uniao-preparacao li:nth-child(2) button:not([hidden])');})()`),'ponte progressiva Electron');
  assert.match(await ev('document.querySelector("#pdf-uniao-preparacao li:nth-child(2) .pdf-import-state").textContent'),/restrições de operações/);
  assert.equal(await ev('document.querySelector("#pdf-uniao-gerar").disabled'),true);
  await click('#pdf-uniao-preparacao li:nth-child(2) button');await ready();
  assert.deepEqual(await ev('Array.from(document.querySelectorAll("#pdf-uniao-lista .pdf-uniao-nome"),e=>e.textContent.split(" · ")[0])'),['1. primeiro.pdf','2. ultimo.jpg']);
  assert.equal(await ev('document.querySelector("#pdf-total").textContent'),'1');
  await click('#pdf-uniao-gerar');
  await wait(()=>ev('!document.querySelector("#pdf-uniao-janela").open&&document.querySelector("#pdf-total").textContent==="3"'),'lote nativo aplicado ao editor');
 }
 assert.deepEqual(errors,[]);await writeFile(path.join(folder,'results.json'),JSON.stringify({passed:true,electron,progressive:true,cancelAndReopen:true,protectedAndInvalidInMixedBatch:true,workerCancellation:true,pages:[1,100,350],pendingEditsProtected:true,noPrinting:true},null,2));console.log('PASSOU conferência progressiva '+(electron?'Electron':'Chrome')+'; evidências: '+folder);
}catch(error){console.error(error);try{console.error(await ev('({status:document.querySelector("#pdf-uniao-status")?.textContent,count:document.querySelector("#pdf-uniao-progresso")?.textContent})'));await screenshot('failure');}catch{}process.exitCode=1;}
finally{socket?.close();app?.kill();server?.kill();}
