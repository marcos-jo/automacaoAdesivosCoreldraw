// Chrome ou Electron de desenvolvimento, perfis isolados; não instala nem imprime.
import assert from 'node:assert/strict';
import {spawn} from 'node:child_process';
import {mkdir,readFile,writeFile} from 'node:fs/promises';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {PDFDocument} from '../public/vendor/pdf-lib/pdf-lib.js';
import {conferirTransformacoes} from './helpers/pdf-imposition-proof.mjs';
import {testarPoliticasEscalaNup} from './helpers/pdf-imposition-scale-policy-browser.mjs';
const root=fileURLToPath(new URL('../',import.meta.url)),electron=process.env.NUP_SELECTION_ELECTRON==='1';
const folder=path.resolve(process.env.NUP_SELECTION_RESULTS||`.desktop-test/nup25/${electron?'electron':'chrome'}`);await mkdir(folder,{recursive:true});
let server,app,socket,id=0;const pending=new Map(),errors=[],reports=[];
const sleep=ms=>new Promise(r=>setTimeout(r,ms));
const wait=async(fn,label)=>{for(let i=0;i<400;i++){if(await fn())return;await sleep(50);}throw Error(label);};
const cdp=(method,params={})=>new Promise((resolve,reject)=>{const n=++id,t=setTimeout(()=>{pending.delete(n);reject(Error(method));},30000);pending.set(n,{resolve:r=>{clearTimeout(t);resolve(r);},reject});socket.send(JSON.stringify({id:n,method,params}));});
const ev=async expression=>{const r=await cdp('Runtime.evaluate',{expression,awaitPromise:true,returnByValue:true,userGesture:true});if(r.exceptionDetails)throw Error(r.exceptionDetails.exception?.description||r.exceptionDetails.text);return r.result.value;};
const ready=()=>wait(()=>ev('(()=>{if(u.selectionError)throw Error(u.get("status").textContent);if(!u.get("tentar").hidden)throw Error(u.get("status").textContent);return !!u.result&&!u.busy&&!u.dirty;})()'),'prévia concluída');
const set=async(id,value)=>{await ev(`(()=>{const e=u.get(${JSON.stringify(id)});if(e.type==='checkbox')e.checked=${JSON.stringify(value)};else e.value=${JSON.stringify(String(value))};e.dispatchEvent(new Event('input',{bubbles:true}));})()`);await ready();};
const screenshot=async name=>{const r=await cdp('Page.captureScreenshot');await writeFile(path.join(folder,name+'.png'),Buffer.from(r.data,'base64'));};
const inspect=async(count,capacity,name)=>{
 const r=await ev(`(async()=>{const b=u.result.bytes,p=u.result.plan,j=await import('/vendor/pdfjs/pdf.js'),task=j.getDocument({data:b.slice()}),texts=[];try{const d=await task.promise;for(let i=1;i<=d.numPages;i++){const t=await(await d.getPage(i)).getTextContent();texts.push([...t.items.map(t=>t.str).join(' ').matchAll(/PAGE ([0-9]+)/g)].map(m=>Number(m[1])));}}finally{await task.destroy();}let s='';for(let i=0;i<b.length;i+=32768)s+=String.fromCharCode(...b.subarray(i,i+32768));return {plan:p,pdf:btoa(s),texts,unchanged:u.originalBytes.every((v,i)=>v===sourceBefore[i])};})()`);
 const bytes=Buffer.from(r.pdf,'base64');delete r.pdf;const expected=Array.from({length:count},(_,i)=>i+1);
 assert.equal(r.plan.totalFolhas,Math.ceil(count/capacity));assert.deepEqual(r.plan.folhas.flat().map(p=>p.origem),expected);assert.deepEqual(r.texts.flat(),expected);assert.equal(r.plan.folhas.at(-1).length,count%capacity||capacity);assert.equal(r.unchanged,true);
 assert.equal((await PDFDocument.load(bytes)).getPageCount(),Math.ceil(count/capacity));assert.equal(await conferirTransformacoes(bytes,r.plan),count);await writeFile(path.join(folder,name+'.pdf'),bytes);reports.push({name,...r});
};
try{
 const profile=path.join(folder,'profile-'+Date.now());await mkdir(profile);
 if(!electron){try{await fetch('http://localhost:3000');}catch{server=spawn(process.execPath,['app.js'],{cwd:root,windowsHide:true,stdio:'ignore'});}await wait(async()=>{try{return(await fetch('http://localhost:3000')).ok;}catch{return false;}},'servidor');}
 const env={...process.env};delete env.ELECTRON_RUN_AS_NODE;
 let exe,args;
 if(electron){env.PRINTSHOPS_TEST_FOLDER=path.dirname(profile);await writeFile(path.join(env.PRINTSHOPS_TEST_FOLDER,'command.json'),JSON.stringify({serial:0}));exe=path.join(root,'node_modules/electron/dist/electron.exe');args=['tests/helpers/desktop-entry.mjs','--remote-debugging-port=0','--disable-renderer-backgrounding'];}
 else{exe=process.env.BROWSER_PATH||'C:/Program Files/Google/Chrome/Application/chrome.exe';args=['--headless=new','--disable-gpu','--remote-debugging-port=0','--no-first-run','--user-data-dir='+profile,'http://localhost:3000'];}
 app=spawn(exe,args,{cwd:root,env,windowsHide:true,stdio:'ignore'});
 const actualProfile=electron?path.join(folder,'profile'):profile;let port;
 await wait(async()=>{try{port=(await readFile(path.join(actualProfile,'DevToolsActivePort'),'utf8')).split('\n')[0];return port;}catch{return false;}},'porta');
 let target;await wait(async()=>{target=(await(await fetch('http://127.0.0.1:'+port+'/json')).json()).find(x=>x.type==='page');return !!target;},'janela');socket=new WebSocket(target.webSocketDebuggerUrl);await new Promise(r=>socket.addEventListener('open',r,{once:true}));
 socket.addEventListener('message',e=>{const m=JSON.parse(e.data);if(m.id){m.error?pending.get(m.id)?.reject(Error(m.error.message)):pending.get(m.id)?.resolve(m.result);pending.delete(m.id);}if(m.method==='Runtime.exceptionThrown')errors.push(m.params.exceptionDetails);});await cdp('Runtime.enable');await cdp('Page.enable');await cdp('Emulation.setDeviceMetricsOverride',{width:1366,height:768,deviceScaleFactor:1,mobile:false});
 await wait(()=>ev('!!document.querySelector("#app-shell")'),'interface');
 await ev('window.selectionErrors=[];window.addEventListener("unhandledrejection",e=>selectionErrors.push(String(e.reason)));window.addEventListener("error",e=>selectionErrors.push(e.message));');
 if(process.env.NUP_POLICY_ONLY==='1'){
  await testarPoliticasEscalaNup({ev,wait,ready,set,screenshot,folder});assert.deepEqual(await ev('selectionErrors'),[]);assert.deepEqual(errors,[]);console.log('PASSOU: políticas de escala N-up, PDFs, prévia e confirmações em '+(electron?'Electron':'Chrome')+'. Sem impressão.');
 }else{
 for(const count of [1,4,5,8,10]){
  await ev(`(async()=>{const{PDFDocument}=await import('/vendor/pdf-lib/pdf-lib.js'),p=await PDFDocument.create();for(let i=1;i<=${count};i++)p.addPage([105*72/25.4,148*72/25.4]).drawText('PAGE '+i,{x:20,y:100});const b=await p.save();window.sourceBefore=b.slice();const{PDFWorkspace}=await import('/js/tools/pdf-workspace.js');window.w=new PDFWorkspace();await w.abrir(new File([b],'selection-${count}.pdf'));const{JanelaMontagem}=await import('/js/tools/pdf-imposition-ui.js');window.testRoot=document.createElement('div');testRoot.className='pdf-editor';document.body.append(testRoot);window.u=new JanelaMontagem(testRoot,{snapshot:()=>w.obterBytes()});u.get('modoMarcas').value='sem';await u.open();window.sourceBefore=u.originalBytes.slice();})()`);await ready();
  await set('modo','nup');assert.equal(await ev('u.get("escalaModo").value'),'real');await set('escalaModo','reduzir');await set('modoMarcas','sem');assert.equal(await ev('u.get("nupSelecao").value'),'todas');assert.equal(await ev('u.get("paginas").checkVisibility()'),false);assert.equal(await ev('u.get("quantidade").checkVisibility()'),false);assert.equal(await ev('u.get("grade").checkVisibility()'),true);assert.match(await ev('u.get("total-paginas").textContent'),new RegExp(count+' página'));
  for(const distribution of ['celulas','compacto']){await set('nupDistribuicao',distribution);for(const capacity of [2,4,6,8]){await set('grade',capacity);await inspect(count,capacity,`${count}-paginas-${distribution}-${capacity}up`);}}
  await set('nupDistribuicao','celulas');await set('grade','4');await ev('u.get("avancadas").open=true');await set('nupSelecao','especificas');assert.equal(await ev('u.get("paginas").checkVisibility()'),true);
  await set('paginas',`1-${count}`);await ev('window.previousResult=u.result;window.previousCanvas=u.dialog.querySelector("canvas");window.previousPixels=previousCanvas.toDataURL();');
  for(const invalid of [String(count+1),`1-${count+1}`,'0','-1','abc','2-1','','1,,2']){
   await ev(`u.get('paginas').value=${JSON.stringify(invalid)};u.get('paginas').dispatchEvent(new Event('input',{bubbles:true}));`);await sleep(310);
   const state=await ev('({same:u.result===previousResult,canvas:u.dialog.querySelector("canvas")===previousCanvas,pixels:u.dialog.querySelector("canvas").toDataURL()===previousPixels,busy:u.busy,canSave:u.canOutput(),status:u.get("status").textContent})');assert.equal(state.same,true);assert.equal(state.canvas,true);assert.equal(state.pixels,true);assert.equal(state.busy,false);assert.equal(state.canSave,false);assert.match(state.status,/Prévia anterior mantida/);
  }
  await screenshot(`erro-selecao-${count}`);await set('paginas',`1-${count}`);await inspect(count,4,`${count}-paginas-selecao-recuperada`);
  // Um resultado pendente jamais pode reaparecer após entrada inválida.
  await ev(`u.get('orientacao').value='paisagem';u.get('orientacao').dispatchEvent(new Event('input',{bubbles:true}));u.get('paginas').value='${count+1}';u.get('paginas').dispatchEvent(new Event('input',{bubbles:true}));`);await sleep(500);assert.equal(await ev('u.busy'),false);assert.equal(await ev('u.canOutput()'),false);
  await set('nupSelecao','todas');
  await ev(`for(let i=0;i<30;i++){u.get('nupDistribuicao').value=i%2?'compacto':'celulas';u.get('orientacao').value=i%2?'paisagem':'retrato';u.get('grade').value=i%2?'2':'4';u.controls();u.schedule();}`);await ready();await inspect(count,2,`${count}-paginas-alternancia-rapida`);
  await screenshot(`nup-todas-${count}`);await ev('u.destroy();testRoot.remove()');
 }
 // Fonte real do workspace: reabrir após exclusão, união e substituição captura o total vigente.
 await ev(`(async()=>{const{PDFDocument}=await import('/vendor/pdf-lib/pdf-lib.js'),p=await PDFDocument.create();for(let i=1;i<=5;i++)p.addPage([180,250]).drawText('PAGE '+i,{x:20,y:100});const{PDFWorkspace}=await import('/js/tools/pdf-workspace.js');window.w=new PDFWorkspace();await w.abrir(new File([await p.save()],'edicoes.pdf'));const{JanelaMontagem}=await import('/js/tools/pdf-imposition-ui.js');window.testRoot=document.createElement('div');testRoot.className='pdf-editor';document.body.append(testRoot);window.u=new JanelaMontagem(testRoot,{snapshot:()=>w.obterBytes()});u.get('modoMarcas').value='sem';await u.open();})()`);await ready();await set('modo','nup');assert.equal(await ev('u.result.plan.quantidade'),5);
 await ev('u.close();w.excluirPaginas([5]);');await ev('u.open()');await ready();assert.equal(await ev('u.totalPaginas'),4);assert.deepEqual(await ev('u.result.config.paginas'),[1,2,3,4]);assert.equal(await ev('w.alterado'),true);
 await set('paginas','4');await ev('u.close();w.excluirPaginas([4]);');await ev('u.open()');await wait(()=>ev('!!u.selectionError'),'seleção antiga inválida após exclusão');assert.equal(await ev('u.totalPaginas'),3);assert.equal(await ev('u.get("paginas").value'),'4');assert.equal(await ev('u.canOutput()'),false);assert.equal(await ev('u.busy'),false);await set('nupSelecao','todas');assert.deepEqual(await ev('u.result.config.paginas'),[1,2,3]);
 await ev(`(async()=>{await u.close();const{PDFDocument}=await import('/vendor/pdf-lib/pdf-lib.js'),p=await PDFDocument.create();p.addPage([180,250]);await w.adicionarArquivos([new File([await p.save()],'nova.pdf')]);await u.open();})()`);await ready();assert.equal(await ev('u.totalPaginas'),4);assert.equal(await ev('u.result.plan.quantidade'),4);
 await set('modo','repeat');await set('paginas','3');await set('quantidade','7');assert.equal(await ev('u.result.plan.quantidade'),7);assert.equal(await ev('u.result.plan.folhas.flat().every(p=>p.origem===3)'),true);await set('modo','nup');assert.equal(await ev('u.get("nupSelecao").value'),'todas');assert.equal(await ev('u.result.plan.quantidade'),4);
 await ev(`(async()=>{await u.close();const{PDFDocument}=await import('/vendor/pdf-lib/pdf-lib.js'),p=await PDFDocument.create();p.addPage([180,250]);await w.abrir(new File([await p.save()],'outro.pdf'));await u.open();})()`);await ready();assert.equal(await ev('u.totalPaginas'),1);assert.deepEqual(await ev('u.result.config.paginas'),[1]);await ev('u.destroy();testRoot.remove()');
 assert.deepEqual(await ev('selectionErrors'),[]);assert.deepEqual(errors,[]);
 await writeFile(path.join(folder,'results.json'),JSON.stringify({passed:true,environment:electron?'Electron development':'Chrome',reports,workspaceEdits:true,invalidRetainsCanvasAndBlocksOutput:true,unhandled:errors,noPrinting:true},null,2));
 console.log(`PASSOU: N-up automático, ${reports.length} PDFs/planos, 1/4/5/8/10 páginas, seleções inválidas/recuperação, canvas preservado, alternância rápida, edição/substituição do workspace, Step & Repeat; sem impressão.`);
 }
}catch(e){try{await screenshot('failure');console.error(await ev('({status:u?.get("status").textContent,busy:u?.busy,total:u?.totalPaginas})'));}catch{}throw e;}
finally{socket?.close();app?.kill();server?.kill();}
