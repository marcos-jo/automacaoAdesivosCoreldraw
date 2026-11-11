import {testarMontagemReal} from './helpers/pdf-imposition-real-browser.mjs';
import {testarExcessoMontagem} from './helpers/pdf-imposition-overflow-browser.mjs';
import {testarInterfaceBasicaMontagem} from './helpers/pdf-imposition-basic-browser.mjs';
import {testarEscalasMontagem} from './helpers/pdf-imposition-scale-browser.mjs';
import {testarNupCompacto} from './helpers/pdf-imposition-compact-browser.mjs';
import {testarLivreto} from './helpers/pdf-booklet-browser.mjs';
import {testarCompatibilidadeMontagem} from './helpers/pdf-imposition-compatibility-browser.mjs';
import {spawn,execFile} from 'node:child_process';
import {promisify} from 'node:util';
import {readFile,writeFile,mkdir,unlink,readdir} from 'node:fs/promises';
import {setTimeout as pause} from 'node:timers/promises';
import path from 'node:path';
import assert from 'node:assert/strict';
import {createRequire} from 'node:module';
import {PDFCombinacao} from './helpers/pdf-combine-modules.mjs';
import {fixtures} from './helpers/pdf-combine-fixtures.mjs';
const desktop=process.env.IMPOSITION_DESKTOP==='1',folder=path.resolve(process.env.IMPOSITION_RESULTS||'.desktop-test/booklet24/'+(desktop?'installed':'browser'));await mkdir(folder,{recursive:true});
const currentVersion=createRequire(import.meta.url)('../electron-builder.hold-native-test.cjs').extraMetadata.version;
const draft=new PDFCombinacao();await draft.add(await fixtures());const bytes=await draft.build(),fixture=path.join(folder,'São José imagens e PDFs.pdf');await writeFile(fixture,bytes);
const exec=promisify(execFile),ps=path.join(process.env.SystemRoot,'System32/WindowsPowerShell/v1.0/powershell.exe'),quote=value=>"'"+value.replaceAll("'","''")+"'";
const powershell=async command=>(await exec(ps,['-NoProfile','-NonInteractive','-Command',command],{windowsHide:true,timeout:50000})).stdout.trim();
const wait=async(fn,label)=>{for(let i=0;i<600;i++){if(await fn())return;await pause(100);}throw Error('Espera esgotada '+label);};
let app,server,socket,id=0;const pending=new Map(),errors=[];
const cdp=(method,params={})=>new Promise((resolve,reject)=>{const n=++id,timer=setTimeout(()=>{pending.delete(n);reject(Error(method));},90000);pending.set(n,{resolve:r=>{clearTimeout(timer);resolve(r);},reject});socket.send(JSON.stringify({id:n,method,params}));});
const ev=async expression=>{const r=await cdp('Runtime.evaluate',{expression,returnByValue:true,awaitPromise:true,userGesture:true});if(r.exceptionDetails)throw Error(r.exceptionDetails.exception?.description||r.exceptionDetails.text);return r.result.value;};
const click=selector=>ev(`document.querySelector(${JSON.stringify(selector)}).click()`);
const key=async letter=>{await cdp('Input.dispatchKeyEvent',{type:'rawKeyDown',key:letter,code:'Key'+letter.toUpperCase(),windowsVirtualKeyCode:letter.toUpperCase().charCodeAt(0),modifiers:2});await cdp('Input.dispatchKeyEvent',{type:'keyUp',key:letter,code:'Key'+letter.toUpperCase(),windowsVirtualKeyCode:letter.toUpperCase().charCodeAt(0),modifiers:2});};
const screenshot=async name=>{await pause(250);const r=await cdp('Page.captureScreenshot');await writeFile(path.join(folder,name+'.png'),Buffer.from(r.data,'base64'));};
try{
 let profile,exe,args;
 if(desktop){
  const reg="HKCU:/Software/Microsoft/Windows/CurrentVersion/Uninstall/90ff2ff0-f013-571c-b109-7edb265dacc8";
  const info=JSON.parse(await powershell(`Get-ItemProperty '${reg}'|Select-Object DisplayVersion,UninstallString|ConvertTo-Json -Compress`));
  assert.equal(await powershell("@(Get-Process '4PrintShops Hold Test' -ErrorAction SilentlyContinue).Count"),'0');
  const installation=path.dirname(info.UninstallString.match(/^"([^"]+)"/)[1]);assert.equal(path.basename(info.UninstallString.match(/^"([^"]+)"/)[1]),'Uninstall 4PrintShops Hold Test.exe');
  if(process.env.IMPOSITION_INSTALL==='1')await new Promise((resolve,reject)=>{const p=spawn(path.resolve(`dist-desktop/hold-native-test/4PrintShops-Hold-Test-${currentVersion}-Windows-x64-Setup.exe`),['/S'],{windowsHide:true,stdio:'ignore'});p.on('error',reject);p.on('exit',code=>code===0?resolve():reject(Error('Setup '+code)));});
  assert.equal(await powershell(`(Get-ItemProperty '${reg}').DisplayVersion`),currentVersion);
  profile=path.join(process.env.APPDATA,'4PrintShops Hold Test');await unlink(path.join(profile,'DevToolsActivePort')).catch(()=>{});exe=path.join(installation,'4PrintShops Hold Test.exe');args=['--remote-debugging-port=0','--disable-renderer-backgrounding','--open-pdf',fixture];
 }else{
  try{await fetch('http://localhost:3000');}catch{server=spawn(process.execPath,['app.js'],{windowsHide:true,stdio:'ignore'});}await wait(async()=>{try{return(await fetch('http://localhost:3000')).ok;}catch{return false;}},'servidor');profile=path.join(folder,'profile-'+Date.now());await mkdir(profile);exe='C:/Program Files/Google/Chrome/Application/chrome.exe';args=['--headless=new','--disable-gpu','--no-first-run','--remote-debugging-port=0','--user-data-dir='+profile,'http://localhost:3000'];
 }
 const env={...process.env};delete env.ELECTRON_RUN_AS_NODE;app=spawn(exe,args,{env,windowsHide:true,stdio:'ignore'});
 let port;await wait(async()=>{try{port=(await readFile(path.join(profile,'DevToolsActivePort'),'utf8')).split('\n')[0];return !!port;}catch{return false;}},'porta');let target;await wait(async()=>{try{target=(await(await fetch('http://127.0.0.1:'+port+'/json')).json()).find(t=>t.type==='page');return !!target;}catch{return false;}},'janela');socket=new WebSocket(target.webSocketDebuggerUrl);await new Promise(r=>socket.addEventListener('open',r,{once:true}));socket.addEventListener('message',event=>{const m=JSON.parse(event.data);if(m.id){m.error?pending.get(m.id)?.reject(Error(m.error.message)):pending.get(m.id)?.resolve(m.result);pending.delete(m.id);}if(m.method==='Runtime.exceptionThrown')errors.push(m.params.exceptionDetails.text);});await cdp('Runtime.enable');await cdp('Page.enable');await cdp('Page.bringToFront');await cdp('Emulation.setDeviceMetricsOverride',{width:1366,height:768,deviceScaleFactor:1,mobile:false});
 if(!desktop){await wait(()=>ev("!!document.querySelector('[data-tool=pdfs]')"),'início');await click('[data-tool=pdfs]');await ev(`(()=>{const data=Uint8Array.from(atob(${JSON.stringify(Buffer.from(bytes).toString('base64'))}),c=>c.charCodeAt(0)),d=new DataTransfer();d.items.add(new File([data],'combinado.pdf'));const i=document.querySelector('#pdf-arquivo');i.files=d.files;i.dispatchEvent(new Event('change',{bubbles:true}));})()`);}
 await wait(()=>ev("!!document.querySelector('#pdf-visualizacao canvas')&&!document.querySelector('#pdf-arquivo').disabled"),'PDF');await click('#pdf-girar-direita');await wait(()=>ev("!document.querySelector('#pdf-arquivo').disabled"),'giro');
 const statusBefore=await ev("document.querySelector('#pdf-alteracoes-pendentes').hidden");assert.equal(statusBefore,false);
 if(process.env.BOOKLET_ONLY!=='1'){
 await click('#app-menu-editar');await ev("document.querySelector('[data-command=pdf-montagem]').click()");
 await wait(()=>ev("!!document.querySelector('.pdf-montagem[open]')"),'janela montagem');
 const field=id=>`.pdf-montagem [data-id=${id}]`;
 const set=async(id,value)=>ev(`(()=>{const el=document.querySelector(${JSON.stringify(field(id))});el.value=${JSON.stringify(String(value))};el.dispatchEvent(new Event('input',{bubbles:true}));})()`);
 await set('rotacao','90');await set('pecaLargura','90');await set('pecaAltura','');
 await testarInterfaceBasicaMontagem({ev,cdp,wait,screenshot,folder});
 await wait(()=>ev("!document.querySelector('.pdf-montagem [data-id=salvar]').disabled"),'prévia vetorial');
 await ev("(async()=>{const {aplicarTema}=await import('/js/app/theme.js');aplicarTema('light');})()");await screenshot('montagem-claro');
 await ev("(async()=>{const {aplicarTema}=await import('/js/app/theme.js');aplicarTema('dark');})()");await screenshot('montagem-escuro');
 const geometry=await ev("(()=>{const d=document.querySelector('.pdf-montagem'),b=d.querySelector('[data-id=salvar]').getBoundingClientRect();return {bottom:b.bottom,height:innerHeight,canvas:d.querySelector('canvas').width,text:d.querySelector('[data-id=status]').textContent};})()");assert.ok(geometry.bottom<=geometry.height);assert.ok(geometry.canvas>0);assert.match(geometry.text,/^Folha \d+ × \d+ mm/);
 await set('modoMarcas','grupo');await wait(()=>ev("!document.querySelector('.pdf-montagem [data-id=salvar]').disabled"),'contorno externo');await screenshot('grupo-externo');
 await set('modoMarcas','bordas');await set('colunas','2');await set('linhas','1');
 await wait(()=>ev("!document.querySelector('.pdf-montagem [data-id=salvar]').disabled"),'contorno com divisões');await screenshot('contorno-divisoes-2x1');
 assert.equal(await ev("document.querySelector('.pdf-montagem [data-id=modoMarcas]').value"),'bordas');
 if(desktop){
    const saveConsent=e=>{const m=JSON.parse(e.data);if(m.method==='Page.javascriptDialogOpening'&&m.params.message.startsWith('Conferir montagem antes de gerar o PDF?'))cdp('Page.handleJavaScriptDialog',{accept:true}).catch(()=>{});};socket.addEventListener('message',saveConsent);
    const destination=path.join(folder,'Montagem A3.pdf');await unlink(destination).catch(()=>{});await click(field('salvar'));
    const nativeSource=await readFile(new URL('./helpers/windows-save-dialog.ps1',import.meta.url),'utf8');
    await powershell(`& {${nativeSource}} -DialogProcess ${app.pid} -Destination ${quote(destination)} -Action save`);
    await wait(()=>ev("document.querySelector('.pdf-montagem [data-id=status]').textContent.startsWith('PDF de montagem salvo')"),'salvamento nativo');
    socket.removeEventListener('message',saveConsent);
    const {PDFDocument}=await import('../public/vendor/pdf-lib/pdf-lib.js');const saved=await PDFDocument.load(await readFile(destination));assert.ok(saved.getPageCount()>1);assert.ok(Math.abs(saved.getPage(0).getWidth()-297*72/25.4)<1e-6);
 }
 await set('modo','nup');await set('escalaModo','reduzir');await set('grade','4');await set('paginas','1-4');
 await wait(()=>ev("!document.querySelector('.pdf-montagem [data-id=salvar]').disabled"),'N-up');await screenshot('nup-4');
 // Testa cancelamento de prévias antigas e recuperação de uma grade inválida.
 await set('espacoX','1000');await wait(()=>ev("!document.querySelector('.pdf-montagem [data-id=excesso]').hidden&&document.querySelector('.pdf-montagem canvas').width>0"),'grade excedente visível');
 assert.equal(await ev("document.querySelector('.pdf-montagem [data-id=salvar]').disabled"),true);
 for(const value of ['1','2','3','0'])await set('espacoX',value);
 await wait(()=>ev("!document.querySelector('.pdf-montagem [data-id=salvar]').disabled"),'recuperação da prévia');
 await click(field('cancelar'));await wait(()=>ev("!document.querySelector('.pdf-montagem').open"),'fechamento');assert.equal(await ev("document.querySelector('#pdf-alteracoes-pendentes').hidden"),false);
 await click('#app-menu-editar');await ev("document.querySelector('[data-command=pdf-montagem]').click()");await wait(()=>ev("!document.querySelector('.pdf-montagem [data-id=salvar]').disabled"),'reabertura');
 if(desktop){
   const accept=event=>{const m=JSON.parse(event.data);if(m.method==='Page.javascriptDialogOpening')cdp('Page.handleJavaScriptDialog',{accept:m.params.message.startsWith('Imprimir a montagem já imposta?')}).catch(()=>{});};
   socket.addEventListener('message',accept);await click(field('imprimir'));
   await wait(()=>ev("document.querySelector('#print-config-dialog').open&&!document.querySelector('#print-config-close').disabled"),'impressão da montagem');socket.removeEventListener('message',accept);
   await ev("document.querySelector('#print-config-printer').value='Microsoft Print to PDF';document.querySelector('#print-config-printer').dispatchEvent(new Event('change'))");
   await wait(()=>ev("!document.querySelector('#print-config-close').disabled"),'seleção virtual');
   assert.equal(await ev("document.querySelector('#print-config-printer').value"),'Microsoft Print to PDF');
   assert.equal(await ev("document.querySelector('#print-config-scale').value"),'100');
   const virtual=path.join(folder,'Montagem virtual.pdf');await unlink(virtual).catch(()=>{});
   await click('#print-config-print');
   const nativeSource=(await readFile(new URL('./helpers/windows-save-dialog.ps1',import.meta.url),'utf8')).replaceAll('Salvar cópia PDF','Salvar saída da impressora PDF');
   await powershell(`& {${nativeSource}} -DialogProcess ${app.pid} -Destination ${quote(virtual)} -Action save`);
   await wait(()=>ev("!document.querySelector('#print-config-close').disabled"),'impressão virtual');assert.match(await ev("document.querySelector('#print-config-status').textContent"),/Trabalho.*enviado/);
   const {PDFDocument}=await import('../public/vendor/pdf-lib/pdf-lib.js');assert.equal((await PDFDocument.load(await readFile(virtual))).getPageCount(),1);
   await click('#print-config-close');assert.equal(await ev("document.querySelector('#pdf-alteracoes-pendentes').hidden"),false);
 }else await click(field('cancelar'));
 if(!desktop){
   const sample=await ev(`(async()=>{const {PDFDocument,rgb}=await import('/vendor/pdf-lib/pdf-lib.js');const {gerarMontagem}=await import('/js/tools/pdf-imposition-engine.js');const source=await PDFDocument.create();const page=source.addPage([90*72/25.4,50*72/25.4]);page.drawRectangle({x:0,y:0,width:90*72/25.4,height:50*72/25.4,color:rgb(.94,.94,1)});page.drawText('Cartao 90 x 50 mm',{x:30,y:65,size:14});const output=await gerarMontagem(await source.save(),{modo:'repeat',paginas:[1],largura:297,altura:210,automatico:false,colunas:2,linhas:1,quantidade:2,modoMarcas:'bordas',espacoX:0,espacoY:0});const pdfjs=await import('/vendor/pdfjs/pdf.js');const task=pdfjs.getDocument({data:output.bytes.slice()});try{const doc=await task.promise,page=await doc.getPage(1),viewport=page.getViewport({scale:2}),canvas=document.createElement('canvas');canvas.width=Math.ceil(viewport.width);canvas.height=Math.ceil(viewport.height);await page.render({canvas,canvasContext:canvas.getContext('2d'),viewport,background:'white',intent:'print'}).promise;let binary='';for(let i=0;i<output.bytes.length;i+=32768)binary+=String.fromCharCode(...output.bytes.subarray(i,i+32768));const result={pdf:btoa(binary),png:canvas.toDataURL('image/png').split(',')[1],marks:output.plan.cortesGrupo[0].length};canvas.width=0;return result;}finally{await task.destroy();}})()`);
   assert.equal(sample.marks,10);await writeFile(path.join(folder,'Contorno divisoes 2x1.pdf'),Buffer.from(sample.pdf,'base64'));await writeFile(path.join(folder,'Contorno divisoes 2x1.png'),Buffer.from(sample.png,'base64'));
   const comparison=await ev(`(async()=>{const {PDFDocument,PDFName,rgb}=await import('/vendor/pdf-lib/pdf-lib.js');const {gerarMontagem}=await import('/js/tools/pdf-imposition-engine.js');const pdf=await PDFDocument.create();const page=pdf.addPage([72,72]);page.drawRectangle({x:4,y:8,width:40,height:40,color:rgb(1,0,0),opacity:.5});page.drawRectangle({x:20,y:20,width:40,height:40,color:rgb(0,0,1),opacity:.5});page.drawText('ABC',{x:5,y:60,size:9});page.node.set(PDFName.of('Group'),pdf.context.obj({S:'Transparency',CS:'DeviceRGB',I:true}));const source=await pdf.save();const result=await gerarMontagem(source,{modo:'repeat',paginas:[1],largura:101.6,altura:101.6,automatico:false,colunas:1,linhas:1,quantidade:1,centralizar:false,margens:{esquerda:0,direita:0,superior:0,inferior:0},marcas:false});const pdfjs=await import('/vendor/pdfjs/pdf.js');async function pixels(bytes){const task=pdfjs.getDocument({data:bytes.slice()});try{const doc=await task.promise,p=await doc.getPage(1),view=p.getViewport({scale:2}),canvas=document.createElement('canvas');canvas.width=Math.ceil(view.width);canvas.height=Math.ceil(view.height);const ctx=canvas.getContext('2d');await p.render({canvas,canvasContext:ctx,viewport:view,background:'white',intent:'print'}).promise;const data=ctx.getImageData(0,0,144,144).data;canvas.width=0;return data;}finally{await task.destroy();}}const a=await pixels(source),b=await pixels(result.bytes);let max=0,different=0;for(let i=0;i<a.length;i++){const d=Math.abs(a[i]-b[i]);max=Math.max(max,d);if(d)different++;}return {components:a.length,max,different};})()`);
   assert.ok(comparison.max<=1,JSON.stringify(comparison));await writeFile(path.join(folder,'fidelity.json'),JSON.stringify(comparison,null,2));
   const result=await ev(`(async()=>{const {PDFDocument,PDFName,rgb}=await import('/vendor/pdf-lib/pdf-lib.js');const {gerarMontagem}=await import('/js/tools/pdf-imposition-engine.js');const source=await PDFDocument.create();const page=source.addPage([90*72/25.4,50*72/25.4]);page.drawText('Montagem vetorial 90 x 50 mm',{x:10,y:100,size:12});page.drawRectangle({x:10,y:20,width:100,height:50,color:rgb(.5,0,.8),opacity:.5});page.node.set(PDFName.of('Group'),source.context.obj({S:'Transparency',CS:'DeviceRGB',I:true}));const bytes=await source.save();const output=await gerarMontagem(bytes,{modo:'repeat',largura:297,altura:420,automatico:true,quantidade:25,paginas:[1],marcas:true});const pdfjs=await import('/vendor/pdfjs/pdf.js');const task=pdfjs.getDocument({data:output.bytes.slice()});const pdf=await task.promise;const text=await (await pdf.getPage(1)).getTextContent();const count=text.items.filter(t=>t.str.includes('Montagem vetorial')).length;await task.destroy();let binary='';for(let i=0;i<output.bytes.length;i+=32768)binary+=String.fromCharCode(...output.bytes.subarray(i,i+32768));return {count,pages:output.plan.totalFolhas,base64:btoa(binary)};})()`);
   assert.equal(result.count,12);assert.equal(result.pages,3);await writeFile(path.join(folder,'Cartoes 90x50 A3.pdf'),Buffer.from(result.base64,'base64'));
 }
 const stress=await ev(`(async()=>{const {JanelaMontagem}=await import('/js/tools/pdf-imposition-ui.js');const data=Uint8Array.from(atob(${JSON.stringify(Buffer.from(bytes).toString('base64'))}),c=>c.charCodeAt(0));const root=document.createElement('div');document.body.append(root);const ui=new JanelaMontagem(root,{snapshot:async()=>data.slice(),imprimir:async()=>{throw Error('Não imprimir no ensaio de falha');}});const originalBytes=data.slice();const original=ui.render.bind(ui);let starts=0,aborts=0;const sleep=ms=>new Promise(r=>setTimeout(r,ms));async function until(fn){for(let i=0;i<500;i++){if(fn())return;await sleep(30);}throw Error('Prévia não recuperou');}ui.get('escalaModo').value='medidas';ui.get('pecaLargura').value='90';ui.render=async()=>{throw Error('Falha de prévia injetada');};await ui.open();await until(()=>!ui.get('tentar').hidden);const error=ui.get('status').textContent;if(!error.includes('Falha de prévia'))throw Error(error);ui.render=original;ui.get('tentar').click();await until(()=>!ui.busy);if(ui.get('salvar').disabled)throw Error('Exportação presa');ui.get('pecaLargura').value='90';ui.get('colunas').value='2';ui.get('linhas').value='1';ui.get('modoMarcas').value='grupo';ui.render=async(signal,result,number)=>{starts++;signal.addEventListener('abort',()=>aborts++,{once:true});if(starts<=8)setTimeout(()=>{const field=ui.get(starts===4?'guias':'espacoX');if(starts===4)field.checked=!field.checked;else field.value=String(starts%3);field.dispatchEvent(new Event('input',{bubbles:true}));},5);return original(signal,result,number);};ui.schedule();await until(()=>starts>=9&&!ui.busy);const result={starts,aborts,canSave:!ui.get('salvar').disabled,canPrint:!ui.get('imprimir').disabled,cuts:ui.result.plan.cortesGrupo[0].length,quantity:ui.result.plan.quantidade,sourceUnchanged:data.every((v,i)=>v===originalBytes[i])};let binary='';for(let i=0;i<ui.result.bytes.length;i+=32768)binary+=String.fromCharCode(...ui.result.bytes.subarray(i,i+32768));result.pdf=btoa(binary);await ui.destroy();root.remove();return result;})()`);
 assert.ok(stress.starts>=9);assert.ok(stress.aborts>=8);assert.equal(stress.canSave,true);assert.equal(stress.canPrint,true);assert.equal(stress.cuts,8);assert.equal(stress.sourceUnchanged,true);await writeFile(path.join(folder,'Grupo apos alteracoes.pdf'),Buffer.from(stress.pdf,'base64'));delete stress.pdf;await writeFile(path.join(folder,'preview-stress.json'),JSON.stringify(stress,null,2));
 await testarExcessoMontagem({ev,wait,socket,cdp,folder,screenshot});
 await testarEscalasMontagem({ev,wait,socket,cdp,folder,screenshot});
 await testarCompatibilidadeMontagem({ev,click,set,wait,socket,cdp,folder,screenshot,desktop,powershell,quote,app});
 await testarNupCompacto({ev,wait,socket,cdp,folder,screenshot,desktop,powershell,quote,app});
 await testarMontagemReal({ev,socket,cdp,folder,screenshot});
 }
 await testarLivreto({ev,wait,socket,cdp,folder,screenshot,desktop,powershell,quote,app});
 assert.deepEqual(errors,[]);console.log('PASSOU montagem '+(desktop?'instalada':'Chrome')+': prévia, grupo, falha/retry, cancelamentos sucessivos, salvar, temas, origem pendente. Nenhuma impressão física.');
}catch(error){console.error(error);console.error(errors);try{console.error(await ev("Array.from(document.querySelectorAll('#print-config-dialog')).map(d=>({open:d.open,closeDisabled:d.querySelector('#print-config-close')?.disabled,status:d.querySelector('#print-config-status')?.textContent}))"));console.error(await ev("document.querySelector('.pdf-montagem [data-id=status]')?.textContent"));await screenshot('failure');}catch{}process.exitCode=1;}
finally{socket?.close();app?.kill();server?.kill();}
