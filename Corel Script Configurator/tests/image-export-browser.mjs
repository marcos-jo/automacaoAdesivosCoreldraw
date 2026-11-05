import {spawn,execFile} from 'node:child_process';
import {promisify} from 'node:util';
import {readFile,writeFile,mkdir,unlink,readdir} from 'node:fs/promises';
import {setTimeout as pause} from 'node:timers/promises';
import path from 'node:path';
import assert from 'node:assert/strict';
import {PDFCombinacao} from './helpers/pdf-combine-modules.mjs';
import {fixtures} from './helpers/pdf-combine-fixtures.mjs';
const desktop=process.env.IMAGES_DESKTOP==='1',folder=path.resolve('.desktop-test/images16/'+(desktop?'installed':'browser'));await mkdir(folder,{recursive:true});
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
  if(process.env.IMAGES_INSTALL==='1')await new Promise((resolve,reject)=>{const p=spawn(path.resolve('dist-desktop/hold-native-test/4PrintShops-Hold-Test-1.0.0-experimental.16-Windows-x64-Setup.exe'),['/S'],{windowsHide:true,stdio:'ignore'});p.on('error',reject);p.on('exit',code=>code===0?resolve():reject(Error('Setup '+code)));});
  assert.equal(await powershell(`(Get-ItemProperty '${reg}').DisplayVersion`),'1.0.0-experimental.16');
  profile=path.join(process.env.APPDATA,'4PrintShops Hold Test');await unlink(path.join(profile,'DevToolsActivePort')).catch(()=>{});exe=path.join(installation,'4PrintShops Hold Test.exe');args=['--remote-debugging-port=0','--disable-renderer-backgrounding','--open-pdf',fixture];
 }else{
  try{await fetch('http://localhost:3000');}catch{server=spawn(process.execPath,['app.js'],{windowsHide:true,stdio:'ignore'});}await wait(async()=>{try{return(await fetch('http://localhost:3000')).ok;}catch{return false;}},'servidor');profile=path.join(folder,'profile-'+Date.now());await mkdir(profile);exe='C:/Program Files/Google/Chrome/Application/chrome.exe';args=['--headless=new','--disable-gpu','--no-first-run','--remote-debugging-port=0','--user-data-dir='+profile,'http://localhost:3000'];
 }
 const env={...process.env};delete env.ELECTRON_RUN_AS_NODE;app=spawn(exe,args,{env,windowsHide:true,stdio:'ignore'});
 let port;await wait(async()=>{try{port=(await readFile(path.join(profile,'DevToolsActivePort'),'utf8')).split('\n')[0];return !!port;}catch{return false;}},'porta');let target;await wait(async()=>{try{target=(await(await fetch('http://127.0.0.1:'+port+'/json')).json()).find(t=>t.type==='page');return !!target;}catch{return false;}},'janela');socket=new WebSocket(target.webSocketDebuggerUrl);await new Promise(r=>socket.addEventListener('open',r,{once:true}));socket.addEventListener('message',event=>{const m=JSON.parse(event.data);if(m.id){m.error?pending.get(m.id)?.reject(Error(m.error.message)):pending.get(m.id)?.resolve(m.result);pending.delete(m.id);}if(m.method==='Runtime.exceptionThrown')errors.push(m.params.exceptionDetails.text);});await cdp('Runtime.enable');await cdp('Page.enable');await cdp('Page.bringToFront');await cdp('Emulation.setDeviceMetricsOverride',{width:1366,height:768,deviceScaleFactor:1,mobile:false});
 if(!desktop){await wait(()=>ev("!!document.querySelector('[data-tool=pdfs]')"),'início');await click('[data-tool=pdfs]');await ev(`(()=>{const data=Uint8Array.from(atob(${JSON.stringify(Buffer.from(bytes).toString('base64'))}),c=>c.charCodeAt(0)),d=new DataTransfer();d.items.add(new File([data],'combinado.pdf'));const i=document.querySelector('#pdf-arquivo');i.files=d.files;i.dispatchEvent(new Event('change',{bubbles:true}));})()`);}
 await wait(()=>ev("!!document.querySelector('#pdf-visualizacao canvas')&&!document.querySelector('#pdf-arquivo').disabled"),'PDF');await click('#pdf-girar-direita');await wait(()=>ev("!document.querySelector('#pdf-arquivo').disabled"),'giro');
 // A ativação começa no canvas, fora dos campos em edição.
 await ev("document.activeElement?.blur()");await key('s');await wait(()=>ev("document.querySelector('#app-exportacao-dialogo').open"),'Ctrl+S');assert.equal(await ev("document.querySelector('#app-exportacao-formato').value"),'pdf');await click('#app-exportacao-cancelar');
 await ev("document.activeElement?.blur()");await key('e');await wait(()=>ev("document.querySelector('#app-exportacao-dialogo').open"),'Ctrl+E');await ev("document.querySelector('#app-exportacao-formato').value='png';document.querySelector('#app-exportacao-formato').dispatchEvent(new Event('change'))");await ev("(async()=>{const {aplicarTema}=await import('/js/app/theme.js');aplicarTema('light');})()");await screenshot('exportar-claro');
 await ev("(async()=>{const {aplicarTema}=await import('/js/app/theme.js');aplicarTema('dark');})()");await screenshot('exportar-escuro');await click('#app-exportacao-cancelar');
 if(!desktop){
  // Destino em memória somente no teste. PDF.js e os encoders são reais.
  const results=[];
  for(const config of [{format:'png',dpi:150,quality:.92,pages:[1]},{format:'png',dpi:300,quality:.92,pages:[1]},{format:'png',dpi:600,quality:.92,pages:[1]},{format:'jpeg',dpi:150,quality:.4,pages:[2]},{format:'jpeg',dpi:150,quality:.95,pages:[2]}]){
   const r=await ev(`(async()=>{const {exportImages}=await import('/js/tools/pdf-image-export.js');const {PDFDocument,degrees}=await import('/vendor/pdf-lib/pdf-lib.js');const pdf=await PDFDocument.load(Uint8Array.from(atob(${JSON.stringify(Buffer.from(bytes).toString('base64'))}),c=>c.charCodeAt(0)));pdf.getPage(0).setRotation(degrees(90));let output;await exportImages(await pdf.save(),${JSON.stringify(config)},{destination:{async write(index,data){output=data;},async finish(){return {saved:true};},async cancel(){throw Error('Cancelamento inesperado');}}});const blob=new Blob([output],{type:'image/${config.format}'}),image=await createImageBitmap(blob);let text='';for(let i=0;i<output.length;i+=32768)text+=String.fromCharCode(...output.subarray(i,i+32768));const result={width:image.width,height:image.height,bytes:output.length,base64:btoa(text)};image.close();return result;})()`);
   const ext=config.format==='png'?'png':'jpg';await writeFile(path.join(folder,`pagina-${config.pages[0]}-${config.dpi}-${Math.round(config.quality*100)}.${ext}`),Buffer.from(r.base64,'base64'));delete r.base64;results.push({...config,...r});
   if(config.pages[0]===1){assert.equal(r.width,Math.ceil(841.8898*config.dpi/72-1e-7));assert.equal(r.height,Math.ceil(595.2756*config.dpi/72-1e-7));}
  }
  const large=await ev(`(async()=>{const {PDFDocument}=await import('/vendor/pdf-lib/pdf-lib.js');const {exportImages}=await import('/js/tools/pdf-image-export.js');const pdf=await PDFDocument.create();for(let i=0;i<350;i++)pdf.addPage([40,60]);let count=0;return await exportImages(await pdf.save(),{format:'png',quality:.9,dpi:150,pages:Array.from({length:350},(_,i)=>i+1)},{destination:{async write(){count++;},async finish(){return {count};},async cancel(){}}});})()`);assert.equal(large.count,350);  assert.ok(results[3].bytes<results[4].bytes);await writeFile(path.join(folder,'images.json'),JSON.stringify(results,null,2));
  const canceled=await ev(`(async()=>{const {exportImages}=await import('/js/tools/pdf-image-export.js');const abort=new AbortController();let cleaned=false;try{await exportImages(Uint8Array.from(atob(${JSON.stringify(Buffer.from(bytes).toString('base64'))}),c=>c.charCodeAt(0)),{format:'png',quality:.9,dpi:150,pages:[1,2,3]},{signal:abort.signal,progress:()=>abort.abort(),destination:{write:async()=>{},finish:async()=>{},cancel:async()=>{cleaned=true;}}});}catch(error){return {name:error.name,cleaned};}})()`);assert.deepEqual(canceled,{name:'AbortError',cleaned:true});
 }else{
  await ev("document.activeElement?.blur()");await key('e');await wait(()=>ev("document.querySelector('#app-exportacao-dialogo').open"),'exportação nativa');await ev("document.querySelector('#app-exportacao-formato').value='png';document.querySelector('#app-exportacao-formato').dispatchEvent(new Event('change'));document.querySelector('#app-exportacao-alvo').value='all';document.querySelector('#app-exportacao-dpi').value='150'");await click('#app-exportacao-confirmar');
  const nativeSource=(await readFile(new URL('./helpers/windows-save-dialog.ps1',import.meta.url),'utf8')).replaceAll("'Abrir PDFs'","'Exportar imagens — escolher pasta'").replaceAll('(id==1001||id==1148)','(id==1001||id==1148||id==1152)').replaceAll("@('1001','1148')","@('1001','1148','1152')").replaceAll('label!="Open"','label!="Open"&&label!="Selecionar pasta"&&label!="Select Folder"').replaceAll("@('Salvar','Save','Abrir','Open')","@('Salvar','Save','Abrir','Open','Selecionar pasta','Select Folder')");
  await powershell(`& {${nativeSource}} -DialogProcess ${app.pid} -Destination ${quote(folder)} -Action open`);
  await wait(()=>ev("!document.querySelector('#app-exportacao-dialogo').open"),'exportação PNG');
  const status=await ev("document.querySelector('#pdf-status').textContent");
  const location=status.match(/imagem\(ns\) exportada\(s\) em (.+)\. PDF preservado\./)?.[1];assert.ok(location,status);assert.equal(path.dirname(location),folder);
  const outputNames=(await readdir(location)).sort();assert.equal(outputNames.length,7);assert.ok(outputNames.every(name=>name.endsWith('.png')));
  const first=await readFile(path.join(location,outputNames[0]));assert.equal(first.readUInt32BE(16),1754);assert.equal(first.readUInt32BE(20),1241);
  await writeFile(path.join(folder,'native-output.json'),JSON.stringify({location,names:outputNames},null,2));
  assert.equal(await ev("document.querySelector('#pdf-alteracoes-pendentes').hidden"),false);
  await ev("document.activeElement?.blur()");await key('p');await wait(()=>ev("document.querySelector('#print-config-dialog').open&&!document.querySelector('#print-config-close').disabled"),'Ctrl+P');await screenshot('atalho-impressao');await click('#print-config-close');await pause(200);
  await ev("document.activeElement?.blur()");await key('o');
  const openSource=await readFile(new URL('./helpers/windows-save-dialog.ps1',import.meta.url),'utf8');
  await powershell(`& {${openSource}} -DialogProcess ${app.pid} -Destination ${quote(fixture)} -Action open`);
  await wait(()=>ev("document.querySelector('#pdf-confirmar-descarte').open"),'Ctrl+O e proteção');await click('#pdf-continuar-editando');
  await wait(()=>ev("!document.querySelector('#pdf-arquivo').disabled"),'documento preservado');
  const beforeCancel=(await readdir(folder)).sort();
  await ev("document.activeElement?.blur()");await key('e');await wait(()=>ev("document.querySelector('#app-exportacao-dialogo').open"),'exportação cancelável');
  await ev("document.querySelector('#app-exportacao-formato').value='png';document.querySelector('#app-exportacao-formato').dispatchEvent(new Event('change'));document.querySelector('#app-exportacao-alvo').value='current';document.querySelector('#app-exportacao-dpi').value='600'");await click('#app-exportacao-confirmar');
  await powershell(`& {${nativeSource}} -DialogProcess ${app.pid} -Destination ${quote(folder)} -Action open`);
  await wait(()=>ev("document.querySelector('#app-exportacao-status').textContent.startsWith('Página')"),'renderização cancelável');await click('#app-exportacao-cancelar');
  await wait(()=>ev("!document.querySelector('#pdf-arquivo').disabled"),'cancelamento concluído');assert.deepEqual((await readdir(folder)).sort(),beforeCancel);assert.equal(await ev("document.querySelector('#pdf-alteracoes-pendentes').hidden"),false);
 }
 assert.equal(await ev("document.querySelector('#pdf-alteracoes-pendentes').hidden"),false);assert.deepEqual(errors,[]);console.log('PASSOU imagens/atalhos '+(desktop?'instalado':'Chrome')+'. Nenhum envio de impressão.');
}catch(error){console.error(error);console.error(errors);try{console.error(await ev("document.querySelector('#app-exportacao-status')?.textContent"));await screenshot('failure');}catch{}process.exitCode=1;}
finally{socket?.close();app?.kill();server?.kill();}
