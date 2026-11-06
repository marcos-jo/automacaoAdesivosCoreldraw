import assert from 'node:assert/strict';import {spawn} from 'node:child_process';
import {mkdir,readFile,writeFile,readdir} from 'node:fs/promises';import path from 'node:path';import {createHash} from 'node:crypto';
import {PDFDocument,PDFName} from '../public/vendor/pdf-lib/pdf-lib.js';
import {criarFixtureImpressao} from './helpers/pdf-print-fixture.mjs';
import {exercitarRedimensionamentoMarcado} from './helpers/pdf-tagged-resize-browser.mjs';
const electron=process.env.TAGGED_ELECTRON==='1',folder=path.resolve(process.env.TAGGED_RESULTS||`.desktop-test/midias29/${electron?'electron-final':'chrome-final'}`);await mkdir(folder,{recursive:true});
const realFile=process.env.TAGGED_REAL_PATH,real=realFile?await readFile(realFile):null,hash=b=>createHash('sha256').update(b).digest('hex'),originalHash=real&&hash(real);
const p=await PDFDocument.load(await criarFixtureImpressao({actualText:'Texto essencial preservado'}));p.getPage(0).setSize(720,540);const synthetic=await p.save();
const unsafe=await PDFDocument.load(synthetic);unsafe.catalog.lookup(PDFName.of('StructTreeRoot')).set(PDFName.of('ClassMap'),unsafe.context.obj({LayoutClass:{O:'Layout',BBox:[0,0,100,100]}}));
const fixtures=[{name:'marked-embedded-fonts-symbols-transparency',data:Buffer.from(synthetic).toString('base64')},...(real?[{name:'original-local-authorized',data:real.toString('base64')}]:[]),{name:'unsupported-layout-attributes',data:Buffer.from(await unsafe.save()).toString('base64'),blocked:true}];
let app,server,socket,id=0;const pending=new Map(),errors=[],sleep=ms=>new Promise(r=>setTimeout(r,ms));
const wait=async(fn,label)=>{for(let i=0;i<600;i++){if(await fn())return;await sleep(100);}throw Error('Espera: '+label);};
const cdp=(method,params={})=>new Promise((resolve,reject)=>{const n=++id,t=setTimeout(()=>{pending.delete(n);reject(Error(method));},240000);pending.set(n,{resolve:r=>{clearTimeout(t);resolve(r);},reject});socket.send(JSON.stringify({id:n,method,params}));});
const ev=async expression=>{const r=await cdp('Runtime.evaluate',{expression,returnByValue:true,awaitPromise:true,userGesture:true});if(r.exceptionDetails)throw Error(r.exceptionDetails.exception?.description||r.exceptionDetails.text);return r.result.value;};
try{
 const env={...process.env};delete env.ELECTRON_RUN_AS_NODE;let profile,exe,args;
 if(electron){env.PRINTSHOPS_TEST_FOLDER=folder;profile=path.join(folder,'profile');exe=path.resolve('node_modules/electron/dist/electron.exe');args=['tests/helpers/desktop-tagged-resize-entry.mjs','--remote-debugging-port=0','--disable-renderer-backgrounding'];}
 else{try{await fetch('http://localhost:3000');}catch{server=spawn(process.execPath,['app.js'],{windowsHide:true,stdio:'ignore'});}await wait(async()=>{try{return(await fetch('http://localhost:3000')).ok;}catch{return false;}},'servidor');profile=path.join(folder,'profile-'+Date.now());await mkdir(profile);exe='C:/Program Files/Google/Chrome/Application/chrome.exe';args=['--headless=new','--disable-gpu','--no-first-run','--remote-debugging-port=0','--user-data-dir='+profile,'http://localhost:3000'];}
 app=spawn(exe,args,{env,windowsHide:true,stdio:'ignore'});let target;
 await wait(async()=>{try{const port=(await readFile(path.join(profile,'DevToolsActivePort'),'utf8')).split('\n')[0];target=(await(await fetch('http://127.0.0.1:'+port+'/json')).json()).find(t=>t.type==='page');return !!target;}catch{return false;}},'janela');
 socket=new WebSocket(target.webSocketDebuggerUrl);await new Promise(r=>socket.addEventListener('open',r,{once:true}));socket.addEventListener('message',e=>{const m=JSON.parse(e.data);if(m.id){m.error?pending.get(m.id)?.reject(Error(m.error.message)):pending.get(m.id)?.resolve(m.result);pending.delete(m.id);}if(m.method==='Runtime.exceptionThrown')errors.push(m.params.exceptionDetails);});
 await cdp('Runtime.enable');await cdp('Page.enable');await cdp('Page.bringToFront');await cdp('Emulation.setFocusEmulationEnabled',{enabled:true});await wait(()=>ev('!!document.querySelector("[data-tool=pdfs]")'),'interface');
 const result=await ev(`(${exercitarRedimensionamentoMarcado.toString()})(${JSON.stringify(fixtures)})`);assert.deepEqual(errors,[]);
 if(electron){let count=0;for(const name of await readdir(folder)){if(!name.startsWith('Synthetic-local')||!name.includes('_PNG_'))continue;const files=await readdir(path.join(folder,name));for(const file of files){assert.ok((await readFile(path.join(folder,name,file))).length>100);count++;}}assert.equal(count,fixtures.filter(f=>!f.blocked).length);result.realIPCExports=count;}
 await writeFile(path.join(folder,'results.json'),JSON.stringify({...result,originalHashPreserved:!!real,errors},null,2));console.log(`PASSOU ${electron?'Electron':'Chrome'}: redimensionamento vetorial, cancelamento, ActualText/fontes/símbolos/transparências, PNG/JPEG150/300/600 e alternativa protegida. ${result.reports.length} casos.`);
}catch(e){console.error(e);try{const r=await cdp('Page.captureScreenshot');await writeFile(path.join(folder,'failure.png'),Buffer.from(r.data,'base64'));}catch{}process.exitCode=1;}
finally{socket?.close();app?.kill();server?.kill();if(realFile)assert.equal(hash(await readFile(realFile)),originalHash);}
