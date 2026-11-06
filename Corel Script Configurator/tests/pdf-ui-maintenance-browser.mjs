// Chrome/Electron em perfis isolados. Nunca instala nem envia trabalhos de impressão.
import assert from 'node:assert/strict';import {spawn} from 'node:child_process';
import {mkdir,readFile,writeFile} from 'node:fs/promises';import path from 'node:path';
import {testarAvisosMontagem} from './helpers/pdf-imposition-warnings-browser.mjs';
import {diagnosticarZoom,testarMatrizZoom} from './helpers/pdf-zoom-anchor-browser.mjs';
const electron=process.env.UX27_ELECTRON==='1',folder=path.resolve(process.env.UX27_RESULTS||`.desktop-test/ux27/${electron?'electron':'chrome'}`);await mkdir(folder,{recursive:true});
let app,server,socket,id=0;const pending=new Map(),errors=[],sleep=ms=>new Promise(r=>setTimeout(r,ms));
const wait=async(fn,label)=>{for(let i=0;i<600;i++){if(await fn())return;await sleep(100);}throw Error('Espera: '+label);};
const cdp=(method,params={})=>new Promise((resolve,reject)=>{const n=++id,t=setTimeout(()=>{pending.delete(n);reject(Error(method));},60000);pending.set(n,{resolve:r=>{clearTimeout(t);resolve(r);},reject});socket.send(JSON.stringify({id:n,method,params}));});
const ev=async expression=>{const r=await cdp('Runtime.evaluate',{expression,returnByValue:true,awaitPromise:true,userGesture:true});if(r.exceptionDetails)throw Error(r.exceptionDetails.exception?.description||r.exceptionDetails.text);return r.result.value;};
const screenshot=async name=>{const r=await cdp('Page.captureScreenshot');await writeFile(path.join(folder,name+'.png'),Buffer.from(r.data,'base64'));};
try{
 const env={...process.env};delete env.ELECTRON_RUN_AS_NODE;let profile,exe,args;
 if(electron){env.PRINTSHOPS_TEST_FOLDER=folder;await writeFile(path.join(folder,'command.json'),JSON.stringify({serial:0}));profile=path.join(folder,'profile');exe=path.resolve('node_modules/electron/dist/electron.exe');args=['tests/helpers/desktop-entry.mjs','--remote-debugging-port=0','--disable-renderer-backgrounding'];}
 else{try{await fetch('http://localhost:3000');}catch{server=spawn(process.execPath,['app.js'],{windowsHide:true,stdio:'ignore'});}await wait(async()=>{try{return(await fetch('http://localhost:3000')).ok;}catch{return false;}},'servidor');profile=path.join(folder,'profile-'+Date.now());await mkdir(profile);exe=process.env.BROWSER_PATH||'C:/Program Files/Google/Chrome/Application/chrome.exe';args=['--headless=new','--disable-gpu','--no-first-run','--remote-debugging-port=0','--user-data-dir='+profile,'http://localhost:3000'];}
 app=spawn(exe,args,{env,windowsHide:true,stdio:'ignore'});let port,target;
 await wait(async()=>{try{port=(await readFile(path.join(profile,'DevToolsActivePort'),'utf8')).split('\n')[0];target=(await(await fetch('http://127.0.0.1:'+port+'/json')).json()).find(t=>t.type==='page');return !!target;}catch{return false;}},'janela');
 socket=new WebSocket(target.webSocketDebuggerUrl);await new Promise(r=>socket.addEventListener('open',r,{once:true}));socket.addEventListener('message',e=>{const m=JSON.parse(e.data);if(m.id){m.error?pending.get(m.id)?.reject(Error(m.error.message)):pending.get(m.id)?.resolve(m.result);pending.delete(m.id);}if(m.method==='Runtime.exceptionThrown')errors.push(m.params.exceptionDetails);});
 await cdp('Runtime.enable');await cdp('Page.enable');await cdp('Page.bringToFront');await cdp('Emulation.setFocusEmulationEnabled',{enabled:true});await cdp('Emulation.setDeviceMetricsOverride',{width:1366,height:768,deviceScaleFactor:1,mobile:false});await wait(()=>ev('!!document.querySelector("[data-tool=pdfs]")'),'interface');
 if(process.env.UX27_PHASE!=='zoom')await testarAvisosMontagem({ev,cdp,wait,folder,screenshot});
 if(process.env.UX27_PHASE==='zoom'){await diagnosticarZoom({ev,cdp,wait,folder,screenshot});if(process.env.UX27_ZOOM_DIAG!=='1')await testarMatrizZoom({ev,cdp,wait,folder,screenshot});}
 assert.deepEqual(errors,[]);console.log('Concluído UX27 '+(process.env.UX27_PHASE||'avisos')+' em '+(electron?'Electron':'Chrome')+'. Sem impressão.');
}catch(e){console.error(e);console.error(errors);try{await screenshot('failure');}catch{}process.exitCode=1;}
finally{socket?.close();app?.kill();server?.kill();}
