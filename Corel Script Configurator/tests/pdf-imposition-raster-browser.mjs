import assert from 'node:assert/strict';import {spawn} from 'node:child_process';
import {mkdir,readFile,writeFile} from 'node:fs/promises';import path from 'node:path';import {createHash} from 'node:crypto';
import {exercitarRasterMontagem} from './helpers/pdf-imposition-raster-browser.mjs';
const electron=process.env.RASTER_ELECTRON==='1',folder=path.resolve(process.env.RASTER_RESULTS||`.desktop-test/raster28/${electron?'electron':'chrome'}`);await mkdir(folder,{recursive:true});
const realFile=process.env.RASTER_REAL_PATH,real=realFile?await readFile(realFile):null,hash=b=>createHash('sha256').update(b).digest('hex');
const originalHash=real&&hash(real),icc=await readFile('C:/Windows/System32/spool/drivers/color/sRGB Color Space Profile.icm');
let app,server,socket,id=0;const pending=new Map(),errors=[],sleep=ms=>new Promise(r=>setTimeout(r,ms));
const wait=async(fn,label)=>{for(let i=0;i<600;i++){if(await fn())return;await sleep(100);}throw Error('Espera: '+label);};
const cdp=(method,params={})=>new Promise((resolve,reject)=>{const n=++id,t=setTimeout(()=>{pending.delete(n);reject(Error(method));},180000);pending.set(n,{resolve:r=>{clearTimeout(t);resolve(r);},reject});socket.send(JSON.stringify({id:n,method,params}));});
const ev=async expression=>{const r=await cdp('Runtime.evaluate',{expression,returnByValue:true,awaitPromise:true,userGesture:true});if(r.exceptionDetails)throw Error(r.exceptionDetails.exception?.description||r.exceptionDetails.text);return r.result.value;};
try{
 const env={...process.env};delete env.ELECTRON_RUN_AS_NODE;let profile,exe,args;
 if(electron){env.PRINTSHOPS_TEST_FOLDER=folder;await writeFile(path.join(folder,'command.json'),JSON.stringify({serial:0,save:{canceled:false,filePath:path.join(folder,'ui-export.pdf')}}));profile=path.join(folder,'profile');exe=path.resolve('node_modules/electron/dist/electron.exe');args=['tests/helpers/desktop-entry.mjs','--remote-debugging-port=0','--disable-renderer-backgrounding'];}
 else{try{await fetch('http://localhost:3000');}catch{server=spawn(process.execPath,['app.js'],{windowsHide:true,stdio:'ignore'});}await wait(async()=>{try{return(await fetch('http://localhost:3000')).ok;}catch{return false;}},'servidor');profile=path.join(folder,'profile-'+Date.now());await mkdir(profile);exe='C:/Program Files/Google/Chrome/Application/chrome.exe';args=['--headless=new','--disable-gpu','--no-first-run','--remote-debugging-port=0','--user-data-dir='+profile,'http://localhost:3000'];}
 app=spawn(exe,args,{env,windowsHide:true,stdio:'ignore'});let target;
 await wait(async()=>{try{const port=(await readFile(path.join(profile,'DevToolsActivePort'),'utf8')).split('\n')[0];target=(await(await fetch('http://127.0.0.1:'+port+'/json')).json()).find(t=>t.type==='page');return !!target;}catch{return false;}},'janela');
 socket=new WebSocket(target.webSocketDebuggerUrl);await new Promise(r=>socket.addEventListener('open',r,{once:true}));socket.addEventListener('message',e=>{const m=JSON.parse(e.data);if(m.id){m.error?pending.get(m.id)?.reject(Error(m.error.message)):pending.get(m.id)?.resolve(m.result);pending.delete(m.id);}if(m.method==='Runtime.exceptionThrown')errors.push(m.params.exceptionDetails);});
 await cdp('Runtime.enable');await cdp('Page.enable');await cdp('Page.bringToFront');await cdp('Emulation.setFocusEmulationEnabled',{enabled:true});await cdp('Emulation.setDeviceMetricsOverride',{width:1366,height:768,deviceScaleFactor:1,mobile:false});await wait(()=>ev('!!document.querySelector("[data-tool=pdfs]")'),'interface');
 if(!electron){await mkdir(path.join(folder,'downloads'),{recursive:true});await cdp('Browser.setDownloadBehavior',{behavior:'allow',downloadPath:path.join(folder,'downloads')});}
 const result=await ev(`(${exercitarRasterMontagem.toString()})(${JSON.stringify(real?.toString('base64')||null)},${JSON.stringify(icc.toString('base64'))})`);
 for(const [i,r]of result.records.entries()){for(const key of ['prepared','output','nup','preview'])if(r[key]){await writeFile(path.join(folder,`${i}-${r.tool}-${r.dpi}-${key}.${key==='preview'?'png':'pdf'}`),Buffer.from(r[key],'base64'));delete r[key];}}
 const last=result.records.length-1;
 const exports=electron?[[path.join(folder,'ui-export.pdf'),path.join(folder,`${last}-livreto-600-output.pdf`)]]:[[path.join(folder,'downloads','Livreto_4PrintShops.pdf'),path.join(folder,`${last}-livreto-600-output.pdf`)],[path.join(folder,'downloads','Montagem_4PrintShops.pdf'),path.join(folder,`${last-1}-montagem-600-nup.pdf`)]];
 for(const [saved,expected]of exports)await wait(async()=>{try{return (await readFile(saved)).equals(await readFile(expected));}catch{return false;}},'exportação idêntica aos bytes validados');
 result.exportsMatchValidatedBytes=true;
 assert.deepEqual(errors,[]);if(realFile)assert.equal(hash(await readFile(realFile)),originalHash);
 await writeFile(path.join(folder,'results.json'),JSON.stringify({...result,originalHashPreserved:!!real,errors},null,2));
 console.log(`PASSOU raster ${electron?'Electron':'Chrome'}: ${result.records.length} preparações UI, original/pendências/configurações, N-up/Step & Repeat/livreto, falha/retry e cancelamento. Sem impressão.`);
}catch(e){console.error(e);console.error(errors);try{const r=await cdp('Page.captureScreenshot');await writeFile(path.join(folder,'failure.png'),Buffer.from(r.data,'base64'));}catch{}process.exitCode=1;}
finally{socket?.close();app?.kill();server?.kill();if(realFile)assert.equal(hash(await readFile(realFile)),originalHash);}
