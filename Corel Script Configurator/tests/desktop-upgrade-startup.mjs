// Opt-in: open only the isolated upgrade identity, or legacy modules with an
// explicit laboratory profile. Never launch a legacy executable/uninstaller.
import assert from 'node:assert/strict';
import {spawn} from 'node:child_process';
import {readFile,writeFile,mkdtemp,mkdir} from 'node:fs/promises';
import {createRequire} from 'node:module';
import {tmpdir} from 'node:os';
import {pathToFileURL,fileURLToPath} from 'node:url';
import path from 'node:path';
const require=createRequire(import.meta.url),asar=require('@electron/asar');
const project=fileURLToPath(new URL('../',import.meta.url));
const root=path.resolve(process.argv[2]);
const legacy=process.argv.includes('--legacy-modules');
const archive=path.join(root,'resources/app.asar');
const pkg=JSON.parse(asar.extractFile(archive,'package.json'));
const evidence=await mkdtemp(path.join(tmpdir(),'4ps-startup-'));
const electron=path.join(project,'node_modules/electron/dist/electron.exe');
const env={...process.env};delete env.ELECTRON_RUN_AS_NODE;
if(!legacy){
 const id=JSON.parse(asar.extractFile(archive,path.join('build/distribution/identity.json')));
 assert.equal(id.lab,true);assert.equal(id.appId,'br.com.4printshops.upgradelab.transaction36');assert.equal(id.exe,'4PrintShops Upgrade Lab.exe');
 const probe=spawn(electron,[path.join(project,'tests/helpers/upgrade-profile-probe.mjs'),root],{env,windowsHide:true,stdio:['ignore','pipe','pipe']});
 let log='';probe.stdout.on('data',d=>log+=d);probe.stderr.on('data',d=>log+=d);
 const code=await new Promise((r,j)=>{probe.once('error',j);probe.once('close',r);});
 await writeFile(path.join(evidence,'profile-probe.log'),log);assert.equal(code,0,'Unsafe lab profile: '+log);
 if(process.argv.includes('--probe-only')){console.log(JSON.stringify({passed:true,evidence,profileProbe:true}));process.exit(0);}
}else {assert.equal(pkg.version,'1.0.0-experimental.34');assert.ok(root.toLowerCase().startsWith(path.join(tmpdir(),'4ps-full-upgrade-').toLowerCase()),'Legacy modules must belong to this disposable laboratory');}
const profile=path.resolve(root,'../lab-profile');await mkdir(profile,{recursive:true});
let exe,args;
if(legacy){
 // Import unchanged .34 modules, but force profile, taskbar identity and name
 // before their first await. The old inert executable is never run.
 const entry=path.join(evidence,'legacy-entry.mjs');
 await writeFile(entry,`import {createRequire} from 'node:module';\nimport {startDesktop} from ${JSON.stringify(pathToFileURL(path.join(root,'resources/app.asar/desktop/app.mjs')).href)};\nconst e=createRequire(${JSON.stringify(path.join(project,'package.json'))})('electron');\ne.app.setPath('userData',${JSON.stringify(profile)});\nObject.defineProperty(process,'resourcesPath',{value:${JSON.stringify(path.join(root,'resources'))},configurable:true});\nconst app=new Proxy(e.app,{get:(target,key)=>key==='isPackaged'?true:key==='getVersion'?()=>${JSON.stringify(pkg.version)}:key==='setAppUserModelId'?()=>target.setAppUserModelId('br.com.4printshops.upgradelab.transaction36'):key==='setName'?()=>target.setName('4PrintShops Upgrade Lab'):typeof target[key]==='function'?target[key].bind(target):target[key]});\nstartDesktop({...e,app},{profile:${JSON.stringify(profile)}}).catch(error=>{console.error(error);e.app.exit(1)});\n`);
 exe=electron;args=[entry,'--remote-debugging-port=0'];
}else{exe=path.join(root,'4PrintShops Upgrade Lab.exe');args=['--remote-debugging-port=0'];}
let socket,serial=0,output='';const pending=new Map(),errors=[];
const child=spawn(exe,args,{cwd:project,env,windowsHide:true,stdio:['ignore','pipe','pipe']});
child.stdout.on('data',d=>output+=d);child.stderr.on('data',d=>output+=d);
const pause=ms=>new Promise(r=>setTimeout(r,ms));
async function wait(fn,label){const end=Date.now()+60000;while(Date.now()<end){if(child.exitCode!==null)throw Error('Startup ended '+child.exitCode+' '+output);const v=await fn();if(v)return v;await pause(80);}throw Error('Timeout: '+label);}
const cdp=(method,params={})=>new Promise((resolve,reject)=>{const id=++serial,timer=setTimeout(()=>{pending.delete(id);reject(Error(method));},30000);pending.set(id,{resolve:v=>{clearTimeout(timer);resolve(v);},reject:e=>{clearTimeout(timer);reject(e);}});socket.send(JSON.stringify({id,method,params}));});
const evaluate=async expression=>{const r=await cdp('Runtime.evaluate',{expression,awaitPromise:true,returnByValue:true,userGesture:true});if(r.exceptionDetails)throw Error(r.exceptionDetails.exception?.description||r.exceptionDetails.text);return r.result.value;};
try{
 const target=await wait(async()=>{try{const port=(await readFile(path.join(profile,'DevToolsActivePort'),'utf8')).split('\n')[0];return(await(await fetch('http://127.0.0.1:'+port+'/json')).json()).find(t=>t.type==='page'&&t.url.startsWith('app://4printshops'));}catch{return false;}},'application page');
 socket=new WebSocket(target.webSocketDebuggerUrl);await new Promise(r=>socket.addEventListener('open',r,{once:true}));
 socket.addEventListener('message',event=>{const m=JSON.parse(event.data);if(m.id){const p=pending.get(m.id);m.error?p?.reject(Error(m.error.message)):p?.resolve(m.result);pending.delete(m.id);}if(m.method==='Runtime.exceptionThrown')errors.push(m.params.exceptionDetails.text);});
 await cdp('Runtime.enable');await wait(()=>evaluate('!!document.querySelector("[data-tool=pdfs]")'),'home');
 const storageBefore=await evaluate('Object.fromEntries(Object.keys(localStorage).sort().map(k=>[k,localStorage.getItem(k)]))');
 await evaluate('localStorage.setItem("upgrade36.synthetic.preference","preserved");document.querySelector("[data-tool=pdfs]").click()');
 await wait(()=>evaluate('!!document.querySelector("#pdf-arquivo")'),'editor');
 await evaluate(`(async()=>{const {PDFDocument}=await import('/vendor/pdf-lib/pdf-lib.js');const p=await PDFDocument.create();p.addPage([300,400]).drawText('SYNTHETIC UPGRADE CHECK');const file=new File([await p.save()],'synthetic-upgrade.pdf',{type:'application/pdf'});document.querySelector('#pdf-arquivo').dispatchEvent(new CustomEvent('change',{bubbles:true,detail:{files:[file]}}));})()`);
 await wait(()=>evaluate('(document.querySelector("#pdf-uniao-janela")?.open&&!document.querySelector("#pdf-uniao-gerar").disabled)||(document.querySelector("#pdf-total").textContent==="1"&&!!document.querySelector("#pdf-visualizacao canvas"))'),'imported synthetic PDF');
 if(await evaluate('document.querySelector("#pdf-uniao-janela")?.open'))await evaluate('document.querySelector("#pdf-uniao-gerar").click()');
 await wait(()=>evaluate('document.querySelector("#pdf-total").textContent==="1"&&!!document.querySelector("#pdf-visualizacao canvas")'),'rendered PDF');
 await evaluate('document.querySelector("#pdf-girar-direita").click()');
 await wait(()=>evaluate('!document.querySelector("#pdf-arquivo").disabled&&!document.querySelector("#pdf-alteracoes-pendentes").hidden'),'editing available');
 const dimensionsAfterRotation=await evaluate('document.querySelector("#pdf-tamanho").textContent');assert.match(dimensionsAfterRotation,/141,11.*105,83/);
 const shot=await cdp('Page.captureScreenshot',{format:'png'});await writeFile(path.join(evidence,'pdf.png'),Buffer.from(shot.data,'base64'));
 assert.deepEqual(errors,[]);
 await evaluate('document.querySelector("[data-close-tool=pdfs]").click()');
 await wait(()=>evaluate('document.querySelector("#pdf-confirmar-descarte")?.open'),'close protection');
 await evaluate('document.querySelector("#pdf-descartar-sair").click()');
 await wait(()=>evaluate('!document.querySelector("#workspace-pdfs")'),'closed PDF tab');
 await cdp('Runtime.evaluate',{expression:'setTimeout(()=>window.close(),0)'});
 const ended=await Promise.race([new Promise(r=>{if(child.exitCode!==null)r(child.exitCode);else child.once('close',r);}),pause(20000).then(()=>{throw Error('Normal close did not complete');})]);assert.equal(ended,0);
 const result={passed:true,root,legacyModules:legacy,version:pkg.version,profile,storageBefore,syntheticPreference:'preserved',pdfRendered:true,editingAvailable:true,dimensionsAfterRotation,closeProtection:true,normalExit:true,evidence};
 await writeFile(path.join(evidence,'result.json'),JSON.stringify(result,null,2));console.log(JSON.stringify(result));
}finally{socket?.close();await writeFile(path.join(evidence,'process.log'),output);if(child.exitCode===null)child.kill();}
