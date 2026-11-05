// Prova local antes/depois da atualização instalada. Não toca PDFs nem impressoras.
import assert from 'node:assert/strict';
import {spawn,execFileSync} from 'node:child_process';
import {readFile,writeFile,mkdir,unlink} from 'node:fs/promises';
import path from 'node:path';
import {setTimeout as pause} from 'node:timers/promises';
const phase=process.argv[2];assert.ok(['before','after'].includes(phase));
const folder=path.resolve('.desktop-test/booklet24'),file=path.join(folder,'preferences-before.json');await mkdir(folder,{recursive:true});
const ps=code=>execFileSync('powershell.exe',['-NoProfile','-NonInteractive','-Command',code],{encoding:'utf8',windowsHide:true}).trim();
assert.equal(ps("@(Get-Process '4PrintShops Hold Test' -ErrorAction SilentlyContinue).Count"),'0','Feche a Hold Test antes do ensaio.');
const registry='HKCU:/Software/Microsoft/Windows/CurrentVersion/Uninstall/90ff2ff0-f013-571c-b109-7edb265dacc8',info=JSON.parse(ps(`Get-ItemProperty '${registry}'|Select-Object DisplayVersion,UninstallString|ConvertTo-Json -Compress`));
const exe=path.join(path.dirname(info.UninstallString.match(/^"([^"]+)"/)[1]),'4PrintShops Hold Test.exe'),profile=path.join(process.env.APPDATA,'4PrintShops Hold Test');
const saved=phase==='after'?JSON.parse(await readFile(file,'utf8')):null;
if(phase==='before')assert.equal(info.DisplayVersion,'1.0.0-experimental.23');else{assert.equal(info.DisplayVersion,'1.0.0-experimental.24');assert.equal(exe,saved.exe);assert.equal(profile,saved.profile);}
await unlink(path.join(profile,'DevToolsActivePort')).catch(()=>{});const env={...process.env};delete env.ELECTRON_RUN_AS_NODE;
const app=spawn(exe,['--remote-debugging-port=0'],{env,windowsHide:true,stdio:'ignore'});let socket;
try{
    let port,targets;for(let i=0;i<300;i++){try{port=(await readFile(path.join(profile,'DevToolsActivePort'),'utf8')).split('\n')[0];targets=await(await fetch('http://127.0.0.1:'+port+'/json')).json();if(targets.some(t=>t.url.startsWith('app://4printshops')))break;}catch{}await pause(100);}
    socket=new WebSocket(targets.find(t=>t.url.startsWith('app://4printshops')).webSocketDebuggerUrl);await new Promise(r=>socket.addEventListener('open',r,{once:true}));let id=0;const pending=new Map();socket.addEventListener('message',e=>{const m=JSON.parse(e.data);if(m.id){pending.get(m.id)?.(m);pending.delete(m.id);}});
    const evaluate=expression=>new Promise((resolve,reject)=>{const n=++id,t=setTimeout(()=>reject(Error('Consulta do perfil expirou.')),10000);pending.set(n,m=>{clearTimeout(t);if(m.result?.exceptionDetails)reject(Error(m.result.exceptionDetails.text));else resolve(m.result?.result?.value);});socket.send(JSON.stringify({id:n,method:'Runtime.evaluate',params:{expression,returnByValue:true,awaitPromise:true}}));});
    for(let i=0;i<200;i++){if(await evaluate("!!document.querySelector('[data-tool=pdfs]')"))break;await pause(100);}
    const actual=await evaluate('({origin:location.origin,values:Object.fromEntries(Object.entries(localStorage).filter(([k])=>k!=="4printshops.test.booklet24"))})');assert.equal(actual.origin,'app://4printshops');
    if(phase==='before'){await evaluate('localStorage.setItem("4printshops.test.booklet24","preservado-24")');await writeFile(file,JSON.stringify({...actual,exe,profile,version:info.DisplayVersion},null,2));}
    else{assert.equal(await evaluate('localStorage.getItem("4printshops.test.booklet24")'),'preservado-24');for(const [k,v]of Object.entries(saved.values))assert.equal(actual.values[k],v,`Preferência alterada: ${k}`);await evaluate('localStorage.removeItem("4printshops.test.booklet24")');await writeFile(path.join(folder,'preferences-after.json'),JSON.stringify({version:info.DisplayVersion,keysCompared:Object.keys(saved.values).length,sameDirectory:true,sameProfile:true,sameOrigin:true,markerPreserved:true},null,2));}
    console.log('PASSOU perfil '+phase+': '+info.DisplayVersion+'; '+Object.keys(actual.values).length+' preferências, origem '+actual.origin+'.');
    socket.send(JSON.stringify({id:++id,method:'Page.close'}));for(let i=0;i<100&&app.exitCode===null;i++)await pause(100);
}finally{socket?.close();if(app.exitCode===null)app.kill();}
