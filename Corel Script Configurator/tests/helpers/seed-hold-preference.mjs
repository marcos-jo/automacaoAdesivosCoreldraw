// Usa a experimental.9 preservada para provar a leitura da preferência pela .10.
// Abre somente Início, sem documento e sem comandos de impressão.
import {spawn} from 'node:child_process';
import {readFile,writeFile,unlink} from 'node:fs/promises';
import path from 'node:path';
import {setTimeout as pause} from 'node:timers/promises';
import {createRequire} from 'node:module';
const require=createRequire(import.meta.url),asar=require('@electron/asar');
const folder=path.resolve('dist-desktop/hold-test/win-unpacked');
const metadata=JSON.parse(asar.extractFile(path.join(folder,'resources/app.asar'),'package.json').toString('utf8'));
if(metadata.version!=='1.0.0-experimental.9')throw Error('Base .9 preservada indisponível.');
const profile=path.join(process.env.APPDATA,'4PrintShops Hold Test');
await unlink(path.join(profile,'DevToolsActivePort')).catch(()=>{});
const env={...process.env};delete env.ELECTRON_RUN_AS_NODE;
const app=spawn(path.join(folder,'4PrintShops Hold Test.exe'),['--remote-debugging-port=0'],{env,windowsHide:true,stdio:'ignore'});
let socket,id=0;const calls=new Map();
try {
    let port;const end=Date.now()+30000;
    while(!port){try{port=(await readFile(path.join(profile,'DevToolsActivePort'),'utf8')).split('\n')[0];}catch{if(Date.now()>end)throw Error('Inicialização .9 sem resposta.');await pause(100);}}
    const target=(await (await fetch(`http://127.0.0.1:${port}/json`)).json()).find(item=>item.type==='page');
    socket=new WebSocket(target.webSocketDebuggerUrl);await new Promise(resolve=>socket.addEventListener('open',resolve,{once:true}));
    socket.addEventListener('message',event=>{const message=JSON.parse(event.data);if(message.id){calls.get(message.id)?.(message);calls.delete(message.id);}});
    const evaluate=expression=>new Promise((resolve,reject)=>{const next=++id;calls.set(next,message=>{if(message.error||message.result?.exceptionDetails)reject(Error(message.error?.message||message.result.exceptionDetails.text));else resolve(message.result.result.value);});socket.send(JSON.stringify({id:next,method:'Runtime.evaluate',params:{expression,returnByValue:true}}));});
    const readyEnd=Date.now()+30000;
    while(!await evaluate("location.origin==='app://4printshops'&&!!document.querySelector('#app-shell')")){if(Date.now()>readyEnd)throw Error('Interface .9 indisponível.');await pause(100);}
    const original=await evaluate("localStorage.getItem('4printshops.print-config-check')");
    await evaluate("localStorage.setItem('4printshops.print-config-check','preservado')");
    await pause(1500);
    await writeFile('.desktop-test/hold-native/preference-seed.json',JSON.stringify({version:metadata.version,profile,originalValue:original,key:'4printshops.print-config-check'},null,2));
    await evaluate('setTimeout(()=>window.close(),100); true');
    const closeEnd=Date.now()+15000;
    while(app.exitCode===null){if(Date.now()>closeEnd)throw Error('Janela .9 não fechou normalmente.');await pause(100);}
    console.log('Preferência gravada pela .9 real e aplicativo fechado normalmente.');
} finally {socket?.close();if(app.exitCode===null)app.kill();}
