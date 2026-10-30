// Testa o aplicativo NSIS instalado; não injeta mocks nem hooks no pacote.
import assert from 'node:assert/strict';
import {spawn,execFile} from 'node:child_process';
import {promisify} from 'node:util';
import {mkdtemp,readFile,writeFile,unlink,realpath} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import path from 'node:path';
import {setTimeout as pause} from 'node:timers/promises';
import {PDFDocument} from '../public/vendor/pdf-lib/pdf-lib.js';
import {testarOrientacoesBrowser} from './helpers/pdf-orientation-browser.mjs';

const exec=promisify(execFile), installation=(await readFile(path.join(tmpdir(),'4printshops-instalacao-teste.txt'),'utf8')).trim();
assert.ok((await realpath(installation)).toLowerCase().startsWith(path.join(await realpath(tmpdir()),'4printshops-instalado-').toLowerCase()),'Somente instalação temporária de teste');
const exe=path.join(installation,'4PrintShops.exe'), profile=path.join(process.env.APPDATA,'4PrintShops');
const folder=await mkdtemp(path.join(tmpdir(),'4printshops-installed-test-'));
let child,socket,id=0;const requests=new Map(),errors=[],diagnostics=[];
const wait=async(fn,label)=>{const end=Date.now()+30000;while(Date.now()<end){const r=await fn();if(r)return r;await pause(50);}throw new Error('Tempo esgotado: '+label);};
const cdp=(method,params={})=>new Promise((resolve,reject)=>{const n=++id,t=setTimeout(()=>{requests.delete(n);reject(new Error('Tempo CDP: '+method));},30000);requests.set(n,{resolve:r=>{clearTimeout(t);resolve(r);},reject:e=>{clearTimeout(t);reject(e);}});socket.send(JSON.stringify({id:n,method,params}));});
const evaluate=async(expression)=>{const r=await cdp('Runtime.evaluate',{expression,awaitPromise:true,returnByValue:true,userGesture:true});if(r.exceptionDetails)throw new Error(r.exceptionDetails.exception?.description||r.exceptionDetails.text);return r.result.value;};
const native=async(action,destination='')=>{
    const script=await readFile(new URL('./helpers/windows-save-dialog.ps1',import.meta.url),'utf8');
    const quote=v=>"'"+v.replaceAll("'","''")+"'";
    return exec('powershell.exe',['-NoProfile','-Command',`& {${script}} -DialogProcess ${child.pid} -Destination ${quote(destination)} -Action ${quote(action)}`],{windowsHide:true,timeout:25000});
};
async function launch(){
    await unlink(path.join(profile,'DevToolsActivePort')).catch(()=>{});
    const env={...process.env,PATH:path.join(process.env.SystemRoot,'System32')+';'+process.env.SystemRoot};delete env.ELECTRON_RUN_AS_NODE;
    // A janela GUI precisa estar visível para testar os diálogos reais por UIA.
    child=spawn(exe,['--remote-debugging-port=0'],{cwd:installation,env,windowsHide:false,stdio:['ignore','ignore','ignore']});
    const port=await wait(async()=>{try{return(await readFile(path.join(profile,'DevToolsActivePort'),'utf8')).split('\n')[0];}catch{return false;}},'Executável sem Node no PATH');
    const target=await wait(async()=>{try{return(await(await fetch(`http://127.0.0.1:${port}/json`)).json()).find(t=>t.type==='page');}catch{return false;}},'Janela instalada');
    socket=new WebSocket(target.webSocketDebuggerUrl);await new Promise(r=>socket.addEventListener('open',r,{once:true}));
    socket.addEventListener('message',e=>{const m=JSON.parse(e.data);if(m.id){if(m.error)requests.get(m.id)?.reject(new Error(m.error.message));else requests.get(m.id)?.resolve(m.result);requests.delete(m.id);}if(m.method==='Runtime.exceptionThrown')errors.push(m.params.exceptionDetails.text);if(m.method==='Log.entryAdded')diagnostics.push(m.params.entry.text);if(m.method==='Page.fileChooserOpened')diagnostics.push('Seletor interceptado: '+JSON.stringify(m.params));});
    await cdp('Runtime.enable');await cdp('Log.enable');await cdp('Page.enable');await wait(()=>evaluate(`!!document.querySelector('[data-tool="pdfs"]')`),'Início');
}
const ready=()=>wait(()=>evaluate(`!!document.querySelector('#pdf-arquivo')&&!document.querySelector('#pdf-arquivo').disabled`),'PDF pronto');
const click=async s=>{
    await evaluate(`document.querySelector(${JSON.stringify(s)}).click()`);
    if(s==='#pdf-salvar-copia'){
        const name=await evaluate(`document.querySelector('#app-exportacao-nome').value`);
        await evaluate(`document.querySelector('#app-exportacao-confirmar').click()`);
        const result=await native('save',path.join(folder,name));console.log('Salvar nativo:',result.stdout.trim());
    }
};
const screenshot=async name=>{const r=await cdp('Page.captureScreenshot');await writeFile(path.join(folder,name+'.png'),Buffer.from(r.data,'base64'));};
const download=async(name,total)=>{await wait(async()=>{try{return(await readFile(path.join(folder,name))).length;}catch{return false;}},'Gravação nativa');const p=await PDFDocument.load(await readFile(path.join(folder,name)));assert.equal(p.getPageCount(),total);return p;};
try{
    await launch();assert.equal(await evaluate(`typeof process`),'undefined');
    const second=spawn(exe,[],{cwd:installation,env:{...process.env,ELECTRON_RUN_AS_NODE:''},windowsHide:true,stdio:'ignore'});
    await wait(()=>second.exitCode!==null,'Segunda instância encerrada');assert.equal(second.exitCode,0);assert.equal(child.exitCode,null);
    await click('[data-tool="pdfs"]');await ready();
    await click('#app-menu-arquivo');await click('[data-command="verificar-atualizacoes"]');await native('updates');
    await testarOrientacoesBrowser({avaliar:evaluate,clicar:click,pronto:ready,esperarTela:(e,l)=>wait(()=>evaluate(e),l),imagem:screenshot,baixar:download});
    // Documento modificado e aba oculta: diálogo REAL do Windows ao fechar.
    const original=await PDFDocument.create();original.addPage([500,700]);const originalBytes=await original.save(),originalPath=path.join(folder,'fechar.pdf');await writeFile(originalPath,originalBytes);
    await cdp('Page.bringToFront');await cdp('Emulation.setFocusEmulationEnabled',{enabled:true});
    const importPoint=await evaluate(`(()=>{const r=document.querySelector('.pdf-importar').getBoundingClientRect();return{x:r.x+r.width/2,y:r.y+r.height/2}})()`);
    console.log('Importação:',await evaluate(`JSON.stringify({focus:document.hasFocus(),alvo:document.elementFromPoint(${importPoint.x},${importPoint.y})?.className,disabled:document.querySelector('#pdf-arquivo').disabled})`));
    await cdp('Input.dispatchMouseEvent',{type:'mousePressed',...importPoint,button:'left',clickCount:1});await cdp('Input.dispatchMouseEvent',{type:'mouseReleased',...importPoint,button:'left',clickCount:1});await native('open',originalPath);
    await wait(()=>evaluate(`document.querySelector('#pdf-nome').textContent==='fechar.pdf'`),'Importação pelo diálogo real');await ready();await click('#pdf-girar-direita');await ready();
    await evaluate(`document.querySelector('#pdf-salvar-copia').click();document.querySelector('#app-exportacao-confirmar').click()`);await native('cancel-save');await ready();
    assert.equal(await evaluate(`document.querySelector('#pdf-alteracoes-pendentes').hidden`),false,'Cancelar Salvar nativo conserva a pendência');await click('#app-exportacao-cancelar');
    assert.deepEqual(await readFile(originalPath),Buffer.from(originalBytes),'PDF original preservado');await click('[data-tool="inicio"]');
    const closing=evaluate('window.close()').catch(()=>{});await native('continue');await closing;await pause(200);assert.equal(child.exitCode,null);
    assert.equal(await evaluate(`document.querySelector('#app-shell').dataset.activeTool`),'inicio');
    await evaluate(`localStorage.setItem('4printshops.desktop.installed-test','persistiu')`);await screenshot('instalado-inicio');
    const exit=evaluate('window.close()').catch(()=>{});await native('discard');await wait(()=>child.exitCode!==null,'Fechamento confirmado');await exit;
    await launch();assert.equal(await evaluate(`localStorage.getItem('4printshops.desktop.installed-test')`),'persistiu');
    await evaluate('window.close()');await wait(()=>child.exitCode!==null,'Reinício encerrado');assert.deepEqual(errors,[]);
    console.log('PASSOU: NSIS instalado, execução sem Node no PATH, instância única, PDFs/QR, giro/reorganização/exclusão/união/resize, exportação com diálogo REAL do Windows, proteção nativa e preferências após reiniciar.');
    console.log('Instalação: '+installation);console.log('Artefatos: '+folder);
}catch(error){console.error('Artefatos: '+folder);console.error(error.stderr || error.message);console.error('Diagnóstico:',diagnostics);try{console.error('Estado:',await evaluate(`({visibility:document.visibilityState,exportacao:document.querySelector('#app-exportacao-status')?.textContent,status:document.querySelector('#pdf-status')?.textContent})`));await screenshot('falha');}catch{}process.exitCode=1;}
finally{socket?.close();if(child&&child.exitCode===null)child.kill();}
