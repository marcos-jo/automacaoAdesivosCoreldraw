// NSIS de laboratório + Caddy portátil reais. Nunca executa o instalador oferecido.
import assert from 'node:assert/strict';
import {spawn,execFile} from 'node:child_process';
import {promisify} from 'node:util';
import {mkdtemp,mkdir,readFile,writeFile,unlink,realpath,readdir,open} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import path from 'node:path';
import {setTimeout as pause} from 'node:timers/promises';
import {sign} from 'node:crypto';
import {loadPublishingKey} from '../scripts/assisted/keys.mjs';
import {PDFDocument} from '../public/vendor/pdf-lib/pdf-lib.js';

const exec=promisify(execFile),lab=JSON.parse(await readFile('dist-desktop/assisted-lab/lab-local.json','utf8'));
const folder=await mkdtemp(path.join(tmpdir(),'4printshops-assisted-installed-')),installation=path.join(folder,'installed');
const profile=path.join(process.env.APPDATA,'4PrintShops Assisted Lab'),owner=path.join(profile,'4printshops-test-owner.txt');
const base='1.0.0-experimental.5',target='1.0.0-experimental.6';
const setup=path.resolve(`dist-desktop/assisted-lab/${base}/4PrintShops-Assisted-Lab-${base}-Windows-x64-Setup.exe`);
const exe=path.join(installation,'4PrintShops Assisted Lab.exe'),manifestFile=path.join(lab.repository,'channels','internal-lab.json');
const manifest=await readFile(manifestFile),envelope=JSON.parse(manifest),data=JSON.parse(Buffer.from(envelope.payload,'base64url'));
const installerFile=path.join(lab.repository,data.installer.path),privateKey=await loadPublishingKey(lab.key);
const caddy=process.env.FOURPRINT_CADDY||path.join(tmpdir(),'4printshops-caddy-2.11.7','caddy.exe');
let server,child,socket,serial=0;const pending=new Map(),exceptions=[];
const wait=async(fn,label,timeout=40000)=>{const end=Date.now()+timeout;while(Date.now()<end){if(await fn())return;await pause(70);}throw new Error('Tempo esgotado: '+label);};
const cdp=(method,params={})=>new Promise((resolve,reject)=>{const id=++serial,timer=setTimeout(()=>{pending.delete(id);reject(new Error('CDP: '+method));},10000);pending.set(id,{resolve:v=>{clearTimeout(timer);resolve(v);},reject});socket.send(JSON.stringify({id,method,params}));});
const evaluate=async expression=>{const r=await cdp('Runtime.evaluate',{expression,awaitPromise:true,returnByValue:true,userGesture:true});if(r.exceptionDetails)throw new Error(r.exceptionDetails.exception?.description||r.exceptionDetails.text);return r.result.value;};
const click=selector=>evaluate(`document.querySelector(${JSON.stringify(selector)}).click()`);
const state=()=>evaluate('window.desktop4PrintShops.estadoAssistida()');
const invoke=action=>evaluate(`void window.desktop4PrintShops.atualizacaoAssistida(${JSON.stringify(action)})`);
const native=async(action,destination='')=>{
    const script=await readFile(new URL('./helpers/windows-save-dialog.ps1',import.meta.url),'utf8'),quote=v=>"'"+v.replaceAll("'","''")+"'";
    await exec('powershell.exe',['-NoProfile','-Command',`& {${script}} -DialogProcess ${child.pid} -Action ${quote(action)} -Destination ${quote(destination)}`],{windowsHide:true,timeout:25000});
};
const screenshot=async name=>{const shot=await cdp('Page.captureScreenshot');await writeFile(path.join(folder,name+'.png'),Buffer.from(shot.data,'base64'));};
async function startServer(){
    const env={...process.env,PATH:path.join(process.env.SystemRoot,'System32')+';'+process.env.SystemRoot};
    server=spawn(caddy,['run','--config',path.resolve('dist-desktop/assisted-lab/Caddyfile'),'--adapter','caddyfile'],{env,windowsHide:true,stdio:'ignore'});
    await wait(async()=>{try{return(await fetch('http://127.0.0.1:18084/channels/internal-lab.json')).ok;}catch{return false;}},'Caddy loopback');
    assert.equal((await fetch('http://127.0.0.1:18084/')).status,404);
    assert.equal((await fetch('http://127.0.0.1:18084/.publish.lock')).status,404);
    assert.equal((await fetch('http://127.0.0.1:18084/channels/internal-lab.json',{method:'POST'})).status,404);
}
async function stopServer(){if(!server)return;const current=server;server=null;current.kill();await new Promise(r=>current.once('exit',r));}
async function launch(){
    await unlink(path.join(profile,'DevToolsActivePort')).catch(()=>{});
    const env={...process.env,PATH:path.join(process.env.SystemRoot,'System32')+';'+process.env.SystemRoot};delete env.ELECTRON_RUN_AS_NODE;
    child=spawn(exe,['--remote-debugging-port=0'],{env,cwd:installation,windowsHide:true,stdio:'ignore'});
    let port,targetInfo;
    await wait(async()=>{try{port=(await readFile(path.join(profile,'DevToolsActivePort'),'utf8')).split('\n')[0];return port;}catch{return false;}},'App instalado');
    await wait(async()=>{try{targetInfo=(await(await fetch(`http://127.0.0.1:${port}/json`)).json()).find(t=>t.type==='page');return targetInfo;}catch{return false;}},'Janela instalada');
    socket=new WebSocket(targetInfo.webSocketDebuggerUrl);await new Promise(r=>socket.addEventListener('open',r,{once:true}));
    socket.addEventListener('message',event=>{const value=JSON.parse(event.data);if(value.id){const job=pending.get(value.id);pending.delete(value.id);value.error?job?.reject(new Error(value.error.message)):job?.resolve(value.result);}if(value.method==='Runtime.exceptionThrown')exceptions.push(value.params.exceptionDetails.text);});
    await cdp('Runtime.enable');await cdp('Page.enable');await wait(()=>evaluate("!!document.querySelector('[data-tool=pdfs]')"),'Interface');
}
async function checkError(pattern){await invoke('check');await wait(async()=>(await state()).phase==='error','Recusa da publicação');assert.match((await state()).message,pattern);await native('assist-notice');}
async function signedData(change){const bytes=Buffer.from(JSON.stringify({...data,...change}));await writeFile(manifestFile,JSON.stringify({...envelope,payload:bytes.toString('base64url'),signature:sign(null,bytes,privateKey).toString('base64url')}));}
try{
    let existingProfile=false;
    try{await readFile(path.join(profile,'Preferences'));existingProfile=true;}
    catch(error){if(error.code!=='ENOENT')throw error;}
    if(existingProfile)assert.equal(await readFile(owner,'utf8'),'desktop-assisted-installed','Não reutilizar perfil desconhecido.');
    await mkdir(profile,{recursive:true});await writeFile(owner,'desktop-assisted-installed');
    assert.ok(installation.startsWith(folder+path.sep)&&(await realpath(folder)).toLowerCase().startsWith((await realpath(tmpdir())).toLowerCase()+path.sep));
    // Apenas provisionamento inicial do fixture isolado. Nenhum instalador é executado pelo aplicativo.
    await new Promise((resolve,reject)=>{const process=spawn(setup,['/S','/D='+installation],{windowsHide:true,stdio:'ignore'});process.once('error',reject);process.once('exit',code=>code===0?resolve():reject(new Error('NSIS base: '+code)));});
    await startServer();await launch();await wait(async()=>(await state()).phase==='available','Detecção no startup');
    assert.equal((await state()).currentVersion,base);assert.equal((await state()).version,target);assert.match((await state()).notes.join(' '),/Ed25519/);
    assert.equal(await evaluate('typeof require'), 'undefined');
    await invoke('details');await native('assist-later');assert.equal((await state()).phase,'available');
    await click('#desktop-assisted-notice button:last-child');assert.equal(await evaluate("document.querySelector('#desktop-assisted-notice').hidden"),true);
    await click('[data-tool=qrcodes]');await click('[data-tool=inicio]');
    await screenshot('01-adiado');
    console.log('PASSOU: .5 instalada, detecção/notas da .6, Mais tarde e uso de QR sem download/instalação automática.');
    const bad=JSON.parse(manifest);bad.payload=Buffer.from(JSON.stringify({...data,notes:['adulterado']})).toString('base64url');await writeFile(manifestFile,JSON.stringify(bad));await checkError(/Assinatura/);
    await signedData({sequence:data.sequence+1,version:'1.0.0-experimental.4',installer:{...data.installer,path:data.installer.path.replace(target,'1.0.0-experimental.4')}});await checkError(/downgrade/);
    await writeFile(manifestFile,manifest);
    const artifact=await open(installerFile,'r+'),first=Buffer.alloc(1);await artifact.read(first,0,1,0);
    try{await artifact.write(Buffer.from([first[0]^1]),0,1,0);await invoke('check');await native('assist-download');await wait(async()=>(await state()).phase==='error','Integridade adulterada',90000);assert.match((await state()).message,/Integridade/);}
    finally{await artifact.write(first,0,1,0);await artifact.close();}
    console.log('PASSOU: app instalado recusou manifesto adulterado, downgrade assinado e instalador modificado.');
    await invoke('check');await native('assist-download');await wait(async()=>(await state()).phase==='ready','Download validado',90000);
    assert.equal(child.exitCode,null);await screenshot('02-verificado');
    const prefs={'4printshops.assisted.preferencia':'preservada','4printshops.pdf.formatos.v1':JSON.stringify({versao:1,formatos:[{id:'custom-assistido',nome:'Teste assistido',largura:123.5,altura:234.5}]})};
    await evaluate(`Object.entries(${JSON.stringify(prefs)}).forEach(([k,v])=>localStorage.setItem(k,v))`);
    const pdf=await PDFDocument.create();pdf.addPage([200,300]);const original=Buffer.from(await pdf.save()),pdfPath=path.join(folder,'original.pdf');await writeFile(pdfPath,original);
    await click('[data-tool=pdfs]');await wait(()=>evaluate("!document.querySelector('#pdf-arquivo').disabled"),'PDF pronto');
    await evaluate('void document.querySelector("#pdf-arquivo").click()');await native('open',pdfPath);
    await wait(()=>evaluate("document.querySelector('#pdf-nome').textContent==='original.pdf'&&!document.querySelector('#pdf-arquivo').disabled"),'PDF carregado');
    await click('#pdf-girar-direita');await wait(()=>evaluate("!document.querySelector('#pdf-arquivo').disabled"),'Rotação');await click('[data-tool=inicio]');
    await invoke('prepare');await native('assist-notice');assert.equal(child.exitCode,null);
    await click('[data-tool=pdfs]');assert.equal(await evaluate("document.querySelector('#pdf-alteracoes-pendentes').hidden"),false);
    const destination=path.join(folder,'exportado.pdf');await click('#pdf-salvar-copia');await click('#app-exportacao-confirmar');await native('save',destination);
    await wait(()=>evaluate("document.querySelector('#pdf-alteracoes-pendentes').hidden"),'Exportação concluída');await click('#app-exportacao-cancelar');
    const exported=await PDFDocument.load(await readFile(destination));assert.equal(exported.getPage(0).getRotation().angle,90);
    await invoke('prepare');await native('assist-prepare-later');assert.equal((await state()).phase,'ready');
    await invoke('prepare');await native('assist-reveal');await wait(()=>evaluate("!document.querySelector('#app-shell').inert"),'Desbloqueio');
    const files=await readdir(path.join(profile,'assisted-updates','installers'));assert.equal(files.filter(f=>f.endsWith('.exe')).length,1);assert.equal(files.some(f=>f.endsWith('.part')),false);
    assert.equal(child.exitCode,null);assert.deepEqual(await readFile(pdfPath),original);
    console.log('PASSOU: PDF alterado em aba oculta bloqueia preparo; exportação real preserva giro/original; cancelar e localizar não executam instalador nem fecham aplicativo.');
    await stopServer();await checkError(/ECONNREFUSED|indisponível|connect/i);await click('[data-tool=qrcodes]');await click('[data-tool=inicio]');
    await evaluate('window.close()');await wait(()=>child.exitCode!==null,'Fechamento normal');socket.close();
    await launch();await wait(async()=>(await state()).phase==='error','Startup offline');
    for(const [key,value]of Object.entries(prefs))assert.equal(await evaluate(`localStorage.getItem(${JSON.stringify(key)})`),value);
    await startServer();await invoke('check');await native('assist-later');assert.equal((await state()).phase,'available');
    assert.deepEqual(exceptions,[]);await screenshot('03-recuperado');
    await evaluate('window.close()');await wait(()=>child.exitCode!==null,'Fechamento final');socket.close();
    await writeFile(path.join(folder,'result.json'),JSON.stringify({base,target,installation,profile,caddy:'2.11.7',publicationVerified:true,downloadVerified:true,prepared:true,targetExecuted:false,automaticInstall:false,preferencesPreservedAcrossRestart:true},null,2));
    console.log('PASSOU: offline/recuperação sem bloquear uso, preferências preservadas ao reiniciar. Instalador .6 NÃO executado; instalação automática NÃO homologada.');
}catch(error){try{await screenshot('falha');console.error('Estado:',await state());}catch{}console.error(error);process.exitCode=1;}
finally{await writeFile(manifestFile,manifest);socket?.close();if(child&&child.exitCode===null)child.kill();await stopServer();console.log('Artefatos: '+folder);}
