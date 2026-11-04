// Instala e prepara o ensaio real. A única ação automatizada na confirmação de envio é Cancelar.
import assert from 'node:assert/strict';
import {spawn,execFile} from 'node:child_process';
import {promisify} from 'node:util';
import {mkdir,readFile,writeFile,unlink} from 'node:fs/promises';
import path from 'node:path';
import {setTimeout as pause} from 'node:timers/promises';
import {PDFDocument} from '../public/vendor/pdf-lib/pdf-lib.js';

const exec=promisify(execFile);
const folder=path.resolve('.desktop-test/hold11/installed-validation');
// A mesma instalação/perfil da experimental.9. Recusar outro destino registrado.
const installation=path.resolve('.desktop-test/hold-lab/installed-validation/installed');
const setup=path.resolve('dist-desktop/hold-native-test/4PrintShops-Hold-Test-1.0.0-experimental.11-Windows-x64-Setup.exe');
const exe=path.join(installation,'4PrintShops Hold Test.exe');
const profile=path.join(process.env.APPDATA,'4PrintShops Hold Test');
const ps=path.join(process.env.SystemRoot,'System32/WindowsPowerShell/v1.0/powershell.exe');
let app,socket,serial=0,watcher;
const pending=new Map(),exceptions=[];
const wait=async(fn,label,timeout=60000)=>{const end=Date.now()+timeout;while(Date.now()<end){if(await fn())return;await pause(100);}throw Error('Tempo esgotado '+label);};
const cdp=(method,params={})=>new Promise((resolve,reject)=>{const id=++serial;const timer=setTimeout(()=>{pending.delete(id);reject(Error(method));},60000);pending.set(id,{resolve:value=>{clearTimeout(timer);resolve(value);},reject});socket.send(JSON.stringify({id,method,params}));});
const evaluate=async expression=>{const result=await cdp('Runtime.evaluate',{expression,awaitPromise:true,returnByValue:true,userGesture:true});if(result.exceptionDetails)throw Error(result.exceptionDetails.exception?.description||result.exceptionDetails.text);return result.result.value;};
const click=selector=>evaluate(`document.querySelector(${JSON.stringify(selector)}).click()`);
const ready=()=>wait(()=>evaluate("!document.querySelector('#print-config-close').disabled"),'controle liberado');
const powershell=async command=>(await exec(ps,['-NoProfile','-NonInteractive','-Command',command],{windowsHide:true,timeout:40000})).stdout.trim();
const quote=value=>"'"+value.replaceAll("'","''")+"'";
async function native(action){
    const pid=Number(await powershell(`$p=Get-CimInstance Win32_Process -Filter "Name='4PrintShops.PrintConfig.exe'" | Where-Object {$_.ParentProcessId -eq ${app.pid}};if(@($p).Count -ne 1){throw 'Auxiliar de teste não identificado'};$p.ProcessId`));
    const source=await readFile(new URL('./helpers/windows-print-properties.ps1',import.meta.url),'utf8');
    await powershell(`& {${source}} -DriverProcess ${pid} -Action ${quote(action)}`);
}
await mkdir(folder,{recursive:true});
try {
    const defaultBefore=await powershell("(Get-CimInstance Win32_Printer | Where-Object Default).Name");
    assert.equal(await powershell("(Get-ItemProperty -LiteralPath 'HKCU:/Software/Microsoft/Windows NT/CurrentVersion/Windows').LegacyDefaultPrinterMode"),'1','Gerência automática deve estar desativada antes do teste');
    assert.equal(await powershell("@(Get-CimInstance Win32_Process | Where-Object {$_.Name -eq '4PrintShops Hold Test.exe'}).Count"),'0','Feche o Hold Test conscientemente antes do upgrade');
    const oldRegistration=JSON.parse(await powershell("Get-ItemProperty -LiteralPath 'HKCU:/Software/Microsoft/Windows/CurrentVersion/Uninstall/90ff2ff0-f013-571c-b109-7edb265dacc8' | Select-Object DisplayVersion,UninstallString | ConvertTo-Json -Compress"));
    assert.ok(oldRegistration.UninstallString.startsWith('"'+path.join(installation,'Uninstall 4PrintShops Hold Test.exe')+'"'),'Recusar atualização de uma instalação fora do laboratório conhecido');
    await new Promise((resolve,reject)=>{const installer=spawn(setup,['/S'],{windowsHide:true,stdio:'ignore'});installer.once('error',reject);installer.once('exit',code=>code===0?resolve():reject(Error('NSIS '+code)));});
    await unlink(path.join(profile,'DevToolsActivePort')).catch(()=>{});
    const env={...process.env};delete env.ELECTRON_RUN_AS_NODE;
    const fixture=path.join(installation,'resources/hold-native/synthetic.pdf');
    app=spawn(exe,['--remote-debugging-port=0','--open-pdf',fixture],{env,windowsHide:true,stdio:'ignore'});
    let port;
    await wait(async()=>{try{port=(await readFile(path.join(profile,'DevToolsActivePort'),'utf8')).split('\n')[0];return !!port;}catch{return false;}},'aplicativo instalado');
    let target;
    await wait(async()=>{try{target=(await (await fetch(`http://127.0.0.1:${port}/json`)).json()).find(item=>item.type==='page');return !!target;}catch{return false;}},'janela');
    socket=new WebSocket(target.webSocketDebuggerUrl);await new Promise(resolve=>socket.addEventListener('open',resolve,{once:true}));
    socket.addEventListener('message',event=>{const message=JSON.parse(event.data);if(message.id){const call=pending.get(message.id);if(message.error)call?.reject(Error(message.error.message));else call?.resolve(message.result);pending.delete(message.id);}if(message.method==='Runtime.exceptionThrown')exceptions.push(message.params.exceptionDetails.text);});
    await cdp('Runtime.enable');
    await wait(()=>evaluate("!!document.querySelector('#pdf-visualizacao canvas')&&!document.querySelector('#pdf-arquivo').disabled"),'PDF sintético');
    assert.equal(await evaluate("localStorage.getItem('4printshops.print-config-check')"),'preservado','Preferência anterior preservada');
    await click('#app-menu-arquivo');await click('[data-command="configurar-impressao"]');await ready();
    await evaluate("document.querySelector('#print-config-printer').value='KONICA MINOLTA C6100PS';document.querySelector('#print-config-printer').dispatchEvent(new Event('change'))");await ready();
    await click('#print-config-fixture');await ready();
    await click('#print-config-close');await ready();
    for(let n=0;n<2;n++){await click('#pdf-girar-direita');await wait(()=>evaluate("!document.querySelector('#pdf-arquivo').disabled"),'giro');}
    assert.equal(await evaluate("document.querySelector('#pdf-alteracoes-pendentes').hidden"),false);
    const monitorSource=await readFile(new URL('./helpers/windows-print-defaults.ps1',import.meta.url),'utf8');
    watcher=spawn(ps,['-NoProfile','-NonInteractive','-Command',`& {${monitorSource}} -PrinterName 'KONICA MINOLTA C6100PS'`],{windowsHide:true,stdio:['pipe','pipe','pipe']});
    let monitor='';watcher.stdout.on('data',chunk=>monitor+=chunk);
    await wait(()=>monitor.includes('READY'),'monitor somente leitura');
    await click('#app-menu-arquivo');await click('[data-command="configurar-impressao"]');await ready();
    await evaluate("document.querySelector('#print-config-printer').value='KONICA MINOLTA C6100PS';document.querySelector('#print-config-printer').dispatchEvent(new Event('change'))");await ready();
    await click('#print-config-properties');await native('a3-sizes');await native('workflow');await pause(250);await native('hold');await native('accept');await ready();
    await click('#print-config-verify-hold');await ready();
    assert.equal(await evaluate("document.querySelector('#print-config-scale').disabled"),false);
    await evaluate("document.querySelector('#print-config-scale').value='120';document.querySelector('#print-config-scale').dispatchEvent(new Event('change'))");
    assert.equal(await evaluate("document.querySelector('#print-config-send-hold').disabled"),false);
    await click('#print-config-send-hold');
    const cancelSource=await readFile(new URL('./helpers/windows-save-dialog.ps1',import.meta.url),'utf8');
    // Nunca adicionar um comando para confirmar o envio neste harness.
    await powershell(`& {${cancelSource}} -DialogProcess ${app.pid} -Action hold-cancel`);
    await ready();
    assert.match(await evaluate("document.querySelector('#print-config-status').textContent"),/cancelado/);
    assert.equal(await evaluate("document.querySelector('#print-config-new-operation').hidden"),false);
    const newOperation=async()=>{
        await click('#print-config-new-operation');
        await powershell(`& {${cancelSource}} -DialogProcess ${app.pid} -Action hold-new`);await ready();
        assert.equal(await evaluate("document.querySelector('#print-config-send-hold').disabled"),true,'Nova sessão exige Hold novamente');
    };
    await newOperation();
    await click('#print-config-properties');await native('a4-sizes');await native('workflow');await pause(250);await native('hold');await native('accept');await ready();
    await click('#print-config-verify-hold');await ready();
    await click('#print-config-send-hold');
    await powershell(`& {${cancelSource}} -DialogProcess ${app.pid} -Action crop-continue`);
    await powershell(`& {${cancelSource}} -DialogProcess ${app.pid} -Action hold-cancel`);await ready();
    await newOperation();
    await click('#print-config-properties');await native('workflow');await pause(250);await native('hold');await native('accept');await ready();
    await click('#print-config-verify-hold');await ready();
    await evaluate("document.querySelector('#print-config-scale').value='custom';document.querySelector('#print-config-scale').dispatchEvent(new Event('change'));document.querySelector('#print-config-custom').value='87,5';document.querySelector('#print-config-custom').dispatchEvent(new Event('input'))");
    assert.match(await evaluate("document.querySelector('#print-config-measures').textContent"),/87,5|87.5/);
    await click('#print-config-send-hold');
    await powershell(`& {${cancelSource}} -DialogProcess ${app.pid} -Action hold-cancel`);await ready();
    assert.match(await evaluate("document.querySelector('#print-config-ps-details').textContent"),/SHKMColorMode=/);
    assert.equal(await evaluate("document.querySelector('#pdf-alteracoes-pendentes').hidden"),false);
    monitor='';watcher.stdin.write('check\n');await wait(()=>monitor.includes('CHANGED'),'preferências');assert.equal(monitor.trim(),'UNCHANGED');
    const jobs=await powershell("@(Get-PrintJob -PrinterName 'KONICA MINOLTA C6100PS' | Where-Object {$_.DocumentName -like '4PS HOLD EXP *'}).Count");assert.equal(jobs,'0');
    assert.equal(await powershell("(Get-CimInstance Win32_Printer | Where-Object Default).Name"),defaultBefore,'Captura preserva a impressora padrão');
    const image=await cdp('Page.captureScreenshot');await writeFile(path.join(folder,'cancelado.png'),Buffer.from(image.data,'base64'));
    await click('#print-config-close');await click('#app-menu-arquivo');await click('[data-command="configurar-impressao"]');await ready();
    assert.equal(await evaluate("document.querySelectorAll('#print-config-dialog').length"),1);
    await click('#print-config-close');
    await click('#pdf-salvar-copia');await click('#app-exportacao-confirmar');
    const destination=path.join(folder,'resultado-'+Date.now()+'.pdf');
    await powershell(`& {${cancelSource}} -DialogProcess ${app.pid} -Destination ${quote(destination)} -Action save`);
    await wait(()=>evaluate("document.querySelector('#pdf-alteracoes-pendentes').hidden"),'exportação posterior');
    const exported=await PDFDocument.load(await readFile(destination));assert.equal(exported.getPageCount(),1);assert.equal(exported.getPage(0).getRotation().angle,180);
    assert.deepEqual(exceptions,[]);
    const versionAfter=await powershell("(Get-ItemProperty -LiteralPath 'HKCU:/Software/Microsoft/Windows/CurrentVersion/Uninstall/90ff2ff0-f013-571c-b109-7edb265dacc8').DisplayVersion");assert.equal(versionAfter,'1.0.0-experimental.11');
    await writeFile(path.join(folder,'result.json'),JSON.stringify({passed:true,previousVersion:oldRegistration.DisplayVersion,version:versionAfter,installation,sameInstallationIdentity:true,preferencePreserved:true,editedWorkspaceRotation:180,psCapturedAndInterpreted:true,newOperations:2,customScale87_5:true,cropConfirmationSeparate:true,finalConfirmationsCanceled:3,dirtyPreserved:true,exportedPages:1,exportedRotation:180,defaultsUnchanged:true,defaultPrinterPreserved:defaultBefore,testJobsInProductionQueue:0,physicalSendingTested:false},null,2));
    console.log('PASSOU: instalado, PDFWorkspace girado, captura e PS Hold/A3/120% verificados, confirmação cancelada, pendências e defaults preservados. Envio físico não executado.');
} catch(error){console.error(error);try{console.error(await evaluate("document.querySelector('#print-config-status')?.textContent"));}catch{}process.exitCode=1;}
finally{watcher?.stdin.end('quit\n');watcher?.kill();socket?.close();if(app&&app.exitCode===null)app.kill();}
