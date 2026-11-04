// Pacote instalado. Cliques em Imprimir são permitidos somente na fila virtual Microsoft.
import assert from 'node:assert/strict';
import {spawn,execFile} from 'node:child_process';
import {promisify} from 'node:util';
import {mkdir,readFile,writeFile,unlink} from 'node:fs/promises';
import path from 'node:path';
import {setTimeout as pause} from 'node:timers/promises';
import {PDFDocument} from '../public/vendor/pdf-lib/pdf-lib.js';

const exec=promisify(execFile),folder=path.resolve('.desktop-test/print12/installed');await mkdir(folder,{recursive:true});
const installation=path.resolve('.desktop-test/hold-lab/installed-validation/installed');
const setup=path.resolve('dist-desktop/hold-native-test/4PrintShops-Hold-Test-1.0.0-experimental.12-Windows-x64-Setup.exe');
const exe=path.join(installation,'4PrintShops Hold Test.exe'),profile=path.join(process.env.APPDATA,'4PrintShops Hold Test');
const ps=path.join(process.env.SystemRoot,'System32/WindowsPowerShell/v1.0/powershell.exe');
let app,socket,serial=0;const pending=new Map(),exceptions=[],watchers=[];
const wait=async(fn,label,timeout=60000)=>{const end=Date.now()+timeout;while(Date.now()<end){if(await fn())return;await pause(100);}throw Error('Tempo esgotado '+label);};
const cdp=(method,params={})=>new Promise((resolve,reject)=>{const id=++serial;const timer=setTimeout(()=>{pending.delete(id);reject(Error(method));},60000);pending.set(id,{resolve:value=>{clearTimeout(timer);resolve(value);},reject});socket.send(JSON.stringify({id,method,params}));});
const evaluate=async expression=>{const result=await cdp('Runtime.evaluate',{expression,awaitPromise:true,returnByValue:true,userGesture:true});if(result.exceptionDetails)throw Error(result.exceptionDetails.exception?.description||result.exceptionDetails.text);return result.result.value;};
const click=selector=>evaluate(`document.querySelector(${JSON.stringify(selector)}).click()`);
const ready=()=>wait(()=>evaluate("!document.querySelector('#print-config-close').disabled"),'controle liberado');
const powershell=async command=>(await exec(ps,['-NoProfile','-NonInteractive','-Command',command],{windowsHide:true,timeout:40000})).stdout.trim();
const quote=value=>"'"+value.replaceAll("'","''")+"'";
async function native(action){
    const pid=Number(await powershell(`$p=Get-CimInstance Win32_Process -Filter "Name='4PrintShops.PrintConfig.exe'" | Where-Object {$_.ParentProcessId -eq ${app.pid}};if(@($p).Count -ne 1){throw 'Auxiliar não identificado'};$p.ProcessId`));
    const source=await readFile(new URL('./helpers/windows-print-properties.ps1',import.meta.url),'utf8');
    await powershell(`& {${source}} -DriverProcess ${pid} -Action ${quote(action)}`);
}
const select=async printer=>{await evaluate(`document.querySelector('#print-config-printer').value=${JSON.stringify(printer)};document.querySelector('#print-config-printer').dispatchEvent(new Event('change'))`);await ready();};
async function monitor(printer){
    const source=await readFile(new URL('./helpers/windows-print-defaults.ps1',import.meta.url),'utf8');
    const child=spawn(ps,['-NoProfile','-NonInteractive','-Command',`& {${source}} -PrinterName ${quote(printer)}`],{windowsHide:true,stdio:['pipe','pipe','pipe']});watchers.push(child);let text='';child.stdout.on('data',data=>text+=data);
    await wait(()=>text.includes('READY'),'monitor somente leitura');
    return async()=>{text='';child.stdin.write('check\n');await wait(()=>text.includes('CHANGED'),'preferências');assert.equal(text.trim(),'UNCHANGED');};
}
try{
    assert.equal(await powershell("@(Get-CimInstance Win32_Process | Where-Object {$_.Name -eq '4PrintShops Hold Test.exe'}).Count"),'0','Feche conscientemente a Hold Test antes do upgrade');
    assert.equal(await powershell("(Get-ItemProperty 'HKCU:/Software/Microsoft/Windows NT/CurrentVersion/Windows').LegacyDefaultPrinterMode"),'1');
    const defaultBefore=await powershell('(Get-CimInstance Win32_Printer | Where-Object Default).Name');
    const registration=JSON.parse(await powershell("Get-ItemProperty 'HKCU:/Software/Microsoft/Windows/CurrentVersion/Uninstall/90ff2ff0-f013-571c-b109-7edb265dacc8' | Select-Object DisplayVersion,UninstallString | ConvertTo-Json -Compress"));
    assert.ok(registration.UninstallString.startsWith('"'+path.join(installation,'Uninstall 4PrintShops Hold Test.exe')+'"'),'Recusar instalação fora do laboratório conhecido');
    await new Promise((resolve,reject)=>{const installer=spawn(setup,['/S'],{windowsHide:true,stdio:'ignore'});installer.once('error',reject);installer.once('exit',code=>code===0?resolve():reject(Error('NSIS '+code)));});
    await unlink(path.join(profile,'DevToolsActivePort')).catch(()=>{});
    const env={...process.env};delete env.ELECTRON_RUN_AS_NODE;
    app=spawn(exe,['--remote-debugging-port=0','--open-pdf',path.resolve('.desktop-test/print12/native/source.pdf')],{env,windowsHide:true,stdio:'ignore'});
    let port;await wait(async()=>{try{port=(await readFile(path.join(profile,'DevToolsActivePort'),'utf8')).split('\n')[0];return !!port;}catch{return false;}},'aplicativo instalado');
    let target;await wait(async()=>{try{target=(await (await fetch(`http://127.0.0.1:${port}/json`)).json()).find(item=>item.type==='page');return !!target;}catch{return false;}},'janela');
    socket=new WebSocket(target.webSocketDebuggerUrl);await new Promise(resolve=>socket.addEventListener('open',resolve,{once:true}));
    socket.addEventListener('message',event=>{const message=JSON.parse(event.data);if(message.id){const call=pending.get(message.id);if(message.error)call?.reject(Error(message.error.message));else call?.resolve(message.result);pending.delete(message.id);}if(message.method==='Runtime.exceptionThrown')exceptions.push(message.params.exceptionDetails.text);});
    await cdp('Runtime.enable');
    await wait(()=>evaluate("!!document.querySelector('#pdf-visualizacao canvas')&&!document.querySelector('#pdf-arquivo').disabled"),'PDF');
    assert.equal(await evaluate("localStorage.getItem('4printshops.print-config-check')"),'preservado');
    await click('#pdf-girar-direita');await wait(()=>evaluate("!document.querySelector('#pdf-arquivo').disabled"),'giro');
    assert.equal(await evaluate("document.querySelector('#pdf-alteracoes-pendentes').hidden"),false);
    await click('#app-menu-arquivo');await click('[data-command="configurar-impressao"]');await ready();
    const unchanged=[];for(const printer of ['WF-C21000 Series(Rede)','KONICA MINOLTA C6100PS','Microsoft Print to PDF'])unchanged.push(await monitor(printer));
    await select('WF-C21000 Series(Rede)');
    assert.equal(await evaluate("document.querySelector('#print-config-print').disabled"),false,'Epson pode imprimir sem Hold');
    assert.equal(await evaluate("document.querySelector('#print-config-diagnostics').hidden"),true);
    const before=await evaluate("document.querySelector('#print-config-summary').textContent");
    await click('#print-config-properties');await native('cancel');await ready();
    assert.equal(await evaluate("document.querySelector('#print-config-summary').textContent"),before);
    await click('#print-config-properties');await native('accept');await ready();
    assert.match(await evaluate("document.querySelector('#print-config-status').textContent"),/confirmada/);
    assert.match(await evaluate("document.querySelector('#print-config-scope').textContent"),/3 página/);
    let image=await cdp('Page.captureScreenshot');await writeFile(path.join(folder,'epson-sem-hold.png'),Buffer.from(image.data,'base64'));
    // Nenhum clique em Imprimir na Epson ou Konica em qualquer parte deste teste.
    await select('KONICA MINOLTA C6100PS');assert.equal(await evaluate("document.querySelector('#print-config-print').disabled"),false);
    assert.equal(await evaluate("document.querySelector('#print-config-diagnostics').hidden"),false);
    await select('Microsoft Print to PDF');
    const saveSource=await readFile(new URL('./helpers/windows-save-dialog.ps1',import.meta.url),'utf8');
    const printSaveSource=saveSource.replaceAll("'Salvar cópia PDF'","'Salvar saída da impressora PDF'");
    const outputs=[];
    for(const scale of ['100','120','custom']){
        await evaluate(`document.querySelector('#print-config-scale').value='${scale}';document.querySelector('#print-config-scale').dispatchEvent(new Event('change'));document.querySelector('#print-config-custom').value='87,5';document.querySelector('#print-config-custom').dispatchEvent(new Event('input'))`);
        assert.equal(await evaluate("document.querySelector('#print-config-printer').value"),'Microsoft Print to PDF','Guarda obrigatória antes do clique de impressão');
        await click('#print-config-print');
        const output=path.join(folder,'impresso-'+scale+'-'+Date.now()+'.pdf');
        await powershell(`& {${printSaveSource}} -DialogProcess ${app.pid} -Destination ${quote(output)} -Action save`);await ready();
        assert.match(await evaluate("document.querySelector('#print-config-status').textContent"),/Trabalho.*enviado/);
        const pdf=await PDFDocument.load(await readFile(output));assert.equal(pdf.getPageCount(),3);outputs.push(output);
        assert.equal(await evaluate("document.querySelector('#pdf-alteracoes-pendentes').hidden"),false,'Impressão não marca exportação');
    }
    image=await cdp('Page.captureScreenshot');await writeFile(path.join(folder,'virtual-impresso.png'),Buffer.from(image.data,'base64'));
    for(const check of unchanged)await check();
    assert.equal(await powershell('(Get-CimInstance Win32_Printer | Where-Object Default).Name'),defaultBefore);
    await click('#print-config-close');await click('#app-menu-arquivo');await click('[data-command="configurar-impressao"]');await ready();
    assert.equal(await evaluate("document.querySelectorAll('#print-config-dialog').length"),1);
    await click('#print-config-close');await click('#pdf-salvar-copia');await click('#app-exportacao-confirmar');
    const exported=path.join(folder,'workspace-'+Date.now()+'.pdf');
    await powershell(`& {${saveSource}} -DialogProcess ${app.pid} -Destination ${quote(exported)} -Action save`);
    await wait(()=>evaluate("document.querySelector('#pdf-alteracoes-pendentes').hidden"),'exportação');
    const pdf=await PDFDocument.load(await readFile(exported));assert.equal(pdf.getPageCount(),3);assert.equal(pdf.getPage(0).getRotation().angle,270);
    assert.deepEqual(exceptions,[]);
    const version=await powershell("(Get-ItemProperty 'HKCU:/Software/Microsoft/Windows/CurrentVersion/Uninstall/90ff2ff0-f013-571c-b109-7edb265dacc8').DisplayVersion");assert.equal(version,'1.0.0-experimental.12');
    await writeFile(path.join(folder,'result.json'),JSON.stringify({passed:true,previousVersion:registration.DisplayVersion,version,sameDirectoryAndProfile:true,preferencePreserved:true,epsonPropertiesCancelAndConfirm:true,printEnabledWithoutHold:true,physicalSendingTested:false,virtualPrintedPages:3,virtualScales:[100,120,87.5],outputs,workspaceDirtyPreserved:true,exportedPages:3,rotation:270,defaultsPreserved:true,defaultPrinter:defaultBefore},null,2));
    console.log('PASSOU: instalação .12, preferências, Epson sem Hold e propriedades, Konica disponível, três impressões virtuais, três páginas mistas, giro do workspace, exportação e padrão preservados. Nenhum clique de impressão em fila física.');
}catch(error){console.error(error);try{console.error(await evaluate("document.querySelector('#print-config-status')?.textContent"));}catch{}process.exitCode=1;}
finally{for(const watcher of watchers){watcher.stdin.end('quit\n');watcher.kill();}socket?.close();if(app&&app.exitCode===null)app.kill();}
