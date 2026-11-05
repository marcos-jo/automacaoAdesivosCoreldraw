// Comparação visual real. Nunca clicar em Imprimir em uma fila física.
import assert from 'node:assert/strict';
import {spawn,execFile} from 'node:child_process';
import {promisify} from 'node:util';
import {mkdir,readFile,writeFile,unlink} from 'node:fs/promises';
import path from 'node:path';
import {setTimeout as pause} from 'node:timers/promises';
import {PDFDocument} from '../public/vendor/pdf-lib/pdf-lib.js';
import {createPaperFixture} from './helpers/windows-print-fixtures.mjs';
const before=process.env.PRINT_UI_STAGE==='before',phase=before?'before':'after';
const folder=path.resolve('.desktop-test/theme15/installed/'+phase);await mkdir(folder,{recursive:true});
const installation=path.resolve(process.env.PRINT_UI_INSTALL_DIR || '.desktop-test/hold-lab/installed-validation/installed/4PrintShops Hold Test');
const exe=path.join(installation,'4PrintShops Hold Test.exe'),profile=path.join(process.env.APPDATA,'4PrintShops Hold Test');
const exec=promisify(execFile),ps=path.join(process.env.SystemRoot,'System32/WindowsPowerShell/v1.0/powershell.exe');
const powershell=async command=>(await exec(ps,['-NoProfile','-NonInteractive','-Command',command],{windowsHide:true,timeout:40000})).stdout.trim();
const quote=value=>"'"+value.replaceAll("'","''")+"'";
let app,socket,serial=0;const pending=new Map(),exceptions=[],watchers=[];
const wait=async(fn,label,timeout=60000)=>{const end=Date.now()+timeout;while(Date.now()<end){if(await fn())return;await pause(100);}throw Error('Tempo esgotado '+label);};
const cdp=(method,params={})=>new Promise((resolve,reject)=>{const id=++serial;const timer=setTimeout(()=>{pending.delete(id);reject(Error(method));},60000);pending.set(id,{resolve:value=>{clearTimeout(timer);resolve(value);},reject});socket.send(JSON.stringify({id,method,params}));});
const evaluate=async expression=>{const result=await cdp('Runtime.evaluate',{expression,awaitPromise:true,returnByValue:true,userGesture:true});if(result.exceptionDetails)throw Error(result.exceptionDetails.exception?.description||result.exceptionDetails.text);return result.result.value;};
const click=selector=>evaluate(`document.querySelector(${JSON.stringify(selector)}).click()`);
const ready=()=>wait(()=>evaluate("!document.querySelector('#print-config-close').disabled"),'controle liberado');
const image=async name=>{await pause(250);const result=await cdp('Page.captureScreenshot');await writeFile(path.join(folder,name+'.png'),Buffer.from(result.data,'base64'));};
const metrics=async(width,height,deviceScaleFactor=1)=>{await cdp('Emulation.setDeviceMetricsOverride',{width,height,deviceScaleFactor,mobile:false});await pause(250);};
const select=async name=>{await ready();await evaluate(`document.querySelector('#print-config-printer').value=${JSON.stringify(name)};document.querySelector('#print-config-printer').dispatchEvent(new Event('change'))`);await ready();};
async function native(action){
    const pid=Number(await powershell(`$p=Get-CimInstance Win32_Process -Filter "Name='4PrintShops.PrintConfig.exe'" | Where-Object {$_.ParentProcessId -eq ${app.pid}};if(@($p).Count -ne 1){throw 'Auxiliar não identificado'};$p.ProcessId`));
    const source=await readFile(new URL('./helpers/windows-print-properties.ps1',import.meta.url),'utf8');
    await powershell(`& {${source}} -DriverProcess ${pid} -Action ${quote(action)}`);
}
const visible=async selector=>evaluate(`(()=>{const r=document.querySelector(${JSON.stringify(selector)}).getBoundingClientRect();return r.top>=0&&r.left>=0&&r.bottom<=innerHeight&&r.right<=innerWidth&&r.width>0&&r.height>0})()`);
async function monitor(printer){
    const source=await readFile(new URL('./helpers/windows-print-defaults.ps1',import.meta.url),'utf8');
    const child=spawn(ps,['-NoProfile','-NonInteractive','-Command',`& {${source}} -PrinterName ${quote(printer)}`],{windowsHide:true,stdio:['pipe','pipe','pipe']});watchers.push(child);let text='';child.stdout.on('data',data=>text+=data);
    await wait(()=>text.includes('READY'),'monitor somente leitura');
    return async()=>{text='';child.stdin.write('check\n');await wait(()=>text.includes('CHANGED'),'preferências');assert.equal(text.trim(),'UNCHANGED');};
}
try{
    assert.equal(await powershell("@(Get-CimInstance Win32_Process | Where-Object {$_.Name -eq '4PrintShops Hold Test.exe'}).Count"),'0','Feche conscientemente o Hold Test');
    const reg="HKCU:/Software/Microsoft/Windows/CurrentVersion/Uninstall/90ff2ff0-f013-571c-b109-7edb265dacc8";
    const registration=JSON.parse(await powershell(`Get-ItemProperty '${reg}' | Select-Object DisplayVersion,UninstallString | ConvertTo-Json -Compress`));
    assert.ok(registration.UninstallString.startsWith('"'+path.join(installation,'Uninstall 4PrintShops Hold Test.exe')+'"'));
    const defaultBefore=await powershell('(Get-CimInstance Win32_Printer | Where-Object Default).Name');
    if(!before){
        const setup=path.resolve('dist-desktop/hold-native-test/4PrintShops-Hold-Test-1.0.0-experimental.15-Windows-x64-Setup.exe');
        await new Promise((resolve,reject)=>{const child=spawn(setup,['/S'],{windowsHide:true,stdio:'ignore'});child.once('error',reject);child.once('exit',code=>code===0?resolve():reject(Error('NSIS '+code)));});
    }
    const source=await PDFDocument.load(await createPaperFixture());
    const second=await PDFDocument.load(await createPaperFixture({width:297,height:420}));
    source.addPage((await source.copyPages(second,[0]))[0]);
    const fixture=path.resolve('.desktop-test/theme15/installed/source.pdf');await writeFile(fixture,await source.save());
    await unlink(path.join(profile,'DevToolsActivePort')).catch(()=>{});
    const env={...process.env};delete env.ELECTRON_RUN_AS_NODE;
    app=spawn(exe,['--remote-debugging-port=0','--disable-renderer-backgrounding','--open-pdf',fixture],{env,windowsHide:true,stdio:'ignore'});
    let port;await wait(async()=>{try{port=(await readFile(path.join(profile,'DevToolsActivePort'),'utf8')).split('\n')[0];return !!port;}catch{return false;}},'Electron');
    let target;await wait(async()=>{try{target=(await (await fetch(`http://127.0.0.1:${port}/json`)).json()).find(item=>item.type==='page');return !!target;}catch{return false;}},'janela');
    socket=new WebSocket(target.webSocketDebuggerUrl);await new Promise(resolve=>socket.addEventListener('open',resolve,{once:true}));
    socket.addEventListener('message',event=>{const m=JSON.parse(event.data);if(m.id){const call=pending.get(m.id);if(m.error)call?.reject(Error(m.error.message));else call?.resolve(m.result);pending.delete(m.id);}if(m.method==='Runtime.exceptionThrown')exceptions.push(m.params.exceptionDetails.exception?.description||m.params.exceptionDetails.text);});
    await cdp('Runtime.enable');await cdp('Page.enable');await cdp('Page.bringToFront');await wait(()=>evaluate("!!document.querySelector('#pdf-visualizacao canvas')&&!document.querySelector('#pdf-arquivo').disabled"),'PDF');
    assert.equal(await evaluate("localStorage.getItem('4printshops.print-config-check')"),'preservado');
    await evaluate("(async()=>{const {aplicarTema}=await import('/js/app/theme.js');aplicarTema('dark');})()");
    await metrics(1366,768);await click('#aba-inicio');await image('home-1366');
    if(!before){
        await click('[data-open-pdf]');
        const source=await readFile(new URL('./helpers/windows-save-dialog.ps1',import.meta.url),'utf8');
        await powershell(`& {${source}} -DialogProcess ${app.pid} -Destination ${quote(fixture)} -Action open`);
        await wait(()=>evaluate("!!document.querySelector('#pdf-visualizacao canvas')&&!document.querySelector('#pdf-arquivo').disabled"),'importação pela tela inicial');
    }else await click('#aba-pdfs');
    await image('pdf-1366');
    if(!before){
        const width=await evaluate("document.querySelector('#pdf-visualizacao').clientWidth");
        await click('#pdf-miniaturas-alternar');await pause(200);
        assert.ok(await evaluate("document.querySelector('#pdf-visualizacao').clientWidth")>width);await image('pdf-thumbnails-collapsed');await click('#pdf-miniaturas-alternar');
        await click('#pdf-painel-alternar');await pause(200);assert.ok(await evaluate("document.querySelector('#pdf-visualizacao').clientWidth")>width);await click('#pdf-painel-alternar');
        await click('#pdf-zoom-mais');await pause(300);await click('#pdf-zoom-pagina');
    }
    await metrics(1920,1080);await image('pdf-1920');await click('#aba-inicio');await image('home-1920');await click('#aba-pdfs');
    await metrics(1366,768);await click('#app-menu-arquivo');await click('[data-command="configurar-impressao"]');await ready();
    const checks=[];for(const printer of ['KONICA MINOLTA C6100PS','WF-C21000 Series(Rede)','Microsoft Print to PDF'])checks.push(await monitor(printer));
    await select('KONICA MINOLTA C6100PS');await click('#print-config-properties');await native('a3-sizes');await native('accept');await ready();
    await evaluate("document.querySelector('#print-config-scale').value='120';document.querySelector('#print-config-scale').dispatchEvent(new Event('change'))");
    if(!before)await wait(()=>evaluate("!!document.querySelector('#print-config-preview image')?.getAttribute('href')"),'conteúdo real na prévia');
    await image('konica-a4-a3-120-1366');
    const layouts=[];
    for(const [width,height,dpr,label]of [[1366,768,1,'1366'],[1920,1080,1,'1920'],[1092,614,1.25,'125percent']]){
        await metrics(width,height,dpr);await image('print-'+label);
        const footerVisible=await visible('#print-config-print');
        layouts.push({width,height,dpr,footerVisible});if(!before)assert.equal(footerVisible,true,'Imprimir acessível sem rolar a janela');
    }
    await metrics(1366,768);
    if(!before){await click('#print-config-next');await wait(()=>evaluate("document.querySelector('#print-config-page').textContent.includes('2')"),'página 2');await image('print-page2');await click('#print-config-previous');}
    await select('WF-C21000 Series(Rede)');await image('epson-1366');
    await click('#print-config-properties');await native('cancel');await ready();
    await image('epson-after-cancel');
    // Verificação visual e propriedades, sem enviar trabalho nesta rodada.
    await click('#print-config-close');
    if(!before){
        await click('#pdf-girar-direita');await wait(()=>evaluate("!document.querySelector('#pdf-arquivo').disabled"),'giro');
        assert.equal(await evaluate("document.querySelector('#pdf-alteracoes-pendentes').hidden"),false);
        await click('#app-menu-arquivo');await click('[data-command="configurar-impressao"]');await ready();
        await select('Microsoft Print to PDF');
        await wait(()=>evaluate("!!document.querySelector('#print-config-preview image')?.getAttribute('href')"),'reabertura da prévia');
        assert.equal(await evaluate("document.querySelectorAll('#print-config-dialog').length"),1);
        await click('#print-config-close');assert.equal(await evaluate("document.querySelector('#pdf-alteracoes-pendentes').hidden"),false);
    }
    await click('#app-menu-arquivo'); await click('[data-command="pdf-importar"]');
    const openSource=await readFile(new URL('./helpers/windows-save-dialog.ps1',import.meta.url),'utf8');
    await powershell(`& {${openSource}} -DialogProcess ${app.pid} -Destination ${quote(fixture)} -Action open`);
    await wait(()=>evaluate("document.querySelector('#pdf-confirmar-descarte').open"),'proteção da abertura pelo menu com documento editado');
    await click('#pdf-continuar-editando');
    assert.equal(await evaluate("document.querySelector('#pdf-alteracoes-pendentes').hidden"),false);
    for(const check of checks)await check();
    assert.equal(await powershell('(Get-CimInstance Win32_Printer | Where-Object Default).Name'),defaultBefore);
    assert.deepEqual(exceptions,[]);
    await writeFile(path.join(folder,'result.json'),JSON.stringify({passed:true,phase,previousVersion:registration.DisplayVersion,layouts,physicalSendingTested:false,preferencePreserved:true,defaultsPreserved:true,defaultPrinter:defaultBefore,scaling125Emulated:true,homeImportTested:!before,panelsAndZoomTested:!before,dirtyAndReentryTested:!before},null,2));
    console.log('PASSOU: '+phase+', capturas Início/PDF/Konica/Epson/A4-A3-120 e três tamanhos. Nenhum envio físico.');
}catch(error){console.error(error);console.error(exceptions);try{console.error(await evaluate("document.body.innerText"));await image('failure');}catch{}process.exitCode=1;}
finally{for(const watcher of watchers){watcher.stdin.end('quit\n');watcher.kill();}socket?.close();if(app&&app.exitCode===null)app.kill();}
