// NSIS e auxiliar reais. Automatiza somente consultas/propriedades, sem envio de trabalhos.
import assert from 'node:assert/strict';
import {spawn,execFile} from 'node:child_process';
import {promisify} from 'node:util';
import {readFile,writeFile,mkdir,unlink} from 'node:fs/promises';
import {setTimeout as pause} from 'node:timers/promises';
import path from 'node:path';
import {PDFDocument} from '../public/vendor/pdf-lib/pdf-lib.js';

const exec=promisify(execFile),folder=path.resolve('.desktop-test/print-config/installed-validation'),installation=path.join(folder,'installed');
await mkdir(folder,{recursive:true});
const setup=path.resolve('dist-desktop/print-config-test/4PrintShops-Print-Config-Test-1.0.0-experimental.8-Windows-x64-Setup.exe');
const exe=path.join(installation,'4PrintShops Print Config Test.exe'),profile=path.join(process.env.APPDATA,'4PrintShops Print Config Test');
const powershell=path.join(process.env.SystemRoot,'System32/WindowsPowerShell/v1.0/powershell.exe');
let app,socket,serial=0;const pending=new Map(),exceptions=[],watchers=[];
const wait=async(fn,label,timeout=30000)=>{const until=Date.now()+timeout;while(Date.now()<until){const value=await fn();if(value)return value;await pause(80);}throw Error('Tempo esgotado: '+label);};
const cdp=(method,params={})=>new Promise((resolve,reject)=>{const id=++serial,timer=setTimeout(()=>{pending.delete(id);reject(Error('CDP: '+method));},30000);pending.set(id,{resolve:value=>{clearTimeout(timer);resolve(value);},reject:error=>{clearTimeout(timer);reject(error);}});socket.send(JSON.stringify({id,method,params}));});
const evaluate=async expression=>{const result=await cdp('Runtime.evaluate',{expression,awaitPromise:true,returnByValue:true,userGesture:true});if(result.exceptionDetails)throw Error(result.exceptionDetails.exception?.description||result.exceptionDetails.text);return result.result.value;};
const click=selector=>evaluate(`document.querySelector(${JSON.stringify(selector)}).click()`);
const quote=value=>"'"+value.replaceAll("'","''")+"'";
const ps=async command=>(await exec(powershell,['-NoProfile','-NonInteractive','-Command',command],{windowsHide:true,timeout:30000})).stdout.trim();
const screenshot=async name=>{const image=await cdp('Page.captureScreenshot');await writeFile(path.join(folder,name+'.png'),Buffer.from(image.data,'base64'));};
const ready=()=>wait(()=>evaluate("!document.querySelector('#print-config-close').disabled"),'Configuração disponível');
const status=()=>evaluate("document.querySelector('#print-config-status').textContent");
async function watcher(printer){
    const source=await readFile(new URL('./helpers/windows-print-defaults.ps1',import.meta.url),'utf8');
    const child=spawn(powershell,['-NoProfile','-NonInteractive','-Command',`& {${source}} -PrinterName ${quote(printer)}`],{windowsHide:true,stdio:['pipe','pipe','pipe']});watchers.push(child);
    let output='',errors='';child.stdout.on('data',chunk=>output+=chunk);child.stderr.on('data',chunk=>errors+=chunk);
    await wait(()=>{if(errors)throw Error('Monitor somente de leitura indisponível.');return output.includes('READY');},'Monitor de preferências');
    return async()=>{output='';child.stdin.write('check\n');await wait(()=>output.includes('CHANGED'),'Conferência de preferências');assert.equal(output.trim(),'UNCHANGED','Preferências globais e do usuário preservadas');};
}
async function native(action){
    const pid=Number(await ps(`$p=Get-CimInstance Win32_Process -Filter "Name='4PrintShops.PrintConfig.exe'" | Where-Object {$_.ParentProcessId -eq ${app.pid}};if(@($p).Count -ne 1){throw 'Auxiliar isolado não identificado.'};$p.ProcessId`));
    assert.ok(pid>0);
    const source=await readFile(new URL('./helpers/windows-print-properties.ps1',import.meta.url),'utf8');
    return (await exec(powershell,['-NoProfile','-NonInteractive','-Command',`& {${source}} -DriverProcess ${pid} -Action ${quote(action)}`],{windowsHide:true,timeout:30000})).stdout.trim();
}
const choose=async name=>{await evaluate(`{const select=document.querySelector('#print-config-printer');select.value=${JSON.stringify(name)};select.dispatchEvent(new Event('change'));}`);await ready();assert.match(await status(),/consultada/);};
const openProperties=async()=>{await click('#print-config-properties');await wait(()=>evaluate("document.querySelector('#print-config-properties').disabled"),'Abertura de propriedades');};
const defaultReader=()=>ps("$k=[Microsoft.Win32.Registry]::CurrentUser.OpenSubKey('Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\FileExts\\.pdf\\UserChoice');if($k){$o=@{};foreach($n in $k.GetValueNames()){$o[$n]=$k.GetValue($n)};$o|ConvertTo-Json -Compress;$k.Close()}else{'{}'}");
const readerBefore=await defaultReader();
try{
    await new Promise((resolve,reject)=>{const installer=spawn(setup,['/S','/D='+installation],{windowsHide:true,stdio:'ignore'});installer.once('error',reject);installer.once('exit',code=>code===0?resolve():reject(Error('NSIS '+code)));});
    assert.equal(await defaultReader(),readerBefore);
    const pdf=await PDFDocument.create();for(let n=0;n<2;n++){const page=pdf.addPage([595.2756,841.8898]);page.drawText('4PrintShops - Teste sintetico sem impressao - '+(n+1),{x:35,y:750,size:14});}
    const document=path.join(folder,'Teste sintético A4.pdf');await writeFile(document,await pdf.save());
    await unlink(path.join(profile,'DevToolsActivePort')).catch(()=>{});
    const env={...process.env};delete env.ELECTRON_RUN_AS_NODE;
    app=spawn(exe,['--remote-debugging-port=0','--open-pdf',document],{env,windowsHide:true,stdio:'ignore'});
    const port=await wait(async()=>{try{return (await readFile(path.join(profile,'DevToolsActivePort'),'utf8')).split('\n')[0];}catch{return false;}},'Inicialização instalada');
    const target=await wait(async()=>{try{return (await (await fetch(`http://127.0.0.1:${port}/json`)).json()).find(t=>t.type==='page');}catch{return false;}},'Janela instalada');
    socket=new WebSocket(target.webSocketDebuggerUrl);await new Promise(resolve=>socket.addEventListener('open',resolve,{once:true}));
    socket.addEventListener('message',event=>{const message=JSON.parse(event.data);if(message.id){const operation=pending.get(message.id);if(message.error)operation?.reject(Error(message.error.message));else operation?.resolve(message.result);pending.delete(message.id);}if(message.method==='Runtime.exceptionThrown')exceptions.push(message.params.exceptionDetails.text);});
    await cdp('Runtime.enable');await wait(()=>evaluate("!!document.querySelector('#pdf-visualizacao canvas')&&!document.querySelector('#pdf-arquivo').disabled"),'PDF.js instalado');
    await click('#pdf-girar-direita');await wait(()=>evaluate("!document.querySelector('#pdf-arquivo').disabled&&!document.querySelector('#pdf-alteracoes-pendentes').hidden"),'PDF modificado');
    await click('#app-menu-arquivo');await click('[data-command="configurar-impressao"]');await ready();
    const names=await evaluate("[...document.querySelector('#print-config-printer').options].map(o=>o.value)");
    assert.ok(names.includes('KONICA MINOLTA C6100PS'));assert.ok(names.includes('Microsoft Print to PDF'));
    const checkKonica=await watcher('KONICA MINOLTA C6100PS');
    const checkVirtual=await watcher('Microsoft Print to PDF');
    await choose('KONICA MINOLTA C6100PS');const before=await evaluate("document.querySelector('#print-config-summary').textContent");
    await openProperties();console.log('Diálogo nativo Konica:',await native('inspect'));await native('cancel');await ready();
    assert.match(await status(),/canceladas/);assert.equal(await evaluate("document.querySelector('#print-config-summary').textContent"),before);await checkKonica();
    await openProperties();await native('landscape');await native('accept');await ready();assert.match(await status(),/confirmada/);await checkKonica();
    assert.match(await evaluate("document.querySelector('#print-config-summary').textContent"),/Paisagem/,'Mudança confirmada retornou ao resumo');
    await click('#print-config-close');await click('#app-menu-arquivo');await click('[data-command="configurar-impressao"]');await ready();await choose('KONICA MINOLTA C6100PS');
    assert.equal(await evaluate("document.querySelector('#print-config-summary').textContent"),before,'Encerrar a sessão descarta ajustes locais e mantém defaults');
    await evaluate("document.querySelector('#print-config-scale').value='120';document.querySelector('#print-config-scale').dispatchEvent(new Event('change'))");
    assert.match(await evaluate("document.querySelector('#print-config-measures').textContent"),/252 × 356,4/);await screenshot('konica-previa-a3-120');
    await evaluate("document.querySelector('#print-config-paper').value='driver';document.querySelector('#print-config-paper').dispatchEvent(new Event('change'))");
    assert.match(await evaluate("document.querySelector('#print-config-warning').textContent"),/ultrapassa o papel/);await screenshot('konica-corte-a4-120');
    await choose('Microsoft Print to PDF');await openProperties();console.log('Diálogo nativo virtual:',await native('inspect'));
    let changed=false;try{await native('a3');changed=true;}catch{console.log('Papel A3 não automatizável nesse diálogo; confirmação de mudança dependerá de validação manual.');}
    await native('accept');await ready();assert.match(await status(),/confirmada/);await checkVirtual();
    if(changed)assert.match(await evaluate("document.querySelector('#print-config-summary').textContent"),/A3/);
    await openProperties();await native('landscape');await native('accept');await ready();await checkVirtual();assert.match(await evaluate("document.querySelector('#print-config-summary').textContent"),/Paisagem/);
    await openProperties();await native('cancel');await ready();await checkVirtual();
    await click('#print-config-close');await wait(()=>evaluate("!document.querySelector('#print-config-dialog').open"),'Fechamento');
    assert.equal(await evaluate("document.querySelector('#pdf-alteracoes-pendentes').hidden"),false,'Configurar não quita alterações');
    await click('#app-menu-arquivo');await click('[data-command="configurar-impressao"]');await ready();assert.equal(await evaluate("document.querySelectorAll('#print-config-dialog').length"),1);await choose('KONICA MINOLTA C6100PS');
    assert.equal(await evaluate("document.querySelector('#print-config-summary').textContent"),before,'Reabrir inicia dos defaults intactos');
    await ps(`$p=Get-CimInstance Win32_Process -Filter "Name='4PrintShops.PrintConfig.exe'" | Where-Object {$_.ParentProcessId -eq ${app.pid}};if(@($p).Count -ne 1){throw 'Auxiliar isolado não identificado.'};Stop-Process -Id $p.ProcessId`);
    await pause(300);await openProperties();await ready();assert.match(await status(),/Sessão desatualizada|encerrado/);await choose('KONICA MINOLTA C6100PS');await checkKonica();await click('#print-config-close');await ready();
    await click('#pdf-salvar-copia');await click('#app-exportacao-confirmar');
    const saveSource=await readFile(new URL('./helpers/windows-save-dialog.ps1',import.meta.url),'utf8'),destination=path.join(folder,'resultado-sintético.pdf');
    await ps(`& {${saveSource}} -DialogProcess ${app.pid} -Destination ${quote(destination)} -Action save`);
    await wait(()=>evaluate("document.querySelector('#pdf-alteracoes-pendentes').hidden"),'Exportação nativa');
    const exported=await PDFDocument.load(await readFile(destination));assert.equal(exported.getPageCount(),2);assert.equal(exported.getPage(0).getRotation().angle,90);
    assert.deepEqual(exceptions,[]);await evaluate("localStorage.setItem('4printshops.print-config-check','preservado')");
    await writeFile(path.join(folder,'result.json'),JSON.stringify({passed:true,setup,installation,exe,version:'1.0.0-experimental.8',konicaCanceled:true,konicaConfirmed:true,konicaOrientationChanged:true,virtualOrientationChanged:true,virtualA3Changed:changed,globalAndUserDefaultsPreserved:true,dirtyPDFPreserved:true,helperFailureRecovered:true,exportedPages:2,exportedRotation:90,noJobsSubmitted:true},null,2));
    console.log('PASSOU: instalação real, abertura Windows/PDF.js, menu, propriedades Konica cancelar/confirmar, virtual, escalas/corte, reabertura sem duplicação, PDF alterado e defaults intactos.');
}catch(error){try{await screenshot('falha');console.error('Estado público:',await status());}catch{}console.error(error);process.exitCode=1;}
finally{for(const child of watchers){child.stdin.end('quit\n');child.kill();}socket?.close();if(app&&app.exitCode===null)app.kill();console.log('Instalação: '+installation);console.log('Artefatos: '+folder);}
