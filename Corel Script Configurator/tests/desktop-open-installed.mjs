// Aplicativo NSIS instalado, registro real e ShellExecuteEx; sem mocks no pacote.
import assert from 'node:assert/strict';
import {spawn,execFile} from 'node:child_process';
import {promisify} from 'node:util';
import {mkdtemp,readFile,writeFile,unlink,realpath} from 'node:fs/promises';
import path from 'node:path';
import {tmpdir} from 'node:os';
import {setTimeout as pause} from 'node:timers/promises';
import {PDFDocument} from '../public/vendor/pdf-lib/pdf-lib.js';

const exec=promisify(execFile),folder=await mkdtemp(path.join(tmpdir(),'4printshops-open-installed-')),installation=path.join(folder,'installed');
const exe=path.join(installation,'4PrintShops PDF Test.exe'),profile=path.join(process.env.APPDATA,'4PrintShops PDF Test');
const setup=path.resolve('dist-desktop/pdf-open-test/4PrintShops-1.0.0-experimental.4-Windows-x64-Setup.exe');
let child,socket,id=0;const pending=new Map(),exceptions=[];
const wait=async(fn,label,timeout=40000)=>{const end=Date.now()+timeout;while(Date.now()<end){const v=await fn();if(v)return v;await pause(60);}throw new Error('Tempo esgotado: '+label);};
const cdp=(method,params={})=>new Promise((resolve,reject)=>{const n=++id,timer=setTimeout(()=>{pending.delete(n);reject(new Error('CDP: '+method));},10000);pending.set(n,{resolve:v=>{clearTimeout(timer);resolve(v);},reject});socket.send(JSON.stringify({id:n,method,params}));});
const evaluate=async expression=>{const r=await cdp('Runtime.evaluate',{expression,awaitPromise:true,returnByValue:true,userGesture:true});if(r.exceptionDetails)throw new Error(r.exceptionDetails.exception?.description||r.exceptionDetails.text);return r.result.value;};
const click=selector=>evaluate(`document.querySelector(${JSON.stringify(selector)}).click()`);
const quote=value=>"'"+value.replaceAll("'","''")+"'";
const ps=async script=>(await exec('powershell.exe',['-NoProfile','-Command',script],{windowsHide:true,timeout:25000})).stdout.trim();
const windowAction=async minimize=>ps(`Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public static class OpenTestWindow{[DllImport("user32.dll")]public static extern bool ShowWindow(IntPtr h,int n);[DllImport("user32.dll")]public static extern bool IsIconic(IntPtr h);}';$h=(Get-Process -Id ${child.pid}).MainWindowHandle;${minimize?'[OpenTestWindow]::ShowWindow($h,6)|Out-Null;':''}[OpenTestWindow]::IsIconic($h)`);
const snapshot=()=>ps(`$r=[Microsoft.Win32.Registry]::CurrentUser;$out=@{};foreach($k in @('Software\\Classes\\.pdf','Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\FileExts\\.pdf\\UserChoice')){$s=$r.OpenSubKey($k);if($s -and $s.GetValueNames().Length -gt 0){$out[$k]=@{};foreach($n in $s.GetValueNames()){$out[$k][$n]=$s.GetValue($n)};$s.Close()}};$out|ConvertTo-Json -Depth 4 -Compress`);
const ownRegistry=()=>ps(`$r=[Microsoft.Win32.Registry]::CurrentUser;$k=$r.OpenSubKey('Software\\Classes\\br.com.4printshops.opentest.PDF\\shell\\open\\command');if($k){$k.GetValue('');$k.Close()}`);
const shellOpen=async file=>{const source=await readFile(new URL('./helpers/windows-open-pdf.ps1',import.meta.url),'utf8');return ps(`& {${source}} -PdfPath ${quote(file)}`);};
async function fixture(name,pages){const p=await PDFDocument.create();for(let n=0;n<pages;n++)p.addPage([300+n*20,500]);const bytes=Buffer.from(await p.save()),file=path.join(folder,name);await writeFile(file,bytes);return {file,bytes};}
async function launch(file){
    await unlink(path.join(profile,'DevToolsActivePort')).catch(()=>{});
    const env={...process.env};delete env.ELECTRON_RUN_AS_NODE;
    child=spawn(exe,['--remote-debugging-port=0','--open-pdf',file],{env,windowsHide:false,stdio:'ignore'});
    const port=await wait(async()=>{try{return(await readFile(path.join(profile,'DevToolsActivePort'),'utf8')).split('\n')[0];}catch{return false;}},'Startup instalado');
    const target=await wait(async()=>{try{return(await(await fetch(`http://127.0.0.1:${port}/json`)).json()).find(t=>t.type==='page');}catch{return false;}},'Janela');
    socket=new WebSocket(target.webSocketDebuggerUrl);await new Promise(resolve=>socket.addEventListener('open',resolve,{once:true}));
    socket.addEventListener('message',event=>{const m=JSON.parse(event.data);if(m.id){if(m.error)pending.get(m.id)?.reject(new Error(m.error.message));else pending.get(m.id)?.resolve(m.result);pending.delete(m.id);}if(m.method==='Runtime.exceptionThrown')exceptions.push(m.params.exceptionDetails.text);});
    await cdp('Runtime.enable');
}
const ready=()=>wait(()=>evaluate("!!document.querySelector('#pdf-arquivo')&&!document.querySelector('#pdf-arquivo').disabled&&!document.querySelector('#app-shell').inert"),'Operação concluída');
const dialog=()=>wait(()=>evaluate("!!document.querySelector('#desktop-open-dialog')?.open"),'Decisão de abertura');
const screenshot=async name=>{const r=await cdp('Page.captureScreenshot');await writeFile(path.join(folder,name+'.png'),Buffer.from(r.data,'base64'));};
const before=await snapshot();
try{
    assert.ok((await realpath(folder)).startsWith(await realpath(tmpdir())));
    await new Promise((resolve,reject)=>{const install=spawn(setup,['/S','/D='+installation],{windowsHide:true,stdio:'ignore'});install.once('error',reject);install.once('exit',code=>code===0?resolve():reject(new Error('NSIS: '+code)));});
    assert.deepEqual(JSON.parse(await snapshot()),JSON.parse(before),'Instalação não muda default nem UserChoice');
    const registered=/^"(.+)" --open-pdf "%1"$/.exec(await ownRegistry());assert.ok(registered);
    assert.equal((await realpath(registered[1])).toLowerCase(),(await realpath(exe)).toLowerCase());
    const a=await fixture('Proposta São José.pdf',2),b=await fixture('Segundo arquivo.pdf',1),c=await fixture('Terceiro arquivo.pdf',3);
    await launch(a.file);await wait(()=>evaluate("document.querySelector('#pdf-nome')?.textContent==='Proposta São José.pdf'"),'PDF do startup');await ready();
    assert.equal(await evaluate("document.querySelector('#app-shell').dataset.activeTool"),'pdfs');
    await wait(()=>evaluate("!!document.querySelector('#pdf-visualizacao canvas')"),'Canvas PDF.js');
    await screenshot('01-startup');
    // Segundo processo via Shell Windows e ProgID realmente instalado.
    await click('#pdf-girar-direita');await ready();await click('[data-tool=inicio]');
    assert.equal(await windowAction(true),'True');
    await shellOpen(b.file);await dialog();await click('#desktop-open-open');
    assert.equal(await windowAction(false),'False','Segunda abertura restaura a janela minimizada');
    await wait(()=>evaluate("document.querySelector('#pdf-confirmar-descarte').open"),'Proteção de alterações');
    await click('#pdf-continuar-editando');await ready();
    assert.equal(await evaluate("document.querySelector('#pdf-nome').textContent"),'Proposta São José.pdf');
    assert.equal(await evaluate("document.querySelector('#pdf-alteracoes-pendentes').hidden"),false);
    console.log('PASSOU: NSIS/registro sem troca de default; startup antes de renderer pronto, espaços/acentos, segunda instância e cancelamento preservando PDF alterado.');
    await shellOpen(b.file);await dialog();await click('#desktop-open-add');await ready();
    await wait(()=>evaluate("document.querySelector('#pdf-pagina-atual').textContent.endsWith('/ 3')"),'Adição sem perder giro');
    // Rajada equivalente à invocação por arquivo feita pelo shell para seleção múltipla.
    await Promise.all([shellOpen(b.file),shellOpen(c.file)]);await dialog();await click('#desktop-open-add');await ready();
    // Dependendo do agendamento do shell, a segunda solicitação pode chegar durante a decisão.
    await wait(async()=>{if(await evaluate("document.querySelector('#desktop-open-dialog').open")){await click('#desktop-open-add');await ready();}return evaluate("document.querySelector('#pdf-pagina-atual').textContent.endsWith('/ 7')");},'Todos os PDFs da rajada');
    await screenshot('02-multiplos');
    // Exportação nativa: deve conservar o giro original e a sequência adicionada.
    await click('#pdf-salvar-copia');await click('#app-exportacao-confirmar');
    const saveScript=await readFile(new URL('./helpers/windows-save-dialog.ps1',import.meta.url),'utf8'),destination=path.join(folder,'resultado.pdf');
    await ps(`& {${saveScript}} -DialogProcess ${child.pid} -Destination ${quote(destination)} -Action save`);await ready();
    const result=await PDFDocument.load(await readFile(destination));assert.equal(result.getPageCount(),7);assert.equal(result.getPage(0).getRotation().angle,90);
    // Caminho inexistente e conteúdo inválido retornam controles responsivos.
    const missing=path.join(folder,'ausente.pdf');const second=spawn(exe,['--open-pdf',missing],{windowsHide:true,stdio:'ignore'});await new Promise(resolve=>second.once('exit',resolve));
    await dialog();await click('#desktop-open-open');await wait(()=>evaluate("!!document.querySelector('#desktop-open-retry')"),'Erro legível');await click('#desktop-open-cancel');
    const bad=path.join(folder,'inválido.pdf');await writeFile(bad,'Não é PDF');await shellOpen(bad);await dialog();await click('#desktop-open-open');await wait(()=>evaluate("!!document.querySelector('#desktop-open-retry')"),'Cabeçalho inválido');await click('#desktop-open-cancel');
    assert.equal(await evaluate("document.querySelector('#pdf-nome').textContent"),'Proposta São José.pdf');
    for(const item of [a,b,c])assert.deepEqual(await readFile(item.file),item.bytes);
    await evaluate("localStorage.setItem('4printshops.open-test','preservado')");await evaluate('window.close()');await wait(()=>child.exitCode!==null,'Fechar');socket.close();
    assert.deepEqual(exceptions,[]);
    const uninstall=path.join(installation,'Uninstall 4PrintShops PDF Test.exe');
    await new Promise((resolve,reject)=>{const p=spawn(uninstall,['/S'],{windowsHide:true,stdio:'ignore'});p.once('error',reject);p.once('exit',code=>code===0?resolve():reject(new Error('Uninstall: '+code)));});
    await wait(async()=>!(await ownRegistry()),'Remoção da associação');assert.deepEqual(JSON.parse(await snapshot()),JSON.parse(before),'Desinstalação preserva default/UserChoice');
    console.log('PASSOU: ShellExecuteEx real para vários PDFs, união preservando giro, exportação de 7 páginas, erros sem perda, originais intactos e remoção de associação sem afetar default.');
    await writeFile(path.join(folder,'result.json'),JSON.stringify({passed:true,version:'1.0.0-experimental.4',registryBefore:JSON.parse(before),registryAfter:JSON.parse(await snapshot()),installation},null,2));
}catch(error){try{await screenshot('falha');console.error('Estado:',await evaluate("({nome:document.querySelector('#pdf-nome')?.textContent,status:document.querySelector('#pdf-status')?.textContent,dialog:document.querySelector('#desktop-open-dialog')?.textContent,inert:document.querySelector('#app-shell')?.inert})"));}catch{}console.error(error);process.exitCode=1;}
finally{socket?.close();if(child&&child.exitCode===null)child.kill();console.log('Artefatos: '+folder);}
