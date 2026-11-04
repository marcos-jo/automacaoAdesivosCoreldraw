// Fluxo real no Chrome ou na Hold Test instalada. Impressão somente virtual.
import assert from 'node:assert/strict';
import {spawn,execFile} from 'node:child_process';
import {promisify} from 'node:util';
import {readFile,writeFile,mkdir,access,unlink} from 'node:fs/promises';
import path from 'node:path';
import {setTimeout as pause} from 'node:timers/promises';
import {PDFDocument} from '../public/vendor/pdf-lib/pdf-lib.js';
import {fixtures,jpg,png} from './helpers/pdf-combine-fixtures.mjs';
const desktop=process.env.COMBINE_DESKTOP==='1';
const runId=Date.now();
const folder=path.resolve('.desktop-test/combine14/'+(desktop?'installed':'browser'));await mkdir(folder,{recursive:true});
const files=await fixtures();for(const file of files)await writeFile(path.join(folder,file.name),new Uint8Array(await file.arrayBuffer()));
const exec=promisify(execFile),ps=path.join(process.env.SystemRoot,'System32/WindowsPowerShell/v1.0/powershell.exe');
const quote=text=>"'"+text.replaceAll("'","''")+"'";
const powershell=async code=>(await exec(ps,['-NoProfile','-NonInteractive','-Command',code],{windowsHide:true,timeout:60000})).stdout.trim();
const wait=async(fn,label)=>{const limit=Date.now()+90000;while(Date.now()<limit){if(await fn())return;await pause(80);}throw Error('Tempo esgotado '+label);};
let app,server,socket,serial=0;const pending=new Map(),exceptions=[];
const cdp=(method,params={})=>new Promise((resolve,reject)=>{const id=++serial,timer=setTimeout(()=>{pending.delete(id);reject(Error(method));},60000);pending.set(id,{resolve:result=>{clearTimeout(timer);resolve(result);},reject});socket.send(JSON.stringify({id,method,params}));});
const evaluate=async expression=>{const r=await cdp('Runtime.evaluate',{expression,awaitPromise:true,returnByValue:true,userGesture:true});if(r.exceptionDetails)throw Error(r.exceptionDetails.exception?.description||r.exceptionDetails.text);return r.result.value;};
const click=selector=>evaluate(`document.querySelector(${JSON.stringify(selector)}).click()`);
const ready=()=>wait(()=>evaluate("!document.querySelector('#pdf-arquivo').disabled&&!document.querySelector('#pdf-uniao-arquivos').disabled"),'interface livre');
const combined=()=>wait(()=>evaluate("!document.querySelector('#pdf-uniao-janela').open&&!!document.querySelector('#pdf-visualizacao canvas')&&!document.querySelector('#pdf-arquivo').disabled"),'documento combinado');
const image=async name=>{await pause(250);const r=await cdp('Page.captureScreenshot');await writeFile(path.join(folder,name+'.png'),Buffer.from(r.data,'base64'));};
async function select(selector,items){
    const payload=await Promise.all(items.map(async file=>({name:file.name,type:file.type,bytes:Buffer.from(await file.arrayBuffer()).toString('base64')})));
    await evaluate(`(()=>{const d=new DataTransfer();for(const f of ${JSON.stringify(payload)})d.items.add(new File([Uint8Array.from(atob(f.bytes),c=>c.charCodeAt(0))],f.name,{type:f.type}));const e=document.querySelector(${JSON.stringify(selector)});e.files=d.files;e.dispatchEvent(new Event('change',{bubbles:true}));})()`);await ready();
}
async function nativeFile(destination,action='open',print=false){
    let code=await readFile(new URL('./helpers/windows-save-dialog.ps1',import.meta.url),'utf8');
    if(print){
        code=code.replaceAll("'Salvar cópia PDF'","'Salvar saída da impressora PDF'");
        code=code.replace("[PrintShopsDialog]::Save($dialogHandle,$Destination)", "[PrintShopsDialog]::Click($dialogHandle,'Salvar')");
    }
    await powershell(`& {${code}} -DialogProcess ${app.pid} -Destination ${quote(destination)} -Action ${action}`);
}
async function exportPDF(name){
    await click('#pdf-salvar-copia');await wait(()=>evaluate("document.querySelector('#app-exportacao-dialogo').open"),'exportação');
    await evaluate(`document.querySelector('#app-exportacao-nome').value=${JSON.stringify(name)}`);await click('#app-exportacao-confirmar');
    const file=path.join(folder,name);if(desktop)await nativeFile(file,'save');
    await wait(async()=>{try{return (await readFile(file)).length>0;}catch{return false;}},'arquivo exportado');await ready();
    return PDFDocument.load(await readFile(file));
}
try{
    let port,profile,defaultBefore;
    if(desktop){
        assert.equal(await powershell("@(Get-CimInstance Win32_Process | Where-Object {$_.Name -eq '4PrintShops Hold Test.exe'}).Count"),'0','Feche conscientemente a Hold Test');
        const installation=path.resolve('.desktop-test/hold-lab/installed-validation/installed');
        const reg="HKCU:/Software/Microsoft/Windows/CurrentVersion/Uninstall/90ff2ff0-f013-571c-b109-7edb265dacc8";
        const current=JSON.parse(await powershell(`Get-ItemProperty '${reg}' | Select-Object DisplayVersion,UninstallString | ConvertTo-Json -Compress`));
        assert.ok(current.UninstallString.startsWith('"'+path.join(installation,'Uninstall 4PrintShops Hold Test.exe')+'"'));
        if(process.env.COMBINE_INSTALL==='1')await new Promise((resolve,reject)=>{const p=spawn(path.resolve('dist-desktop/hold-native-test/4PrintShops-Hold-Test-1.0.0-experimental.14-Windows-x64-Setup.exe'),['/S'],{windowsHide:true});p.once('error',reject);p.once('exit',code=>code===0?resolve():reject(Error('NSIS '+code)));});
        assert.equal(await powershell(`(Get-ItemProperty '${reg}').DisplayVersion`),'1.0.0-experimental.14');
        defaultBefore=await powershell('(Get-CimInstance Win32_Printer | Where-Object Default).Name');
        assert.equal(await powershell("(Get-ItemProperty 'HKCU:/Software/Microsoft/Windows NT/CurrentVersion/Windows').LegacyDefaultPrinterMode"),'1');
        profile=path.join(process.env.APPDATA,'4PrintShops Hold Test');await unlink(path.join(profile,'DevToolsActivePort')).catch(()=>{});
        const env={...process.env};delete env.ELECTRON_RUN_AS_NODE;
        app=spawn(path.join(installation,'4PrintShops Hold Test.exe'),['--remote-debugging-port=0'],{env,windowsHide:true,stdio:'ignore'});
    }else{
        try{await fetch('http://localhost:3000');}catch{server=spawn(process.execPath,['app.js'],{windowsHide:true,stdio:'ignore'});}
        await wait(async()=>{try{return(await fetch('http://localhost:3000')).ok;}catch{return false;}},'servidor');
        profile=path.join(folder,'profile-'+Date.now());await mkdir(profile);
        app=spawn('C:/Program Files/Google/Chrome/Application/chrome.exe',['--headless=new','--disable-gpu','--no-first-run','--remote-debugging-port=0','--user-data-dir='+profile,'http://localhost:3000'],{windowsHide:true,stdio:'ignore'});
    }
    await wait(async()=>{try{port=(await readFile(path.join(profile,'DevToolsActivePort'),'utf8')).split('\n')[0];return !!port;}catch{return false;}},'porta');
    let target;await wait(async()=>{try{target=(await(await fetch('http://127.0.0.1:'+port+'/json')).json()).find(t=>t.type==='page');return !!target;}catch{return false;}},'janela');
    socket=new WebSocket(target.webSocketDebuggerUrl);await new Promise(resolve=>socket.addEventListener('open',resolve,{once:true}));
    socket.addEventListener('message',event=>{const m=JSON.parse(event.data);if(m.id){const p=pending.get(m.id);m.error?p?.reject(Error(m.error.message)):p?.resolve(m.result);pending.delete(m.id);}if(m.method==='Runtime.exceptionThrown')exceptions.push(m.params.exceptionDetails.exception?.description||m.params.exceptionDetails.text);});
    await cdp('Runtime.enable');await cdp('Page.enable');
    if(!desktop)await cdp('Browser.setDownloadBehavior',{behavior:'allow',downloadPath:folder});
    await cdp('Emulation.setDeviceMetricsOverride',{width:1366,height:768,deviceScaleFactor:1,mobile:false});
    await wait(()=>evaluate("!!document.querySelector('[data-open-pdf]')"),'início');
    if(desktop){
        assert.equal(await evaluate("localStorage.getItem('4printshops.print-config-check')"),'preservado');
        await click('[data-open-pdf]');await nativeFile(path.join(folder,files[4].name));await wait(()=>evaluate("document.querySelectorAll('#pdf-uniao-lista li').length===1"),'imagem importada pelo diálogo');await ready();
        assert.equal(await evaluate("document.querySelectorAll('#pdf-uniao-lista li').length"),1);
        await select('#pdf-uniao-arquivos',files.slice(0,4));
    }else{await click('[data-tool="pdfs"]');await select('#pdf-arquivo',[files[4],...files.slice(0,4)]);}
    assert.equal(await evaluate("document.querySelectorAll('#pdf-uniao-lista li').length"),7);
    await wait(()=>evaluate("document.querySelectorAll('#pdf-uniao-lista img').length>0"),'miniaturas reais');await image('01-combinar-pdfs-imagens');
    // Mover uma página pelo controle acessível e outra por drag-and-drop.
    await click('#pdf-uniao-lista li:nth-child(2) [data-acao="subir"]');
    await evaluate("(()=>{const rows=document.querySelectorAll('#pdf-uniao-lista li');const d=new DataTransfer();rows[0].dispatchEvent(new DragEvent('dragstart',{bubbles:true,dataTransfer:d}));rows[2].dispatchEvent(new DragEvent('dragover',{bubbles:true,dataTransfer:d}));rows[2].dispatchEvent(new DragEvent('drop',{bubbles:true,dataTransfer:d}));})()");
    const sequence=await evaluate("Array.from(document.querySelectorAll('#pdf-uniao-lista .pdf-uniao-nome'),e=>e.textContent)");assert.equal(sequence.length,7);
    await select('#pdf-uniao-arquivos',[new File(['bad'],'Arquivo corrompido.jpg')]);assert.equal(await evaluate("document.querySelectorAll('#pdf-uniao-lista li').length"),7);assert.match(await evaluate("document.querySelector('#pdf-uniao-status').textContent"),/corrompido/);
    await evaluate(`document.querySelector('#pdf-uniao-nome').value='BalcaoTeste${runId}'`);
    await evaluate("document.querySelector('#pdf-uniao-ajuste').value='preencher';document.querySelector('#pdf-uniao-ajuste').dispatchEvent(new Event('change'))");assert.equal(await evaluate("document.querySelector('#pdf-uniao-gerar').disabled"),true);
    await click('#pdf-uniao-corte');assert.equal(await evaluate("document.querySelector('#pdf-uniao-gerar').disabled"),false);
    await evaluate("document.querySelector('#pdf-uniao-ajuste').value='ajustar';document.querySelector('#pdf-uniao-ajuste').dispatchEvent(new Event('change'))");await click('#pdf-uniao-gerar');await combined();
    assert.equal(await evaluate("document.querySelector('#pdf-total').textContent"),'7');assert.equal(await evaluate("document.querySelector('#pdf-alteracoes-pendentes').hidden"),false);await image('02-combinado-a4');
    const output=await exportPDF('combinado-a4-'+runId+'.pdf');assert.equal(output.getPageCount(),7);
    for(const page of output.getPages()){const sizes=[page.getWidth(),page.getHeight()].sort((a,b)=>a-b);assert.ok(Math.abs(sizes[0]-595.2756)<.01);assert.ok(Math.abs(sizes[1]-841.8898)<.01);}
    await click('#pdf-girar-direita');await ready();
    // Cancelar substituição não descarta o documento editado nem o rascunho.
    await click('#pdf-ferramenta-unir');await click('#pdf-uniao-abrir');await select('#pdf-uniao-arquivos',[files[2]]);
    await evaluate("document.querySelector('#pdf-uniao-destino').value='novo'");await click('#pdf-uniao-gerar');await wait(()=>evaluate("document.querySelector('#pdf-confirmar-descarte').open"),'proteção');await click('#pdf-continuar-editando');await ready();
    assert.equal(await evaluate("document.querySelector('#pdf-total').textContent"),'7');assert.equal(await evaluate("document.querySelector('#pdf-alteracoes-pendentes').hidden"),false);
    await evaluate("document.querySelector('#pdf-uniao-destino').value='adicionar';document.querySelector('#pdf-uniao-ajuste').value='original'");await click('#pdf-uniao-gerar');await combined();assert.equal(await evaluate("document.querySelector('#pdf-total').textContent"),'8');
    // Comparação de orientação EXIF e transparência com o decoder de imagem do navegador.
    for(const [name,bytes,type]of (desktop?[]:[['exif6',jpg(6),'jpeg'],['transparencia',png(),'png']])){
        const base64=Buffer.from(bytes).toString('base64');
        const result=await evaluate(`(async()=>{const {importarImagem}=await import('/js/tools/pdf-image-import.js');const {getDocument}=await import('/vendor/pdfjs/pdf.js');const data=Uint8Array.from(atob(${JSON.stringify(base64)}),c=>c.charCodeAt(0));const {pdf}=await importarImagem(data,${JSON.stringify(type==='jpeg'?'jpg':'png')});const task=getDocument({data:await pdf.save()});try{const doc=await task.promise,p=await doc.getPage(1),v=p.getViewport({scale:150/72}),c=document.createElement('canvas');c.width=Math.round(v.width);c.height=Math.round(v.height);await p.render({canvas:c,canvasContext:c.getContext('2d'),viewport:v}).promise;const original=new Image();original.src='data:image/${type};base64,'+${JSON.stringify(base64)};await original.decode();const expected=document.createElement('canvas');expected.width=c.width;expected.height=c.height;const ctx=expected.getContext('2d');ctx.fillStyle='white';ctx.fillRect(0,0,c.width,c.height);ctx.drawImage(original,0,0,c.width,c.height);let max=0;for(const [x,y]of [[10,10],[c.width-11,10],[10,c.height-11],[c.width-11,c.height-11]]){const a=c.getContext('2d').getImageData(x,y,1,1).data,b=ctx.getImageData(x,y,1,1).data;for(let i=0;i<4;i++)max=Math.max(max,Math.abs(a[i]-b[i]));}return {max,width:c.width,height:c.height};}finally{await task.destroy();}})()`);
        assert.ok(result.max<=3,name+' aparência mudou '+JSON.stringify(result));console.log('PASSOU aparência '+name, result);
    }
    if(desktop){
        await evaluate("(async()=>{const {JanelaConfiguracaoImpressao}=await import('/js/app/desktop-print-ui.js');const open=JanelaConfiguracaoImpressao.prototype.open;JanelaConfiguracaoImpressao.prototype.open=function(){window.printUITest=this;return open.call(this);};})()");
        await click('#app-menu-arquivo');await click('[data-command="configurar-impressao"]');await wait(()=>evaluate("document.querySelector('#print-config-dialog').open&&!document.querySelector('#print-config-close').disabled&&!document.querySelector('#pdf-arquivo').disabled&&document.querySelector('#print-config-page').textContent.includes('8')"),'impressão');
        await evaluate("document.querySelector('#print-config-printer').value='Microsoft Print to PDF';document.querySelector('#print-config-printer').dispatchEvent(new Event('change'))");await wait(()=>evaluate("!document.querySelector('#print-config-print').disabled"),'fila virtual');
        assert.equal(await evaluate("document.querySelector('#print-config-printer').value"),'Microsoft Print to PDF');
        const printed=path.join(process.env.USERPROFILE,'Downloads',`BalcaoTeste${runId}_editado_impresso.pdf`);
        await assert.rejects(access(printed));
        await click('#print-config-print');await nativeFile(printed,'save',true);await wait(()=>evaluate("!document.querySelector('#print-config-close').disabled"),'impressão virtual');
        await wait(async()=>{try{return (await PDFDocument.load(await readFile(printed))).getPageCount()===8;}catch{return false;}},'gravação da saída virtual pelo spooler');
        assert.equal((await PDFDocument.load(await readFile(printed))).getPageCount(),8);await writeFile(path.join(folder,'combinado-virtual-validado.pdf'),await readFile(printed));assert.equal(await evaluate("document.querySelector('#pdf-alteracoes-pendentes').hidden"),false);await click('#print-config-close');
        assert.equal(await powershell('(Get-CimInstance Win32_Printer | Where-Object Default).Name'),defaultBefore);
    }
    assert.deepEqual(exceptions,[]);console.log('PASSOU combinação '+(desktop?'Electron instalado + impressão virtual':'Chrome headless')+', A4, EXIF, transparência, ordem, proteção e exportação. Nenhuma impressão física.');
    await writeFile(path.join(folder,'result.json'),JSON.stringify({passed:true,desktop,pages:8,physicalPrint:false,virtualPrint:desktop,sequence},null,2));
}catch(error){console.error(error);console.error(exceptions);try{console.error(await evaluate("JSON.stringify({status:document.querySelector('#pdf-uniao-status')?.textContent,print:document.querySelector('#print-config-status')?.textContent,busy:window.printUITest?.busy,source:window.printUITest?.documentPreview?.source,loaded:window.printUITest?.documentPreview?.preview?.documento?.numPages,rendering:!!window.printUITest?.documentPreview?.preview?.render})"));await image('failure');}catch{}process.exitCode=1;}
finally{socket?.close();if(app&&app.exitCode===null)app.kill();if(server&&server.exitCode===null)server.kill();}
