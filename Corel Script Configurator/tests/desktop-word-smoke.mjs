// Importação unificada real em Electron: somente sintéticos, sem impressão.
import assert from 'node:assert/strict';
import {mkdtemp,writeFile,readFile,mkdir} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import path from 'node:path';
import {fileURLToPath,pathToFileURL} from 'node:url';
import {spawn} from 'node:child_process';
import {setTimeout as pause} from 'node:timers/promises';
import {convertWord,runOffice,safeProfile} from '../desktop/word-import.mjs';
import {createOfficeResolver} from '../desktop/libreoffice-runtime.mjs';
import {createWordFixture} from './helpers/word-fixtures.mjs';
import {png,jpg} from './helpers/pdf-combine-fixtures.mjs';
import {PDFDocument} from '../public/vendor/pdf-lib/pdf-lib.js';
const root=fileURLToPath(new URL('../',import.meta.url));
const folder=await mkdtemp(path.join(tmpdir(),'4printshops-import31-'));
const privateResolver=createOfficeResolver({projectRoot:root});
const engine=await privateResolver();assert.equal(engine.mode,'bundled','Prepare o motor privado antes deste ensaio');
const executable=engine.executable;
const word=path.join(folder,'Proposta São José & final.docx'),wordBytes=createWordFixture({font:'FonteInexistente4PrintShops',pages:9});await writeFile(word,wordBytes);
const first=await convertWord(word,{resolveExecutable:privateResolver});assert.equal(first.report.pages,2);assert.equal(first.report.sourcePages,9);await writeFile(path.join(folder,'docx.pdf'),first.bytes);
const generator=path.join(folder,'generator');await mkdir(path.join(generator,'user'),{recursive:true});await writeFile(path.join(generator,'user/registrymodifications.xcu'),safeProfile);
await runOffice(executable,['-env:UserInstallation='+pathToFileURL(generator).href,'--headless','--convert-to','doc:MS Word 97','--outdir',folder,word],{cwd:folder,env:engine.env});
const doc=word.replace(/docx$/,'doc'),docBytes=await readFile(doc);const second=await convertWord(doc,{resolveExecutable:privateResolver});assert.equal(second.report.pages,2);await writeFile(path.join(folder,'doc.pdf'),second.bytes);
const pdf=await PDFDocument.create();pdf.addPage([300,400]).drawText('PDF ORIGINAL');const pdfFile=path.join(folder,'original.pdf');await writeFile(pdfFile,await pdf.save());
const image=path.join(folder,'imagem.png'),photo=path.join(folder,'foto.JPG');await writeFile(image,png());await writeFile(photo,jpg());
const invalid=path.join(folder,'invalido.docx');await writeFile(invalid,'Este arquivo não é Word.');
const sources=[pdfFile,image,photo,word,doc],originals=await Promise.all(sources.map(file=>readFile(file)));
let child,socket,id=0,serial=0,options={},debug='';const requests=new Map(),errors=[];
const wait=async(fn,label,ms=90000)=>{const end=Date.now()+ms;while(Date.now()<end){if(child?.exitCode!=null)throw Error('Electron terminou: '+debug.slice(-1000));const value=await fn();if(value)return value;await pause(70);}throw Error('Tempo esgotado: '+label);};
const cdp=(method,params={})=>new Promise((resolve,reject)=>{const n=++id,t=setTimeout(()=>{requests.delete(n);reject(Error(method));},30000);requests.set(n,{resolve:v=>{clearTimeout(t);resolve(v);}});socket.send(JSON.stringify({id:n,method,params}));});
const evaluate=async expression=>{const r=await cdp('Runtime.evaluate',{expression,awaitPromise:true,returnByValue:true,userGesture:true});if(r.exceptionDetails)throw Error(r.exceptionDetails.exception?.description||r.exceptionDetails.text);return r.result.value;};
const click=selector=>evaluate(`document.querySelector(${JSON.stringify(selector)}).click()`);
const command=async data=>{options={...options,...data,serial:++serial};await writeFile(path.join(folder,'command.json'),JSON.stringify(options));await wait(async()=>{try{return JSON.parse(await readFile(path.join(folder,'status.json'),'utf8')).serial===serial;}catch{return false;}},'comando');};
const ready=()=>wait(()=>evaluate(`!!document.querySelector('#pdf-arquivo')&&!document.querySelector('#pdf-arquivo').disabled&&!document.querySelector('#pdf-import-progress')?.open`),'editor liberado');
const review=()=>wait(()=>evaluate(`(()=>{if(!document.querySelector('#pdf-import-progress')?.open&&document.querySelector('#pdf-status').textContent.startsWith('Não foi possível importar:'))throw Error(document.querySelector('#pdf-status').textContent);return document.querySelector('#pdf-uniao-janela').open&&!document.querySelector('#pdf-uniao-gerar').disabled;})()`),'lote preparado');
try {
    await writeFile(path.join(folder,'command.json'),JSON.stringify({serial:0}));
    const env={...process.env,PRINTSHOPS_TEST_FOLDER:folder};delete env.ELECTRON_RUN_AS_NODE;env.PRINTSHOPS_LIBREOFFICE=path.join(folder,'no-external-office/soffice.com');env.ProgramFiles=folder;env['ProgramFiles(x86)']=folder;env.PATH=folder;
    child=spawn(path.join(root,'node_modules/electron/dist/electron.exe'),['tests/helpers/desktop-entry.mjs','--remote-debugging-port=0'],{cwd:root,env,windowsHide:true,stdio:['ignore','pipe','pipe']});
    child.stderr.on('data',d=>debug+=d);child.stdout.on('data',d=>debug+=d);
    const port=await wait(async()=>{try{return (await readFile(path.join(folder,'profile/DevToolsActivePort'),'utf8')).split('\n')[0];}catch{return false;}},'porta');
    const target=await wait(async()=>{try{return (await(await fetch(`http://127.0.0.1:${port}/json`)).json()).find(t=>t.type==='page'&&t.url.startsWith('app://4printshops'));}catch{return false;}},'página');
    socket=new WebSocket(target.webSocketDebuggerUrl);await new Promise(resolve=>socket.addEventListener('open',resolve,{once:true}));
    socket.addEventListener('message',e=>{const m=JSON.parse(e.data);if(m.id){requests.get(m.id)?.resolve(m.result);requests.delete(m.id);}if(m.method==='Runtime.exceptionThrown')errors.push(m.params.exceptionDetails.text);});
    await cdp('Runtime.enable');await cdp('Page.enable');await wait(()=>evaluate(`!!document.querySelector('[data-tool="pdfs"]')`),'início');await click('[data-tool="pdfs"]');await ready();
    assert.equal(await evaluate(`!!document.querySelector('#pdf-word-abrir')`),false);
    assert.equal(await evaluate(`document.querySelector('#pdf-arquivo').closest('label').textContent.trim()`),'Importar');
    await command({open:{canceled:true,filePaths:[]}});await click('#pdf-arquivo');await ready();assert.equal(await evaluate(`document.querySelector('#pdf-conteudo').hidden`),true);
    // Um PDF continua direto para o editor, sem passar pela combinação.
    await command({open:{canceled:false,filePaths:[pdfFile]}});await click('#pdf-arquivo');await ready();assert.equal(await evaluate(`Number(document.querySelector('#pdf-total').textContent)`),1);
    await click('#pdf-girar-direita');await ready();const previous=await evaluate(`document.querySelector('#pdf-tamanho').textContent`);
    // Falha em Word não admite silenciosamente os PDFs anteriores do lote.
    await command({open:{canceled:false,filePaths:[pdfFile,invalid]}});await click('#pdf-arquivo');await ready();assert.match(await evaluate(`document.querySelector('#pdf-status').textContent`),/não foi possível importar/i);assert.equal(await evaluate(`document.querySelector('#pdf-tamanho').textContent`),previous);
    // PDF + PNG + JPG + DOCX + DOC: sete páginas, seleção original preservada.
    await command({open:{canceled:false,filePaths:sources}});await click('#pdf-arquivo');await review();
    assert.match(await evaluate(`document.querySelector('#pdf-uniao-total').textContent`),/5 arquivo.*7 página/);
    const labels=await evaluate(`[...document.querySelectorAll('.pdf-uniao-nome')].map(e=>e.textContent)`);
    assert.match(labels[0],/original.pdf/);assert.match(labels[1],/imagem.png/);assert.match(labels[2],/foto.JPG/);assert.match(labels[3],/_WORD.pdf/);
    assert.equal(await evaluate(`document.querySelector('#pdf-uniao-word').hidden`),false);
    assert.match(await evaluate(`document.querySelector('#pdf-uniao-word-info').textContent`),/FonteInexistente4PrintShops/);
    await wait(()=>evaluate(`!!document.querySelectorAll('.pdf-uniao-item')[1]?.querySelector('img')`),'miniatura PNG preparada');
    await cdp('Emulation.setDeviceMetricsOverride',{width:1366,height:768,deviceScaleFactor:1,mobile:false});
    const shot=await cdp('Page.captureScreenshot',{format:'png'});await writeFile(path.join(folder,'importar-lote.png'),Buffer.from(shot.data,'base64'));
    await evaluate(`document.querySelector('#pdf-uniao-destino').value='novo'`);await click('#pdf-uniao-gerar');
    await wait(()=>evaluate(`document.querySelector('#pdf-confirmar-descarte').open`),'proteger alterações');await click('#pdf-continuar-editando');await review();assert.equal(await evaluate(`document.querySelector('#pdf-tamanho').textContent`),previous);
    await click('#pdf-uniao-gerar');await wait(()=>evaluate(`document.querySelector('#pdf-confirmar-descarte').open`),'confirmar substituição');await click('#pdf-descartar-sair');await ready();assert.equal(await evaluate(`Number(document.querySelector('#pdf-total').textContent)`),7);
    assert.equal(await evaluate(`document.querySelector('#pdf-arquivo').closest('label').hidden`),false);
    // Exportação usa o fluxo anterior e PDF reaberto conserva as sete páginas.
    const output=path.join(folder,'combinado.pdf');await command({save:{canceled:false,filePath:output}});await click('#pdf-salvar-copia');await click('#app-exportacao-confirmar');await ready();assert.equal((await PDFDocument.load(await readFile(output))).getPageCount(),7);
    // Word sozinho também é identificado automaticamente e chega à conferência.
    await command({open:{canceled:false,filePaths:[word]}});await click('#pdf-arquivo');await review();assert.match(await evaluate(`document.querySelector('#pdf-uniao-total').textContent`),/1 arquivo.*2 página/);await click('#pdf-uniao-limpar');await click('#pdf-uniao-fechar');
    // Adicionar mais arquivos na combinação utiliza o mesmo conversor.
    await click('#pdf-uniao-abrir');await command({open:{canceled:false,filePaths:[doc]}});await click('#pdf-uniao-arquivos');await review();assert.match(await evaluate(`document.querySelector('#pdf-uniao-total').textContent`),/1 arquivo.*2 página/);await click('#pdf-uniao-limpar');await click('#pdf-uniao-fechar');
    // Cancelar durante preparação não troca o workspace nem adota resultado tardio.
    await command({open:{canceled:false,filePaths:[word]}});await click('#pdf-arquivo');await click('#pdf-import-progress button');await ready();assert.equal(await evaluate(`Number(document.querySelector('#pdf-total').textContent)`),7);
    for(let i=0;i<sources.length;i++)assert.deepEqual(await readFile(sources[i]),originals[i]);
    assert.deepEqual(errors,[]);await command({action:'close',answer:1});await pause(500);
    console.log(JSON.stringify({passed:true,folder,libreOffice:executable,mixedPages:7,sourceFiles:5,originalsPreserved:true}));
}finally{socket?.close();if(child&&child.exitCode===null)child.kill();await writeFile(path.join(folder,'electron.log'),debug);console.log('Evidências: '+folder);}
