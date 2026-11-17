// A: orientação/prévia/exportação. B: catálogo v1 compartilhado/persistência.
// Somente Edge headless e perfil temporário, nunca o Hold Test instalado.
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import {tmpdir} from 'node:os';
import {spawn} from 'node:child_process';
import {createServer} from 'node:http';
import {fileURLToPath} from 'node:url';
import {PDFDocument} from '../public/vendor/pdf-lib/pdf-lib.js';
const root=fileURLToPath(new URL('../',import.meta.url));
const evidence=await fs.mkdtemp(path.join(tmpdir(),'4ps-sheet-catalog-')),profile=path.join(evidence,'profile');await fs.mkdir(profile);
console.log('Evidências: '+evidence);
const doc=await PDFDocument.create();doc.addPage([90*72/25.4,50*72/25.4]).drawText('SINTETICO',{x:10,y:50,size:12});const source=await doc.save();
const seed={versao:1,extensao:{preservar:true},formatos:[{id:'custom-f4',nome:'F4',largura:328,altura:478,legado:true},{id:'custom-f6',nome:'F6',largura:318,altura:318},{id:'custom-f8',nome:'F8',largura:238,altura:328},{id:'A3',nome:'A3',largura:297,altura:420},{id:'A4',nome:'A4',largura:210,altura:297}]};
const key='4printshops.pdf.formatos.v1',checks={A:[],B:[]},errors=[];let browser,socket,serial=0;const pending=new Map();
const server=createServer(async(req,res)=>{try{
 if(req.url==='/fixture.pdf'){res.setHeader('Content-Type','application/pdf');res.end(source);return;}
 if(req.url==='/favicon.ico'){res.statusCode=204;res.end();return;}
 const name=decodeURIComponent(new URL(req.url,'http://localhost').pathname),file=path.resolve(root,'public',name==='/'?'index.html':'.'+name);
 if(!file.startsWith(path.join(root,'public')+path.sep))throw Error('path');
 res.setHeader('Content-Type',file.endsWith('.js')?'text/javascript':file.endsWith('.css')?'text/css':file.endsWith('.html')?'text/html':'application/octet-stream');res.end(await fs.readFile(file));
 }catch{res.statusCode=404;res.end();}});
await new Promise(r=>server.listen(0,'127.0.0.1',r));const origin='http://127.0.0.1:'+server.address().port;
const wait=async(fn,label)=>{const end=Date.now()+30000;while(Date.now()<end){if(await fn())return;await new Promise(r=>setTimeout(r,50));}throw Error('Timeout '+label);};
const cdp=(method,params={})=>new Promise((resolve,reject)=>{const id=++serial,timer=setTimeout(()=>{pending.delete(id);reject(Error(method));},30000);pending.set(id,{resolve:v=>{clearTimeout(timer);resolve(v);},reject:e=>{clearTimeout(timer);reject(e);}});socket.send(JSON.stringify({id,method,params}));});
const ev=async expression=>{const r=await cdp('Runtime.evaluate',{expression,awaitPromise:true,returnByValue:true});if(r.exceptionDetails)throw Error(r.exceptionDetails.exception?.description||r.exceptionDetails.text);return r.result.value;};
async function launch(){
 browser=spawn(process.env.BROWSER_PATH||'C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe',['--headless=new','--disable-gpu','--disable-background-networking','--disable-component-update','--no-first-run','--no-default-browser-check','--remote-debugging-port=0','--user-data-dir='+profile,'about:blank'],{windowsHide:true,stdio:'ignore'});
 let port;await wait(async()=>{try{port=(await fs.readFile(path.join(profile,'DevToolsActivePort'),'utf8')).split('\n')[0];return !!(await(await fetch(`http://127.0.0.1:${port}/json/version`)).json()).webSocketDebuggerUrl;}catch{return false;}},'porta ativa do perfil isolado');
 const tab=await(await fetch(`http://127.0.0.1:${port}/json/new?about:blank`,{method:'PUT'})).json();socket=new WebSocket(tab.webSocketDebuggerUrl);await new Promise((r,j)=>{socket.addEventListener('open',r,{once:true});socket.addEventListener('error',j,{once:true});});
 socket.addEventListener('message',e=>{const m=JSON.parse(e.data);if(m.id){const p=pending.get(m.id);if(p){pending.delete(m.id);m.error?p.reject(Error(m.error.message)):p.resolve(m.result);}}if(m.method==='Runtime.exceptionThrown')errors.push(m.params.exceptionDetails.exception?.description||m.params.exceptionDetails.text);});
 await cdp('Runtime.enable');await cdp('Page.enable');await cdp('Page.navigate',{url:origin});await wait(()=>ev("!!document.querySelector('[data-tool=pdfs]')"),'início');
}
async function closeBrowser(){
 const child=browser;try{await cdp('Browser.close');}finally{socket?.close();if(child.exitCode===null)await new Promise(r=>child.once('exit',r));}
}
async function setup(){
 await ev("document.querySelector('[data-tool=pdfs]').click()");await wait(()=>ev("!!document.querySelector('#pdf-redim-formato')"),'Redimensionar');
 await ev(`(async()=>{const {obterCatalogoFormatos}=await import('/js/tools/pdf-format-catalog.js');window.catalog=obterCatalogoFormatos();await catalog.pronta;const {JanelaMontagem}=await import('/js/tools/pdf-imposition-ui.js');const bytes=new Uint8Array(await(await fetch('/fixture.pdf')).arrayBuffer());window.testSource=bytes.slice();window.confirm=()=>true;window.confirmQuestions=[];window.catalogConfirm=false;window.ui=new JanelaMontagem(document.querySelector('#app'),{snapshot:async()=>bytes.slice(),imprimir:()=>{throw Error('IMPRESSAO PROIBIDA NO TESTE');},confirmar:msg=>{confirmQuestions.push(msg);return catalogConfirm;}});window.downloads=[];const create=URL.createObjectURL.bind(URL);window.blobs=new Map();URL.createObjectURL=b=>{const u=create(b);blobs.set(u,b);return u;};HTMLAnchorElement.prototype.click=function(){if(this.download)downloads.push(blobs.get(this.href));};await ui.open();})()`);await ready();
}
const ready=()=>wait(()=>ev('!!ui.result&&!ui.busy&&!ui.dirty&&ui.dialog.querySelector("canvas").width>0'),'prévia');
const edit=async(id,value)=>{await ev(`(()=>{const f=ui.get(${JSON.stringify(id)});f.value=${JSON.stringify(String(value))};f.dispatchEvent(new Event('input',{bubbles:true}));})()`);await ready();};
const state=()=>ev('({config:ui.config(),plan:ui.result.plan,canvas:ui.dialog.querySelector("canvas").width,blocked:ui.get("salvar").disabled})');
const raw=()=>ev(`localStorage.getItem(${JSON.stringify(key)})`);
async function exportSheet(w,h){
 await ev('ui.save()');const sizes=await ev(`(async()=>{const b=downloads.at(-1);if(!b)throw Error('Exportacao ausente');const {PDFDocument}=await import('/vendor/pdf-lib/pdf-lib.js');const pdf=await PDFDocument.load(await b.arrayBuffer());return pdf.getPages().map(p=>[p.getWidth()*25.4/72,p.getHeight()*25.4/72]);})()`);
 assert.ok(sizes.length);for(const [a,b]of sizes){assert.ok(Math.abs(a-w)<.001);assert.ok(Math.abs(b-h)<.001);}
}
async function screenshot(name){const r=await cdp('Page.captureScreenshot',{format:'png'});await fs.writeFile(path.join(evidence,name+'.png'),Buffer.from(r.data,'base64'),{flag:'wx'});}
try{
 await launch();await ev(`localStorage.setItem('sentinela','manter');localStorage.setItem(${JSON.stringify(key)},${JSON.stringify(JSON.stringify(seed))});window.seedReload=true`);await cdp('Page.reload');await wait(()=>ev("!window.seedReload&&!!document.querySelector('[data-tool=pdfs]')"),'catálogo sintético antes da criação do singleton');await setup();
 // Parte A: caminho real do evento, plano, canvas e arquivo de exportação.
 await edit('largura',297);await edit('altura',420);await edit('rotacao',90);
 for(const [o,w,h]of [['paisagem',420,297],['retrato',297,420],['paisagem',420,297],['retrato',297,420]]){
  await edit('orientacao',o);const s=await state();assert.deepEqual([s.config.largura,s.config.altura,s.config.rotacao],[w,h,90]);assert.deepEqual([s.plan.folha.largura,s.plan.folha.altura],[w,h]);assert.ok(s.canvas>0);await exportSheet(w,h);checks.A.push({orientacao:o,largura:w,altura:h,arte:90,exportacao:true});
 }
 await edit('orientacao','paisagem');await screenshot('A-personalizada-paisagem');
 await edit('largura',318);await edit('altura',318);await edit('orientacao','paisagem');assert.deepEqual((await state()).plan.folha,{largura:318,altura:318});await edit('orientacao','retrato');assert.equal((await state()).config.rotacao,90);checks.A.push({quadrada:true});
 for(const [id,w,h]of [['A3',297,420],['A4',210,297]]){await edit('formato',id);await edit('orientacao','paisagem');const s=await state();assert.deepEqual([s.config.largura,s.config.altura,s.config.rotacao],[h,w,90]);checks.A.push({preset:id,paisagem:true});}
 assert.equal(await raw(),JSON.stringify(seed));assert.equal(await ev('ui.originalBytes.every((v,i)=>v===testSource[i])'),true);
 // Parte B: seleção/orientação não grava nem altera os cadastros antigos.
 await edit('formato','custom-f4');await edit('orientacao','retrato');assert.deepEqual([(await state()).config.largura,(await state()).config.altura],[328,478]);await edit('orientacao','paisagem');assert.deepEqual([(await state()).config.largura,(await state()).config.altura],[478,328]);await ev("ui.get('formato-salvar').click()");await wait(()=>ev('!ui.salvandoFormato'),'reusar F4 girado');assert.equal(await raw(),JSON.stringify(seed));
 await edit('formato','custom-f6');await edit('orientacao','retrato');assert.deepEqual([(await state()).config.largura,(await state()).config.altura],[318,318]);checks.B.push({legadosSemGravacao:true});
 await edit('largura',140);await edit('altura',180);await ev("ui.get('formatos-salvos').querySelector('summary').click();ui.get('formato-nome').value='Convite'");await screenshot('B-cadastro-expandido');await ev("ui.get('formato-salvar').click()");await wait(()=>ev("!ui.salvandoFormato&&catalog.listar().some(f=>f.nome==='Convite')"),'salvar na montagem');
 const id=await ev("catalog.listar().find(f=>f.nome==='Convite').id");assert.equal(await ev(`!!document.querySelector('#pdf-redim-formato option[value="${id}"]')`),true);assert.deepEqual(JSON.parse(await raw()).formatos.slice(0,seed.formatos.length),seed.formatos);assert.deepEqual(JSON.parse(await raw()).extensao,seed.extensao);
 await ev("(()=>{window.confirm=()=>false;document.querySelector('#pdf-formato-nome').value='Copia';document.querySelector('#pdf-redim-largura').value='140';document.querySelector('#pdf-redim-altura').value='180';document.querySelector('#pdf-formato-salvar').click();})()");await wait(()=>ev("document.querySelector('#pdf-formatos-status').textContent.includes('cancelado')"),'duplicata no Redimensionar');assert.equal(JSON.parse(await raw()).formatos.length,6);await ev('window.confirm=()=>true');
 await ev("ui.get('formato-salvar').click()");await wait(()=>ev('!ui.salvandoFormato'),'reutilizar mesmo nome');assert.equal(JSON.parse(await raw()).formatos.length,6);
 await ev("ui.get('formato-nome').value='Alias';ui.get('formato-salvar').click()");await wait(()=>ev('!ui.salvandoFormato'),'duplicata por medidas');assert.equal(JSON.parse(await raw()).formatos.length,6);assert.equal(await ev('confirmQuestions.length'),1);
 await ev("catalogConfirm=true;ui.get('formato-nome').value='Alias';ui.get('formato-salvar').click()");await wait(()=>ev("!ui.salvandoFormato&&catalog.listar().some(f=>f.nome==='Alias')"),'novo nome confirmado');const alias=await ev("catalog.listar().find(f=>f.nome==='Alias').id");
 // Criação pelo Redimensionar aparece na montagem sem aplicar outro papel.
 const before=await state();await ev("(()=>{const q=s=>document.querySelector('#pdf-redim-'+s);document.querySelector('#pdf-formato-nome').value='Etiqueta';q('largura').value='90';q('altura').value='50';document.querySelector('#pdf-formato-salvar').click();})()");await wait(()=>ev("catalog.listar().some(f=>f.nome==='Etiqueta')"),'salvar em Redimensionar');const etiqueta=await ev("catalog.listar().find(f=>f.nome==='Etiqueta').id");assert.equal(await ev(`!!ui.get('formato').querySelector('option[value="${etiqueta}"]')`),true);assert.deepEqual((await state()).config,before.config);
 await edit('modo','nup');await edit('formato','custom-f4');await edit('orientacao','retrato');await edit('grade','custom');await edit('colunas',1);await edit('linhas',1);await exportSheet(328,478);checks.B.push({montagemParaRedimensionar:true,redimensionarParaMontagem:true,nupMesmoCatalogo:true,exportacaoNup:true});
 // Editar/excluir um formato em uso preserva as dimensões e a prévia.
 await edit('modo','repeat');await edit('formato',id);await edit('orientacao','paisagem');await ev(`(()=>{const s=document.querySelector('#pdf-redim-formato');s.value=${JSON.stringify(id)};s.dispatchEvent(new Event('change',{bubbles:true}));window.versionAntes=ui.version;})()`);const active=await state();
 await ev(`catalog.salvar({nome:'Convite alterado',largura:150,altura:190},${JSON.stringify(id)})`);assert.equal(await ev("ui.get('formato').value"),'custom');assert.equal(await ev("document.querySelector('#pdf-redim-formato').value"),'personalizado');assert.deepEqual((await state()).config,active.config);assert.equal(await ev('ui.version'),await ev('versionAntes'));
 await edit('formato',alias);const beforeDelete=await state();await ev(`catalog.excluir(${JSON.stringify(alias)})`);assert.equal(await ev("ui.get('formato').value"),'custom');assert.deepEqual((await state()).config,beforeDelete.config);checks.B.push({edicaoExclusaoNaoAplicamDimensoes:true,idsEstaveis:true});
 const persisted=await raw();assert.deepEqual(JSON.parse(persisted).formatos.slice(0,seed.formatos.length),seed.formatos);assert.equal(await ev("localStorage.getItem('sentinela')"),'manter');await screenshot('B-catalogo-compartilhado');
 await closeBrowser();await launch();await setup();assert.equal(await raw(),persisted);await edit('formato',id);await edit('orientacao','retrato');assert.deepEqual([(await state()).config.largura,(await state()).config.altura],[150,190]);checks.B.push({reinicioRealMesmoPerfilOrigem:true});
 // Instâncias concorrentes: uma escreve, outra recusa catálogo desatualizado.
 const concurrency=await ev("(async()=>{const {PDFFormatos}=await import('/js/tools/pdf-formats.js');const a=new PDFFormatos(),b=new PDFFormatos();await Promise.all([a.pronta,b.pronta]);const r=await Promise.allSettled([a.salvar({nome:'Janela A',largura:100,altura:200}),b.salvar({nome:'Janela B',largura:101,altura:201})]);return{locks:!!navigator.locks,ok:r.filter(x=>x.status==='fulfilled').length,refusals:r.filter(x=>x.status==='rejected').map(x=>x.reason.message)};})()");assert.equal(concurrency.locks,true);assert.equal(concurrency.ok,1);assert.match(concurrency.refusals[0],/outra janela/);checks.B.push({concurrency});
 // Corrupção apenas no perfil de teste: qualquer gravação fica bloqueada.
 const valid=await raw();await ev(`localStorage.setItem(${JSON.stringify(key)},'{')`);await cdp('Page.reload');await wait(()=>ev("!!document.querySelector('[data-tool=pdfs]')"),'recarregar');await setup();await ev("ui.get('formato-nome').value='Nao sobrescrever';ui.get('formato-salvar').click()");await wait(()=>ev('!ui.salvandoFormato'),'recusa inválida');assert.equal(await raw(),'{');assert.match(await ev("ui.get('formatos-status').textContent"),/bloqueada/);await ev(`localStorage.setItem(${JSON.stringify(key)},${JSON.stringify(valid)})`);assert.equal(await ev("localStorage.getItem('sentinela')"),'manter');checks.B.push({corrompidoNaoSobrescrito:true,outrosDadosPreservados:true});
 assert.deepEqual(errors,[]);await fs.writeFile(path.join(evidence,'resultado.json'),JSON.stringify({checks,errors,profile,origin,installedAppUsed:false},null,2),{flag:'wx'});console.log(JSON.stringify({checks,errors,evidence}));
}finally{await fs.writeFile(path.join(evidence,'parcial.json'),JSON.stringify({checks,errors,profile},null,2),{flag:'wx'});if(socket?.readyState===WebSocket.OPEN){try{await closeBrowser();}catch{socket.close();}}if(browser&&browser.exitCode===null)browser.kill();server.close();}
