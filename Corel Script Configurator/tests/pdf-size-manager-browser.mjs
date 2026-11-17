// Executar com Node 26; --electron usa Electron preparado, cópia verificada e perfil exclusivo.
import fs from 'node:fs/promises';import path from 'node:path';import assert from 'node:assert/strict';
import {tmpdir} from 'node:os';import {spawn} from 'node:child_process';import {createServer} from 'node:http';
import {createServer as netServer} from 'node:net';import {fileURLToPath} from 'node:url';import {createHash} from 'node:crypto';
import {seed,preparar,executar,conferirReabertura} from './helpers/pdf-size-manager-scenarios.mjs';
import {FORMATOS_INICIAIS} from '../public/js/tools/pdf-formats.js';
const root=fileURLToPath(new URL('../',import.meta.url)),electron=process.argv.includes('--electron');
const evidence=await fs.mkdtemp(path.join(tmpdir(),'4ps-size-manager-')),profile=path.join(evidence,'profile'),source=path.join(evidence,'source');await fs.mkdir(profile);
console.log('Evidências: '+evidence);let server,origin,child,socket,serial=0;const errors=[],pending=new Map(),pids=[],closings=[],hashes=[];
if(electron){
    await fs.mkdir(source);
    async function copy(dir){for(const entry of await fs.readdir(path.join(root,dir),{withFileTypes:true})){const rel=path.join(dir,entry.name);if(entry.isDirectory()){await fs.mkdir(path.join(source,rel),{recursive:true});await copy(rel);}else{const b=await fs.readFile(path.join(root,rel));await fs.writeFile(path.join(source,rel),b,{flag:'wx'});const hash=createHash('sha256').update(b).digest('hex');assert.equal(createHash('sha256').update(await fs.readFile(path.join(source,rel))).digest('hex'),hash);hashes.push({path:rel,sha256:hash});}}}
    for(const dir of ['desktop','public']){await fs.mkdir(path.join(source,dir));await copy(dir);}
    await fs.copyFile(path.join(root,'package.json'),path.join(source,'package.json'),fs.constants.COPYFILE_EXCL);
    await fs.symlink(process.env.DESKTOP_TEST_MODULES||'C:/Dev/4PrintShops-Tools/desktop-prep-2026-11-16/source/node_modules',path.join(source,'node_modules'),'junction');
    await fs.writeFile(path.join(evidence,'source-hashes.json'),JSON.stringify(hashes,null,2));
}else{
    server=createServer(async(req,res)=>{try{const name=decodeURIComponent(new URL(req.url,'http://localhost').pathname);if(name==='/favicon.ico'){res.statusCode=204;res.end();return;}const file=path.resolve(root,'public',name==='/'?'index.html':'.'+name);if(!file.startsWith(path.join(root,'public')+path.sep))throw Error('path');res.setHeader('Content-Type',file.endsWith('.js')?'text/javascript':file.endsWith('.css')?'text/css':file.endsWith('.html')?'text/html':'application/octet-stream');res.end(await fs.readFile(file));}catch{res.statusCode=404;res.end();}});
    await new Promise(r=>server.listen(0,'127.0.0.1',r));origin='http://127.0.0.1:'+server.address().port;
}
const wait=async(fn,label)=>{const end=Date.now()+25000;while(Date.now()<end){if(await fn())return;if(child?.exitCode!==null&&child?.exitCode!==undefined)throw Error(label+': processo encerrou '+child.exitCode);await new Promise(r=>setTimeout(r,60));}throw Error('Timeout '+label);};
const cdp=(method,params={})=>new Promise((resolve,reject)=>{const id=++serial,timer=setTimeout(()=>{pending.delete(id);reject(Error(method));},25000);pending.set(id,{resolve:v=>{clearTimeout(timer);resolve(v);},reject:e=>{clearTimeout(timer);reject(e);}});socket.send(JSON.stringify({id,method,params}));});
const ev=async expression=>{const r=await cdp('Runtime.evaluate',{expression,awaitPromise:true,returnByValue:true,userGesture:true});if(r.exceptionDetails)throw Error(r.exceptionDetails.exception?.description||r.exceptionDetails.text);return r.result.value;};
async function launch(){
    const portServer=netServer();await new Promise(r=>portServer.listen(0,'127.0.0.1',r));const port=portServer.address().port;await new Promise(r=>portServer.close(r));
    const env={...process.env,FORMATS_TEST_FOLDER:evidence,FORMATS_TEST_SOURCE:source};for(const k of Object.keys(env))if(/^(ELECTRON_RUN_AS_NODE|NODE_OPTIONS)$/i.test(k))delete env[k];
    const exe=electron?(process.env.ELECTRON_PATH||'C:/Dev/4PrintShops-Tools/desktop-prep-2026-11-16/source/node_modules/electron/dist/electron.exe'):(process.env.BROWSER_PATH||'C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe');
    const args=electron?[path.join(root,'tests/helpers/pdf-size-manager-electron.cjs')]:['--headless=new','--disable-gpu','--no-first-run','--disable-background-networking','--disable-component-update'];
    args.push('--remote-debugging-address=127.0.0.1','--remote-debugging-port='+port,'--user-data-dir='+profile);if(!electron)args.push('about:blank');
    child=spawn(exe,args,{env,windowsHide:true,stdio:['ignore','pipe','pipe']});child.on('error',e=>errors.push(e.stack));for(const stream of [child.stdout,child.stderr])stream.on('data',b=>fs.appendFile(path.join(evidence,'process-'+child.pid+'.log'),b));pids.push(child.pid);
    await wait(async()=>{try{return !!(await(await fetch('http://127.0.0.1:'+port+'/json/version')).json()).webSocketDebuggerUrl;}catch{return false;}},'CDP');
    let tab;await wait(async()=>{tab=(await(await fetch('http://127.0.0.1:'+port+'/json')).json()).find(t=>t.type==='page'&&(!electron||t.url.startsWith('app://4printshops')));return !!tab;},'janela');
    socket=new WebSocket(tab.webSocketDebuggerUrl);await new Promise((r,j)=>{socket.addEventListener('open',r,{once:true});socket.addEventListener('error',j,{once:true});});
    socket.addEventListener('message',e=>{const m=JSON.parse(e.data);if(m.id){const p=pending.get(m.id);if(p){pending.delete(m.id);m.error?p.reject(Error(m.error.message)):p.resolve(m.result);}}if(m.method==='Runtime.exceptionThrown')errors.push(m.params.exceptionDetails.exception?.description||m.params.exceptionDetails.text);if(m.method==='Runtime.consoleAPICalled'&&m.params.type==='error')errors.push(m.params.args.map(a=>a.value||a.description).join(' '));});
    await cdp('Runtime.enable');await cdp('Page.enable');if(!electron)await cdp('Page.navigate',{url:origin});
    await wait(()=>ev(`!!document.querySelector('[data-tool=pdfs]')`),'interface');
    if(electron){await wait(async()=>{try{return JSON.parse(await fs.readFile(path.join(evidence,'electron-ready.json'),'utf8')).pid===child.pid;}catch{return false;}},'main pronto');const ready=JSON.parse(await fs.readFile(path.join(evidence,'electron-ready.json'),'utf8'));assert.equal(ready.pid,child.pid);assert.equal(path.resolve(ready.userData),path.resolve(profile));assert.equal(ready.preferences.sandbox,true);assert.equal(ready.preferences.contextIsolation,true);assert.equal(ready.preferences.nodeIntegration,false);}
}
async function close(){
    assert.equal(await ev('document.querySelectorAll("dialog[open]").length'),0);
    const started=Date.now();let exit;
    const exited=new Promise(resolve=>{if(child.exitCode!==null)resolve(child.exitCode);else child.once('exit',resolve);});
    if(electron)await fs.writeFile(path.join(evidence,'close-request'),'close');else await cdp('Browser.close');
    exit=await Promise.race([exited,new Promise((_,reject)=>{const timer=setTimeout(()=>reject(Error('Encerramento normal excedeu 15s')),15000);timer.unref();})]);
    assert.equal(exit,0);socket.close();closings.push({pid:child.pid,ms:Date.now()-started,exit});
    if(electron){const events=(await fs.readFile(path.join(evidence,'electron-events.ndjson'),'utf8')).trim().split('\n').map(JSON.parse).filter(e=>e.pid===child.pid);for(const event of ['app.quit','before-quit','will-quit','quit'])assert.ok(events.some(e=>e.event===event));await fs.rename(path.join(evidence,'close-request'),path.join(evidence,'closed-'+child.pid));}
}
let result;
try{
    await launch();
    assert.deepEqual(await ev(`(async()=>{const{obterCatalogoFormatos}=await import('/js/tools/pdf-format-catalog.js');const c=obterCatalogoFormatos();await c.pronta;return c.listar().map(({nome,largura,altura})=>({nome,largura,altura}));})()`),FORMATOS_INICIAIS.map(({nome,largura,altura})=>({nome,largura,altura})));
    assert.equal(await ev(`localStorage.getItem('4printshops.pdf.formatos.v1')`),null);
    await ev(`localStorage.setItem('sentinela-tamanhos','preservar');localStorage.setItem('4printshops.pdf.formatos.v1',${JSON.stringify(JSON.stringify(seed))})`);
    // O singleton foi criado pelo menu antes da semente sintética: recarregar somente o perfil isolado.
    await ev('window.__reloadMarker=true');await cdp('Page.reload');await wait(()=>ev(`!window.__reloadMarker&&!!document.querySelector('[data-tool=pdfs]')`),'interface semeada');
    await ev(`document.querySelector('[data-command=gerenciar-tamanhos]').closest('details').open=true;document.querySelector('[data-command=gerenciar-tamanhos]').click()`);
    await wait(()=>ev(`document.querySelector('#pdf-tamanhos').open`),'gerenciador disponível no Início sem PDF');
    assert.equal(await ev(`localStorage.getItem('4printshops.pdf.formatos.v1')`),JSON.stringify(seed));
    await ev(`document.querySelector('#pdf-tamanhos [data-campo=fechar]').click()`);
    await preparar(ev,wait);result=await executar(ev,wait);
    await ev('openManager()');await wait(()=>ev("document.querySelector('#pdf-tamanhos').open"),'captura do gerenciador');
    await fs.writeFile(path.join(evidence,'gerenciador.png'),Buffer.from((await cdp('Page.captureScreenshot',{format:'png'})).data,'base64'));
    await ev("document.querySelector('#pdf-tamanhos [data-campo=fechar]').click()");await close();await launch();await conferirReabertura(ev,wait,result);await close();await launch();await conferirReabertura(ev,wait,result);
    const backup=await ev(`catalog.exportarBackup()`);await fs.writeFile(path.join(evidence,'catalogo-antes-esvaziar.json'),JSON.stringify(backup,null,2),{flag:'wx'});
    const verified=JSON.parse(await fs.readFile(path.join(evidence,'catalogo-antes-esvaziar.json'),'utf8'));assert.equal(createHash('sha256').update(verified.conteudo).digest('hex'),verified.sha256);
    await ev(`(async()=>{for(const f of catalog.listar())await catalog.excluir(f.id);})()`);
    assert.deepEqual(await ev('catalog.listar()'),[]);assert.equal(await ev(`lists().every(l=>l.length===1)`),true);
    await close();await launch();await preparar(ev,wait);assert.deepEqual(await ev('catalog.listar()'),[]);assert.equal(await ev(`lists().every(l=>l.length===1)`),true);
    const empty=await ev('catalog.exportarBackup()');await fs.writeFile(path.join(evidence,'catalogo-vazio.json'),JSON.stringify(empty,null,2),{flag:'wx'});
    const emptyVerified=JSON.parse(await fs.readFile(path.join(evidence,'catalogo-vazio.json'),'utf8'));assert.equal(createHash('sha256').update(emptyVerified.conteudo).digest('hex'),emptyVerified.sha256);
    await ev(`catalog.redefinir(${JSON.stringify(emptyVerified)})`);
    assert.equal(await ev('catalog.listar().length'),9);assert.equal(await ev(`lists().every(l=>l.length===10)`),true);
    await ev(`catalog.restaurarBackup(${JSON.stringify(verified)})`);assert.equal(await ev(`localStorage.getItem('4printshops.pdf.formatos.v1')`),result.persisted);
    assert.equal(await ev(`lists().every(l=>l.includes(${JSON.stringify(result.persistent)})&&!l.includes('SRA3'))`),true);
    assert.equal(await ev(`localStorage.getItem('sentinela-tamanhos')`),'preservar');await ev('(async()=>{await repeat.close();await nup.close();})()');await close();
    assert.equal(new Set(pids).size,4);assert.deepEqual(errors,[]);
    result.checks.A.push('Nove sugestões no primeiro uso; catálogo vazio preservado após reinício; redefinição e recuperação por backup SHA-256 verificado em arquivo');
    await fs.writeFile(path.join(evidence,'resultado.json'),JSON.stringify({electron,pids,closings,checks:result.checks,threeRestarts:true,emptyRestartConfirmed:true,errors,isolatedProfile:profile,sourceFilesVerified:hashes.length},null,2));
    console.log(JSON.stringify({electron,evidence,pids,closings,checks:result.checks,errors}));
}catch(e){await fs.writeFile(path.join(evidence,'falha.json'),JSON.stringify({error:e.stack,errors,pids,closings},null,2));if(socket?.readyState===WebSocket.OPEN){try{await fs.writeFile(path.join(evidence,'estado-falha.txt'),await ev('document.body.innerText'));}catch{}}throw e;}
finally{socket?.close();server?.close();if(child?.exitCode===null){child.unref();child.stdout?.unref();child.stderr?.unref();}}
