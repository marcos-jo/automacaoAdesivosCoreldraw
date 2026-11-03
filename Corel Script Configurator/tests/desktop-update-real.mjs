// Pacotes/servidor/diálogos reais; sem mocks de electron-updater no app instalado.
import assert from 'node:assert/strict';
import {spawn,execFile} from 'node:child_process';
import {promisify} from 'node:util';
import {mkdtemp,mkdir,readFile,writeFile,copyFile,unlink,realpath} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import path from 'node:path';
import {setTimeout as pause} from 'node:timers/promises';
import {createHash} from 'node:crypto';
import yaml from 'js-yaml';
import {serveUpdateLab} from '../scripts/serve-update-lab.mjs';
import {PDFDocument} from '../public/vendor/pdf-lib/pdf-lib.js';

const exec=promisify(execFile),versions=['1.0.0-experimental.2','1.0.0-experimental.3'];
const folder=await mkdtemp(path.join(tmpdir(),'4printshops-update-real-')),installation=path.join(folder,'installed');
const profile=path.join(process.env.APPDATA,'4PrintShops Update Lab');
const artifacts=version=>path.resolve('dist-desktop/update-lab',version);
const setup=version=>path.join(artifacts(version),`4PrintShops-Update-Lab-${version}-Setup.exe`);
const exe=path.join(installation,'4PrintShops Update Lab.exe');
const requests=[],exceptions=[];let server,child,socket,sequence=0;const pending=new Map();
const wait=async(fn,label,timeout=40000)=>{const end=Date.now()+timeout;while(Date.now()<end){const result=await fn();if(result)return result;await pause(60);}throw new Error('Tempo esgotado: '+label);};
const cdp=(method,params={})=>new Promise((resolve,reject)=>{const id=++sequence,timer=setTimeout(()=>{pending.delete(id);reject(new Error('CDP: '+method));},10000);pending.set(id,{resolve:v=>{clearTimeout(timer);resolve(v);},reject});socket.send(JSON.stringify({id,method,params}));});
const evaluate=async expression=>{const r=await cdp('Runtime.evaluate',{expression,awaitPromise:true,returnByValue:true,userGesture:true});if(r.exceptionDetails)throw new Error(r.exceptionDetails.exception?.description||r.exceptionDetails.text);return r.result.value;};
const click=selector=>evaluate(`document.querySelector(${JSON.stringify(selector)}).click()`);
const state=()=>evaluate('window.desktop4PrintShops.estadoAtualizacao()');
const native=async(action,destination='')=>{
    const script=await readFile(new URL('./helpers/windows-save-dialog.ps1',import.meta.url),'utf8');
    const quote=value=>"'"+value.replaceAll("'","''")+"'";
    return exec('powershell.exe',['-NoProfile','-Command',`& {${script}} -DialogProcess ${child.pid} -Action ${quote(action)} -Destination ${quote(destination)}`],{windowsHide:true,timeout:25000});
};
const screenshot=async name=>{const result=await cdp('Page.captureScreenshot');await writeFile(path.join(folder,name+'.png'),Buffer.from(result.data,'base64'));};
async function install(version){
    assert.ok(installation.startsWith(folder+path.sep)&&(await realpath(folder)).toLowerCase().startsWith((await realpath(tmpdir())).toLowerCase()+path.sep),'Destino somente no laboratório temporário');
    await new Promise((resolve,reject)=>{const process=spawn(setup(version),['/S','/D='+installation],{windowsHide:true,stdio:'ignore'});process.once('error',reject);process.once('exit',code=>code===0?resolve():reject(new Error('NSIS: '+code)));});
}
async function launch(){
    await unlink(path.join(profile,'DevToolsActivePort')).catch(()=>{});
    const env={...process.env,PATH:path.join(process.env.SystemRoot,'System32')+';'+process.env.SystemRoot};delete env.ELECTRON_RUN_AS_NODE;
    child=spawn(exe,['--remote-debugging-port=0'],{cwd:installation,env,windowsHide:false,stdio:'ignore'});
    const port=await wait(async()=>{try{return(await readFile(path.join(profile,'DevToolsActivePort'),'utf8')).split('\n')[0];}catch{return false;}},'App instalado iniciou');
    const target=await wait(async()=>{try{return(await(await fetch(`http://127.0.0.1:${port}/json`)).json()).find(t=>t.type==='page');}catch{return false;}},'Janela');
    socket=new WebSocket(target.webSocketDebuggerUrl);await new Promise(resolve=>socket.addEventListener('open',resolve,{once:true}));
    socket.addEventListener('message',event=>{const value=JSON.parse(event.data);if(value.id){if(value.error)pending.get(value.id)?.reject(new Error(value.error.message));else pending.get(value.id)?.resolve(value.result);pending.delete(value.id);}if(value.method==='Runtime.exceptionThrown')exceptions.push(value.params.exceptionDetails.text);});
    await cdp('Runtime.enable');await cdp('Page.enable');await wait(()=>evaluate("!!document.querySelector('[data-tool=pdfs]')"),'Interface');
}
async function close(){await evaluate('window.close()');await wait(()=>child.exitCode!==null,'Encerramento');socket.close();}
const nativeCheck=async action=>{await evaluate('void window.desktop4PrintShops.verificarAtualizacoes()');await native(action);};

try{
    // Reutilizar só perfil que este harness identificou como laboratório.
    const owner=path.join(profile,'4printshops-test-owner.txt');
    let exists=false;try{await readFile(path.join(profile,'Preferences'));exists=true;}catch(error){if(error.code!=='ENOENT')throw error;}
    if(exists)assert.equal(await readFile(owner,'utf8'),'desktop-update-real','Perfil preexistente precisa pertencer ao harness');
    await mkdir(profile,{recursive:true});await writeFile(owner,'desktop-update-real');
    const manifest=await readFile(path.join(artifacts(versions[1]),'experimental.yml'),'utf8'),info=yaml.load(manifest);
    assert.equal(info.version,versions[1]);
    const installerBytes=await readFile(setup(versions[1]));
    assert.equal(createHash('sha512').update(installerBytes).digest('base64'),info.files[0].sha512,'SHA512 real gerado pelo builder');
    await copyFile(setup(versions[1]),path.join(folder,info.files[0].url));
    await copyFile(setup(versions[1])+'.blockmap',path.join(folder,info.files[0].url+'.blockmap'));
    await writeFile(path.join(folder,'experimental.yml'),manifest);
    server=await serveUpdateLab(folder,{onRequest:request=>requests.push(request),bytesPerSecond:12*1024*1024});
    await install(versions[0]);await launch();
    assert.equal((await state()).currentVersion,versions[0]);
    await evaluate('window.updateTestStates=[];window.desktop4PrintShops.observarAtualizacoes(value=>window.updateTestStates.push(value))');
    await wait(async()=>(await state()).phase==='available','Detecção no startup');
    assert.match((await state()).notes,/Consulta de todas as abas/);
    assert.equal(requests.some(r=>r.name.endsWith('.exe')),false,'Startup não baixa instalador');
    await screenshot('01-notificacao');
    await nativeCheck('update-later');assert.equal((await state()).phase,'available');
    assert.equal(requests.some(r=>r.name.endsWith('.exe')),false,'Mais tarde não baixa');
    await click('#desktop-update-notice button:last-child');
    assert.equal(await evaluate("document.querySelector('#desktop-update-notice').hidden"),true);
    await click('[data-tool=qrcodes]');await click('[data-tool=inicio]');
    const prefs={'4printshops.update-lab.preferencia':'persistiu','4printshops.pdf.formatos.v1':JSON.stringify({versao:1,formatos:[{id:'custom-teste-lab',nome:'Formato Lab',largura:123.5,altura:234.5}]})};
    await evaluate(`Object.entries(${JSON.stringify(prefs)}).forEach(([k,v])=>localStorage.setItem(k,v))`);
    // Uma edição real em aba oculta continua protegida durante o ensaio.
    const pdf=await PDFDocument.create();pdf.addPage([200,300]);const original=Buffer.from(await pdf.save()),pdfPath=path.join(folder,'original.pdf');await writeFile(pdfPath,original);
    await click('[data-tool=pdfs]');await wait(()=>evaluate("!document.querySelector('#pdf-arquivo').disabled"),'PDF pronto');
    await evaluate('void document.querySelector("#pdf-arquivo").click()');await native('open',pdfPath);
    await wait(()=>evaluate("document.querySelector('#pdf-nome').textContent==='original.pdf'&&!document.querySelector('#pdf-arquivo').disabled"),'PDF importado');
    await click('#pdf-girar-direita');await wait(()=>evaluate("!document.querySelector('#pdf-arquivo').disabled"),'Giro');await click('[data-tool=inicio]');
    await nativeCheck('update-download');
    await wait(async()=>(await state()).phase==='error','Rejeição de assinatura',90000);
    const rejected=await state();assert.equal(rejected.code,'ERR_UPDATER_INVALID_SIGNATURE');
    await native('updates');assert.equal(child.exitCode,null);assert.ok(requests.some(r=>r.name.endsWith('.exe')));
    assert.equal(await evaluate("window.updateTestStates.some(s=>s.phase==='downloading'&&s.percent>0&&s.percent<100)"),true,'Progresso intermediário real comunicado à interface');
    await click('[data-tool=pdfs]');assert.equal(await evaluate("document.querySelector('#pdf-alteracoes-pendentes').hidden"),false);
    assert.equal(await evaluate("document.querySelector('#app-shell').inert"),false);await screenshot('02-assinatura-rejeitada');
    console.log('PASSOU: experimental.2 instalada, startup/notificação/novidades, Mais tarde, uso de QR/PDF e download real; pacote sem assinatura rejeitado com ERR_UPDATER_INVALID_SIGNATURE, PDF pendente preservado.');
    // Falha de integridade: metadado de teste alterado deliberadamente, sem substituir o binário.
    const damaged=yaml.load(manifest);damaged.files[0].sha512=Buffer.alloc(64,1).toString('base64');damaged.sha512=damaged.files[0].sha512;
    await writeFile(path.join(folder,'experimental.yml'),yaml.dump(damaged));
    await nativeCheck('update-download');await wait(async()=>(await state()).phase==='error','Checksum rejeitado',90000);
    assert.match((await state()).message,/sha512|checksum/i);await native('updates');
    await writeFile(path.join(folder,'experimental.yml'),manifest);
    console.log('PASSOU: download com SHA512 adulterado rejeitado; nenhuma instalação.');
    // Offline e recuperação sem reiniciar o app.
    await new Promise(resolve=>server.close(resolve));server=null;
    await evaluate('void window.desktop4PrintShops.verificarAtualizacoes()');
    await wait(async()=>(await state()).phase==='error','Erro de conexão',40000);await native('updates');
    server=await serveUpdateLab(folder,{onRequest:request=>requests.push(request)});
    await nativeCheck('update-later');assert.equal((await state()).phase,'available');
    console.log('PASSOU: falha de conexão informada e recuperação após voltar o servidor.');
    // Encerrar conscientemente antes da instalação MANUAL complementar da .3.
    const cancelClose=evaluate('window.close()').catch(()=>{});await native('continue');await cancelClose;assert.equal(child.exitCode,null);
    const discardClose=evaluate('window.close()').catch(()=>{});await native('discard');await wait(()=>child.exitCode!==null,'Descarte consciente');await discardClose;socket.close();
    await install(versions[1]);await launch();assert.equal((await state()).currentVersion,versions[1]);
    await wait(async()=>(await state()).phase==='current','Versão .3 atual');
    for(const [key,value]of Object.entries(prefs))assert.equal(await evaluate(`localStorage.getItem(${JSON.stringify(key)})`),value);
    assert.equal(await evaluate("(async()=>{const {PDFFormatos}=await import('/js/tools/pdf-formats.js');const f=new PDFFormatos();await f.pronta;return f.listar().find(x=>x.id==='custom-teste-lab').largura;})()"),123.5);
    assert.deepEqual(await readFile(pdfPath),original);assert.deepEqual(exceptions,[]);
    await screenshot('03-versao-3-preferencias');await close();
    console.log('PASSOU: instalação MANUAL complementar da experimental.3, versão em execução confirmada, preferências/formatos e PDF original preservados.');
    console.log('LIMITAÇÃO: instalação/reinício pelo updater NÃO executados: falta Authenticode confiável. Não houve bypass nem certificado adicionado ao Windows.');
    await writeFile(path.join(folder,'requests.json'),JSON.stringify(requests,null,2));
    await writeFile(path.join(folder,'result.json'),JSON.stringify({versions,rejected,installerSha512:info.files[0].sha512,automaticInstall:false,manualInstall:true,installation,profile},null,2));
}catch(error){try{await screenshot('falha');console.error('Estado:',await state());}catch{}console.error(error);process.exitCode=1;}
finally{socket?.close();if(child&&child.exitCode===null)child.kill();if(server)await new Promise(resolve=>server.close(resolve));console.log('Artefatos: '+folder);}
