import test from 'node:test';
import assert from 'node:assert/strict';
import {EventEmitter} from 'node:events';
import {UpdateController,configuredUpdater,updateSource} from '../desktop/updates.mjs';
import {updateAction} from '../desktop/update-ui.mjs';
import {workStateRequest} from '../desktop/work-state.mjs';
import {resumirTrabalho} from '../public/js/app/desktop-updates.js';

function setup(){const updater=new EventEmitter();let pending=false,installs=0,confirm=async()=>true;
    updater.checkForUpdates=async()=>{updater.emit('checking-for-update');updater.emit('update-available',{version:'1.0.1',releaseNotes:'<b>Correções</b>'});};
    updater.downloadUpdate=async()=>{updater.emit('download-progress',{percent:55});updater.emit('update-downloaded',{version:'1.0.1'});};
    updater.quitAndInstall=()=>installs++;
    const controller=new UpdateController(updater,{canInstall:()=>!pending,confirmInstall:()=>confirm()});
    return {controller,updater,get installs(){return installs;},set pending(v){pending=v;},set confirm(v){confirm=v;}};
}
test('atualização opcional percorre consulta/download/progresso/pronto sem instalar automaticamente',async()=>{
    const s=setup(),phases=[];s.controller.on('state',v=>phases.push(v.phase));
    await s.controller.check();assert.equal(s.controller.state.notes,'Correções');assert.equal(s.installs,0);
    assert.equal(s.updater.autoDownload,false);assert.equal(s.updater.autoInstallOnAppQuit,false);
    s.controller.later();assert.equal(s.installs,0);await s.controller.download();assert.equal(s.installs,0);
    assert.deepEqual(phases,['checking','available','downloading','downloading','ready']);
    assert.equal(await s.controller.install(),true);assert.equal(s.installs,1);
});
test('nenhuma atualização, falha e repetição são estados explícitos',async()=>{
    const s=setup();s.updater.checkForUpdates=async()=>{throw new Error('offline');};await s.controller.check();assert.equal(s.controller.state.phase,'error');
    s.updater.checkForUpdates=async()=>s.updater.emit('update-not-available');await s.controller.check();assert.equal(s.controller.state.phase,'current');
});
test('pendência, mais tarde e alteração durante confirmação impedem instalação',async()=>{
    const s=setup();await s.controller.check();await s.controller.download();s.pending=true;
    await assert.rejects(s.controller.install(),/Exporte/);assert.equal(s.installs,0);
    s.pending=false;s.confirm=async()=>false;assert.equal(await s.controller.install(),false);
    s.confirm=async()=>{s.pending=true;return true;};await assert.rejects(s.controller.install(),/mudou/);assert.equal(s.installs,0);
});
test('sem configuração não há provedor, consulta de rede ou atualização',async()=>{
    const app={getPath:()=>'/4printshops-config-inexistente',isPackaged:true};assert.equal(await configuredUpdater(app,{}),null);
});

test('origem comercial exige HTTPS, publisher e verificação; loopback é restrito ao laboratório',()=>{
    const meta={publisherName:'Gráfica',url:'http://127.0.0.1:18083/',channel:'experimental'};
    assert.throws(()=>updateSource(meta,{enabled:true,url:meta.url}),/HTTPS/);
    assert.equal(updateSource(meta,null,true).url,meta.url);
    for(const url of ['http://localhost/','http://127.0.0.2/','http://example.com/','https://usuario:senha@example.com/','https://example.com/?token=123'])assert.throws(()=>updateSource({...meta,url},null,true));
    assert.throws(()=>updateSource({...meta,publisherName:null},{enabled:true,url:'https://example.com/'}),/editor/);
    assert.throws(()=>updateSource({...meta,verifyUpdateCodeSignature:false},null,true),/ativa/);
    assert.throws(()=>updateSource({...meta,channel:'../latest'},null,true),/Canal/);
    assert.equal(updateSource(meta,{enabled:true,url:'https://example.com/'}).channel,'experimental');
});

test('consulta agregada inclui abas ocultas e futuras ferramentas com estado próprio',()=>{
    const abas=[{controle:{}},{controle:{estadoTrabalho:()=>({pending:true,busy:false})}},{controle:{estadoTrabalho:()=>({pending:false,busy:true})}}];
    assert.deepEqual(resumirTrabalho(abas),{pending:true,busy:true});
    assert.deepEqual(resumirTrabalho([{controle:{}}]),{pending:false,busy:false});
});

test('consulta IPC exige token/frame da janela e falha fechada no timeout/destruição',async()=>{
    let token;const contents={mainFrame:{},send:(_channel,value)=>token=value};
    const work=workStateRequest(()=>({isDestroyed:()=>false,webContents:contents}),{timeout:15});
    const first=work.request();work.receive({sender:contents,senderFrame:{}},token,{pending:false,busy:false});
    assert.deepEqual(await first,{pending:true,busy:true});
    const second=work.request();work.receive({sender:contents,senderFrame:contents.mainFrame},token,{pending:false,busy:false});
    assert.deepEqual(await second,{pending:false,busy:false});
    const third=work.request();work.destroy();assert.deepEqual(await third,{pending:true,busy:true});
});

test('startup notifica sem modal/download; Mais tarde preserva versão e download posterior exige confirmação separada',async()=>{
    const s=setup(),dialogs=[],states=[],locks=[];let options;
    const run=updateAction({}, {getWindow:()=>({isDestroyed:()=>false,setProgressBar(){}}),publish:v=>states.push(v),lock:v=>locks.push(v),canInstall:()=>true,
        configure:async(_app,opts)=>{options=opts;s.controller.confirmInstall=opts.confirmInstall;return s.controller;},
        dialog:{showMessageBox:async(_window,value)=>{dialogs.push(value);return {response:dialogs.length===1?1:0};}}});
    await run.start();assert.equal(dialogs.length,0);assert.equal(s.installs,0);assert.equal(run.getState().phase,'available');
    await run();assert.equal(s.installs,0);assert.equal(run.getState().phase,'available');
    assert.match(dialogs[0].detail,/Correções/);await run();assert.equal(s.installs,1);
    assert.equal(dialogs.at(-1).title,'Instalar atualização e reiniciar');assert.ok(locks.includes(true));
    assert.ok(states.some(s=>s.phase==='downloading'));run.destroy();assert.equal(s.controller.listenerCount('state'),0);assert.equal(locks.at(-1),false);
});

test('startup offline é discreto e consulta em voo não se duplica',async()=>{
    const s=setup();let finish,count=0,dialogs=0;
    s.updater.checkForUpdates=()=>{count++;return new Promise(resolve=>finish=resolve);};
    const run=updateAction({}, {getWindow:()=>null,dialog:{showMessageBox:async()=>{dialogs++;}},canInstall:()=>true,configure:async()=>s.controller});
    const first=run.start();await new Promise(resolve=>setImmediate(resolve));await run.start();assert.equal(count,1);
    s.updater.emit('error',Object.assign(new Error('offline'),{code:'ECONNREFUSED'}));finish();await first;
    assert.equal(dialogs,0);assert.equal(run.getState().code,'ECONNREFUSED');run.destroy();
});

test('download com checksum/assinatura inválidos nunca alcança instalação e pode ser repetido',async()=>{
    const s=setup();await s.controller.check();s.updater.downloadUpdate=async()=>{throw Object.assign(new Error('assinatura inválida'),{code:'ERR_UPDATER_INVALID_SIGNATURE'});};
    await s.controller.download();assert.equal(s.controller.state.code,'ERR_UPDATER_INVALID_SIGNATURE');
    await assert.rejects(s.controller.install(),/pronta/);assert.equal(s.installs,0);await s.controller.check();assert.equal(s.controller.state.phase,'available');
});

test('falha assíncrona do instalador libera a interface e não deixa instalação agendada',async()=>{
    const s=setup(),locks=[];s.updater.quitAndInstall=()=>setImmediate(()=>s.updater.emit('error',new Error('Falha ao abrir instalador')));
    const run=updateAction({}, {getWindow:()=>null,lock:v=>locks.push(v),canInstall:async()=>true,
        configure:async(_app,opts)=>{s.controller.confirmInstall=opts.confirmInstall;return s.controller;},dialog:{showMessageBox:async()=>({response:0})}});
    await run();await new Promise(resolve=>setImmediate(resolve));assert.equal(run.getState().phase,'error');assert.equal(locks.at(-1),false);
    assert.equal(s.updater.autoInstallOnAppQuit,false);run.destroy();
});
