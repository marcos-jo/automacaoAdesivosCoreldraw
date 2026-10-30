import test from 'node:test';
import assert from 'node:assert/strict';
import {EventEmitter} from 'node:events';
import {UpdateController,configuredUpdater} from '../desktop/updates.mjs';

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
