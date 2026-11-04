// Exercita somente propriedades e leitura do PrintTicket da Konica. O auxiliar não recebe PDFs.
import assert from 'node:assert/strict';
import {spawn,execFile} from 'node:child_process';
import {promisify} from 'node:util';
import {readFile,writeFile,mkdir} from 'node:fs/promises';
import {setTimeout as pause} from 'node:timers/promises';
import path from 'node:path';
import {PrintConfigClient} from '../desktop/print-config.mjs';

const exec=promisify(execFile);
const powershell=path.join(process.env.SystemRoot,'System32/WindowsPowerShell/v1.0/powershell.exe');
const client=new PrintConfigClient({helper:path.resolve('build/native/4PrintShops.PrintConfig.exe'),manifest:path.resolve('build/native/manifest.json'),
    launch:(file,args,options)=>spawn(file,args,{...options,windowsHide:false})});
const source=await readFile(new URL('./helpers/windows-print-properties.ps1',import.meta.url),'utf8');
let monitor,monitorOutput='',monitorError='';
const wait=async(fn,label)=>{const until=Date.now()+30000;while(Date.now()<until){if(await fn())return;await pause(100);}throw Error('Tempo esgotado: '+label);};
const native=async action=>exec(powershell,['-NoProfile','-NonInteractive','-Command',`& {${source}} -DriverProcess ${client.child.pid} -Action '${action}'`],{windowsHide:true,timeout:30000});
const verify=state=>client.request({action:'verify-hold',session:state.session,revision:state.revision});
try {
    const monitorSource=await readFile(new URL('./helpers/windows-print-defaults.ps1',import.meta.url),'utf8');
    monitor=spawn(powershell,['-NoProfile','-NonInteractive','-Command',`& {${monitorSource}} -PrinterName 'KONICA MINOLTA C6100PS'`],{windowsHide:true,stdio:['pipe','pipe','pipe']});
    monitor.stdout.on('data',chunk=>monitorOutput+=chunk);monitor.stderr.on('data',chunk=>monitorError+=chunk);
    await wait(()=>{if(monitorError)throw Error('Monitor de preferências indisponível.');return monitorOutput.includes('READY');},'Monitor');
    let state=await client.request({action:'select',printer:'KONICA MINOLTA C6100PS'});
    const initial=await verify(state);assert.equal(initial.hold.outputMethod,'Print');assert.equal(initial.hold.holdConfirmed,false);
    const change=async(action,cancel=false)=>{
        const previous=state;
        const pending=client.request({action:'properties',session:state.session,revision:state.revision});
        await native('workflow');await pause(250);await native(action);
        await native(cancel?'cancel':'accept');
        state=await pending;
        if(cancel)assert.equal(state.revision,previous.revision);else assert.equal(state.hold.verified,false,'Confirmar invalida a verificação anterior');
        const checked=await verify(state);
        console.log('Métodos públicos:',checked.hold.outputMethod,checked.hold.printMode);
        return checked;
    };
    state=await change('hold');
    assert.equal(state.hold.outputMethod,'Hold');assert.equal(state.hold.printMode,'Print');assert.equal(state.hold.holdConfirmed,true);assert.equal(state.hold.canSend,false);
    const canceled=await change('print',true);assert.equal(canceled.hold.holdConfirmed,true,'Cancelar conserva o Hold confirmado');
    state=canceled;
    state=await change('holdprint');assert.equal(state.hold.outputMethod,'HoldPrint');assert.equal(state.hold.holdConfirmed,false);
    state=await change('print');assert.equal(state.hold.outputMethod,'Print');assert.equal(state.hold.holdConfirmed,false);
    await client.request({action:'release'});
    const reopened=await client.request({action:'select',printer:'KONICA MINOLTA C6100PS'});
    assert.equal((await verify(reopened)).hold.outputMethod,'Print');
    monitorOutput='';monitor.stdin.write('check\n');await wait(()=>monitorOutput.includes('CHANGED'),'Preferências');assert.equal(monitorOutput.trim(),'UNCHANGED');
    await mkdir('.desktop-test/hold-lab',{recursive:true});
    await writeFile('.desktop-test/hold-lab/hold-native-result.json',JSON.stringify({passed:true,observed:['Print','Hold','HoldPrint'],cancelPreservedHold:true,confirmationInvalidated:true,globalAndUserDefaultsPreserved:true,physicalSendingAvailable:false},null,2));
    console.log('PASSOU: Print/Hold/HoldPrint reais, cancelamento, invalidação por revisão e defaults intactos. Envio físico bloqueado.');
} finally {
    client.destroy();
    if(monitor){monitor.stdin.end('quit\n');monitor.kill();}
}
