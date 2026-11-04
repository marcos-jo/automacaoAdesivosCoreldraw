import test from 'node:test';
import assert from 'node:assert/strict';
import {HoldTestCoordinator} from '../desktop/hold-test.mjs';
import {validatePrintRequest,validateHoldTestRequest} from '../desktop/print-config.mjs';

const state={session:'a'.repeat(32),revision:2,appScale:120};
const proof={token:'b'.repeat(32),jobName:'4PS HOLD EXP bbbbbbbb',printer:'KONICA MINOLTA C6100PS',revision:2,paper:'A3',scale:120,driverZoom:100,copies:1,destination:'Hold',outputValidated:true,cutRisk:false,contentWidthMM:252,contentHeightMM:356.4};
const bytes=Buffer.from('%PDF-1.7\nsynthetic');
const window={isDestroyed:()=>false};
function fixture(response=0,override={}){
    const calls=[];
    const client={request:async request=>{calls.push(request.action);if(request.action==='prepare-hold')return {...proof,...override};return {submitted:true};}};
    const coordinator=new HoldTestCoordinator(client,{showMessageBox:async()=>({response})},()=>window);
    return {coordinator,calls,client};
}
test('cancelar confirmação nunca chama envio e preserva possibilidade de preparar depois',async()=>{
    const {coordinator,calls}=fixture();assert.equal((await coordinator.run(state,bytes)).canceled,true);
    assert.deepEqual(calls,['prepare-hold']);assert.equal(coordinator.attempted,true);
});
test('falha de prova, Print/HoldPrint, fila ou revisão diferentes bloqueiam o envio',async()=>{
    for(const override of [{outputValidated:false},{destination:'Print'},{destination:'HoldPrint'},{printer:'outra'},{revision:3},{copies:2},{scale:100},{driverZoom:120}]){
        const {coordinator,calls}=fixture(1,override);await assert.rejects(coordinator.run(state,bytes));assert.deepEqual(calls,['prepare-hold']);
    }
});
test('confirmação com cliente simulado autoriza uma única chamada e bloqueia reenvio',async()=>{
    const {coordinator,calls}=fixture(1);assert.equal((await coordinator.run(state,bytes)).submitted,true);
    await assert.rejects(coordinator.run(state,bytes));assert.deepEqual(calls,['prepare-hold','send-hold']);
});
test('timeout de envio simulado sinaliza incerteza e nunca repete o trabalho',async()=>{
    const {coordinator,calls,client}=fixture(1);const original=client.request;
    client.request=async request=>{if(request.action==='send-hold'){calls.push(request.action);throw Error('timeout');}return original(request);};
    assert.equal((await coordinator.run(state,bytes)).uncertain,true);await assert.rejects(coordinator.run(state,bytes));assert.deepEqual(calls,['prepare-hold','send-hold']);
});
test('canal comum continua recusando envio e canal experimental valida sessão/token',()=>{
    const session={session:state.session,revision:state.revision};
    assert.throws(()=>validatePrintRequest({action:'send-hold',...session,token:proof.token}));
    assert.equal(validateHoldTestRequest({action:'send-hold',...session,token:proof.token}).action,'send-hold');
    assert.throws(()=>validateHoldTestRequest({action:'send-hold',...session,token:'velho'}));
    assert.throws(()=>validateHoldTestRequest({action:'prepare-hold',...session,pdf:'arquivo local.pdf'}));
});

test('escala selecionada chega ao preparo em 100, 120 e percentual decimal',async()=>{
    for(const scale of [100,120,87.5]){
        let received;
        const client={request:async request=>{received=request;return {...proof,scale};}};
        const coordinator=new HoldTestCoordinator(client,{showMessageBox:async()=>({response:0})},()=>window);
        await coordinator.run({...state,appScale:scale},bytes);assert.equal(received.scale,scale);
    }
});
test('corte cancelado interrompe antes da confirmação de envio; Hold inválido nunca vira aviso',async()=>{
    let dialogs=0;
    const client={request:async()=>({...proof,cutRisk:true})};
    const coordinator=new HoldTestCoordinator(client,{showMessageBox:async(_window,options)=>{dialogs++;assert.equal(options.title,'Possível corte na impressão');return {response:0};}},()=>window);
    assert.equal((await coordinator.run(state,bytes)).canceled,true);assert.equal(dialogs,1);
    const invalid=new HoldTestCoordinator({request:async()=>({...proof,cutRisk:true,destination:'Print'})},{showMessageBox:async()=>{throw Error('Aviso indevido');}},()=>window);
    await assert.rejects(invalid.run(state,bytes),/bloqueado/);
});
test('continuar o aviso geométrico ainda exige confirmação individual de envio',async()=>{
    const titles=[],calls=[];
    const coordinator=new HoldTestCoordinator({request:async request=>{calls.push(request.action);return {...proof,cutRisk:true};}},{showMessageBox:async(_window,options)=>{titles.push(options.title);return {response:titles.length===1?1:0};}},()=>window);
    assert.equal((await coordinator.run(state,bytes)).canceled,true);
    assert.deepEqual(calls,['prepare-hold']);assert.equal(titles.length,2);
});
test('nova operação exige decisão, libera a anterior e retorna sessão diferente sem Hold herdado',async()=>{
    const calls=[];let answer=0;
    const coordinator=new HoldTestCoordinator({request:async request=>{calls.push(request.action);return request.action==='select'?{session:'c'.repeat(32),revision:1,hold:{holdConfirmed:false}}:proof;}},{showMessageBox:async()=>({response:answer})},()=>window);
    await coordinator.run(state,bytes);answer=1;
    const fresh=await coordinator.newOperation();assert.notEqual(fresh.session,state.session);assert.equal(fresh.hold.holdConfirmed,false);assert.equal(coordinator.attempted,false);
    assert.deepEqual(calls,['prepare-hold','release','select']);
});
