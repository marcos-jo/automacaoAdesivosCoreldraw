import test from 'node:test';
import assert from 'node:assert/strict';
import {HoldTestCoordinator} from '../desktop/hold-test.mjs';
import {validatePrintRequest,validateHoldTestRequest} from '../desktop/print-config.mjs';

const state={session:'a'.repeat(32),revision:2};
const proof={token:'b'.repeat(32),jobName:'4PS HOLD EXP bbbbbbbb',printer:'KONICA MINOLTA C6100PS',revision:2,paper:'A3',scale:120,copies:1,destination:'Hold',outputValidated:true};
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
    assert.deepEqual(calls,['prepare-hold']);assert.equal(coordinator.attempted,false);
});
test('falha de prova, Print/HoldPrint, fila ou revisão diferentes bloqueiam o envio',async()=>{
    for(const override of [{outputValidated:false},{destination:'Print'},{destination:'HoldPrint'},{printer:'outra'},{revision:3},{copies:2},{scale:100}]){
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
    assert.throws(()=>validatePrintRequest({action:'send-hold',...state,token:proof.token}));
    assert.equal(validateHoldTestRequest({action:'send-hold',...state,token:proof.token}).action,'send-hold');
    assert.throws(()=>validateHoldTestRequest({action:'send-hold',...state,token:'velho'}));
    assert.throws(()=>validateHoldTestRequest({action:'prepare-hold',...state,pdf:'arquivo local.pdf'}));
});
