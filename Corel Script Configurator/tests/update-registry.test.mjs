// Writes only an unpredictable, explicitly isolated UpgradeLab registry namespace.
import test from 'node:test';
import assert from 'node:assert/strict';
import {randomUUID} from 'node:crypto';
import {registryScope,registryRequest} from '../desktop/update-registry.mjs';
test('registry scope refuses non-product and non-lab identities',()=>{
    assert.throws(()=>registryScope({appId:'other.product',guid:randomUUID()}));
    assert.throws(()=>registryScope({appId:'br.com.4printshops.holdtest',guid:randomUUID(),lab:true}));
});
test('real isolated registry snapshot/restore preserves types and never changes PDF default',{skip:process.platform!=='win32'},async()=>{
    const id={appId:'br.com.4printshops.upgradelab.'+randomUUID(),guid:randomUUID(),exe:'4PS Registry Lab.exe',lab:true};
    const scope=registryScope(id),before=await registryRequest(id,{op:'snapshot'});
    assert.ok(before.trees.every(t=>!t.exists));
    const configured=structuredClone(before),tree=configured.trees[0];
    tree.exists=true;tree.children=[];tree.values=[
        {name:'text',kind:'String',data:'Acentos: São José'},
        {name:'expand',kind:'ExpandString',data:'%LOCALAPPDATA%\\test'},
        {name:'binary',kind:'Binary',data:Buffer.from([0,1,128,255]).toString('base64')},
        {name:'number',kind:'DWord',data:'-1'},
        {name:'large',kind:'QWord',data:'9223372036854775807'},
        {name:'multiple',kind:'MultiString',data:['one','dois']}
    ];
    try{
        await registryRequest(id,{op:'restore',snapshot:configured});
        const actual=await registryRequest(id,{op:'snapshot'});
        assert.deepEqual(actual.trees[0],tree);assert.deepEqual(actual.userChoice,before.userChoice);
        await registryRequest(id,{op:'restore',snapshot:before});
        const restored=await registryRequest(id,{op:'snapshot'});assert.deepEqual(restored.trees,before.trees);assert.deepEqual(restored.values.map(({parentExists,...v})=>v),before.values.map(({parentExists,...v})=>v));assert.deepEqual(restored.userChoice,before.userChoice);
        assert.ok(scope.allowedTrees.every(p=>p.startsWith('Software\\4PrintShops\\UpgradeLab\\')));
    }finally{await registryRequest(id,{op:'restore',snapshot:before});}
});
