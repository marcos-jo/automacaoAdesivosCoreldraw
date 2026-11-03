import test from 'node:test';
import assert from 'node:assert/strict';
import path from 'node:path';
import {verifyInstallerSignature} from '../desktop/update-signature.mjs';

const file=path.resolve("installer-'teste.pdf.exe"),publishers=['4PrintShops'];
const valid={status:0,embedded:true,subject:'CN=4PrintShops, O=Grafica',path:file};
const run=result=>verifyInstallerSignature(publishers,file,{execute:async()=>({stdout:JSON.stringify(result)})});
test('somente assinatura válida, incorporada, do editor e arquivo esperados é aceita',async()=>{
    assert.equal(await run(valid),null);
    for(const change of [{status:2},{embedded:false},{subject:'CN=Outro'},{path:file+'.outro'},{subject:null}])assert.ok(await run({...valid,...change}));
});
test('PowerShell ausente, timeout e JSON inválido bloqueiam; executável absoluto e sem shell',async()=>{
    for(const execute of [async()=>{throw new Error('ENOENT');},async()=>{throw new Error('timeout');},async()=>({stdout:'invalid'})])assert.ok(await verifyInstallerSignature(publishers,file,{execute}));
    let invocation;
    assert.equal(await verifyInstallerSignature(publishers,file,{execute:async(...args)=>{invocation=args;return {stdout:JSON.stringify(valid)};}}),null);
    assert.ok(path.isAbsolute(invocation[0]));assert.equal(invocation[2].shell,undefined);assert.equal(invocation[2].windowsHide,true);
    assert.ok(invocation[1].at(-1).includes(file.replaceAll("'","''")));
});
