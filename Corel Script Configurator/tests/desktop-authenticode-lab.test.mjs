import test from 'node:test';
import assert from 'node:assert/strict';
import guard from '../scripts/authenticode-lab/guest-guard.cjs';
const manifest={schema:'4printshops-authenticode-kit-v1',hostFingerprint:'a'.repeat(64)};
test('laboratório recusa host, ausência de confirmação e modelo físico',()=>{
    assert.throws(()=>guard.validateGuest(manifest,{fingerprint:'b'.repeat(64),model:'Virtual Machine'},false),/explicitamente/);
    assert.throws(()=>guard.validateGuest(manifest,{fingerprint:'a'.repeat(64),model:'Virtual Machine'},true),/HOST PRINCIPAL/);
    assert.throws(()=>guard.validateGuest(manifest,{fingerprint:'b'.repeat(64),model:'H410M H'},true),/não reconhecido/);
    assert.throws(()=>guard.validateGuest({}, {},true),/inválido/);
});
test('guest distinto e reconhecido exige manifesto do host e confirmação',()=>{
    const guest={fingerprint:'b'.repeat(64),model:'VirtualBox',manufacturer:'Oracle'};
    assert.equal(guard.validateGuest(manifest,guest,true),guest);
    assert.throws(()=>guard.assertGuest(undefined,false),/confirmação/);
});
