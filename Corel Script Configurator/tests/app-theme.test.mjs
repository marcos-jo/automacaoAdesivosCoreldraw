import test from 'node:test';
import assert from 'node:assert/strict';
import {aplicarTema,iniciarTema} from '../public/js/app/theme.js';
test('tema padrão claro, escolha escura restaurada e alteração imediata',()=>{
 const data=new Map(),storage={getItem:key=>data.get(key),setItem:(key,value)=>data.set(key,value)},root={dataset:{}};
 assert.equal(iniciarTema(storage,root),'light');aplicarTema('dark',storage,root);assert.equal(root.dataset.theme,'dark');assert.equal(iniciarTema(storage,{dataset:{}}),'dark');aplicarTema('light',storage,root);assert.equal(root.dataset.theme,'light');
});
test('preferência inválida ou armazenamento indisponível mantém tema utilizável',()=>{
 const root={dataset:{}};assert.equal(iniciarTema({getItem:()=>{throw Error('indisponível');},setItem:()=>{throw Error('indisponível');}},root),'light');assert.equal(aplicarTema('invalid',{setItem(){}},root),'light');
});
