import test from 'node:test';
import assert from 'node:assert/strict';
import { PDFDocument } from '../public/vendor/pdf-lib/pdf-lib.js';
import { PDFWorkspace } from './helpers/pdf-resize-modules.mjs';
import { criarPDFAssinado } from './helpers/pdf-signature-fixture.mjs';
const codigo='PDF_UNIAO_OBSOLETA';
const deferred=()=>{let resolve,reject;const promise=new Promise((a,b)=>{resolve=a;reject=b;});return {promise,resolve,reject};};
async function file(name='comum.pdf',width=200){const p=await PDFDocument.create();p.addPage([width,300]);const bytes=await p.save();return {name,arrayBuffer:async()=>bytes.slice().buffer};}
async function suspendSave(ws){const entered=deferred(),gate=deferred(),pdf=ws.pdf,save=pdf.save.bind(pdf);pdf.save=async()=>{entered.resolve();await gate.promise;return save();};return {entered:entered.promise,release:()=>gate.resolve(),fail:e=>gate.reject(e),restore:()=>{pdf.save=save;}};}
for(const assinado of [false,true])test(`união obsoleta não substitui nova abertura: assinado=${assinado}`,async()=>{
 const ws=new PDFWorkspace();await ws.abrir(await file());const hold=await suspendSave(ws);
 const task=ws.adicionarArquivos([await file('adicionar.pdf')]);const rejection=assert.rejects(task,{codigo});await hold.entered;
 const bytes=assinado?await criarPDFAssinado():null;
 await ws.abrir(bytes?{name:'assinado.pdf',arrayBuffer:async()=>bytes.slice().buffer}:await file('novo.pdf',400));
 const state=ws.capturarEstado();hold.release();await rejection;assert.deepEqual(ws.capturarEstado(),state);assert.equal(ws.obterResumoAssinaturas().presentes,assinado?1:0);
});
test('união normal publica somente a cópia completa',async()=>{
 const ws=new PDFWorkspace();await ws.abrir(await file());assert.equal(await ws.adicionarArquivos([await file('novo.pdf',400)]),1);assert.equal(ws.getTotalPaginas(),2);assert.equal(ws.nomeArquivo,'comum.pdf');
});
test('edição no mesmo PDF durante união invalida publicação por revisão',async()=>{
 const ws=new PDFWorkspace();await ws.abrir(await file());const hold=await suspendSave(ws),task=ws.adicionarArquivos([await file()]);const rejection=assert.rejects(task,{codigo});await hold.entered;ws.girarPagina(1,90);const state=ws.capturarEstado();hold.release();await rejection;assert.deepEqual(ws.capturarEstado(),state);assert.equal(ws.pdf.getPage(0).getRotation().angle,90);
});
test('duas uniões fora de ordem: somente a mais recente publica',async()=>{
 const ws=new PDFWorkspace();await ws.abrir(await file());const entered=deferred(),gate=deferred(),pdf=ws.pdf,save=pdf.save.bind(pdf);let calls=0;
 pdf.save=async()=>{if(++calls===1){entered.resolve();await gate.promise;}return save();};
 const old=ws.adicionarArquivos([await file('antigo.pdf',250)]),rejection=assert.rejects(old,{codigo});await entered.promise;
 await ws.adicionarArquivos([await file('recente.pdf',450)]);const state=ws.capturarEstado();gate.resolve();await rejection;assert.deepEqual(ws.capturarEstado(),state);assert.equal(ws.pdf.getPage(1).getWidth(),450);
});
for(const name of ['Error','AbortError'])test(`erro tardio ${name} não substitui a nova abertura`,async()=>{
 const ws=new PDFWorkspace();await ws.abrir(await file());const hold=await suspendSave(ws),task=ws.adicionarArquivos([await file()]),rejection=assert.rejects(task,{codigo});await hold.entered;await ws.abrir(await file('novo.pdf',500));const state=ws.capturarEstado();hold.fail(Object.assign(new Error('Operação antiga cancelada/falhou'),{name}));await rejection;assert.deepEqual(ws.capturarEstado(),state);
});
test('restauração de snapshot não revalida uma união anterior (geração monotônica)',async()=>{
 const ws=new PDFWorkspace();await ws.abrir(await file());const original=ws.capturarEstado(),hold=await suspendSave(ws),task=ws.adicionarArquivos([await file()]),rejection=assert.rejects(task,{codigo});await hold.entered;ws.girarPagina(1,90);ws.restaurarEstado(original);const state=ws.capturarEstado();hold.release();await rejection;assert.deepEqual(ws.capturarEstado(),state);
});
