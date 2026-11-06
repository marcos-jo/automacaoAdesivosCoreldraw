import test from 'node:test';import assert from 'node:assert/strict';
import {PDFDocument,PDFName,degrees} from '../public/vendor/pdf-lib/pdf-lib.js';
import {planejarImagemMontagem,avaliarImagemMontagem,prepararImagemMontagem} from './helpers/pdf-imposition-raster-modules.mjs';
const N=PDFName.of;
const fixture=async()=>{const pdf=await PDFDocument.create();pdf.addPage([248,142]);return pdf;};
test('imagem é uma alternativa explícita: autorização obrigatória, sem worker ou alteração ao recusar',async()=>{
    const bytes=await(await fixture()).save(),old=bytes.slice();let workers=0;
    await assert.rejects(prepararImagemMontagem(bytes,{dpi:300,load:async()=>{workers++;}}),/Confirme/);
    assert.equal(workers,0);assert.deepEqual(bytes,old);
});
test('preflight OCG/OutputIntents permite raster sem liberar o diagnóstico vetorial',async()=>{
    const {diagnosticarMontagem}=await import('./helpers/pdf-imposition-modules.mjs');
    for(const type of ['OCProperties','OutputIntents']){const pdf=await fixture();pdf.catalog.set(N(type),pdf.context.obj(type==='OutputIntents'?[]:{}));const diag=diagnosticarMontagem(pdf);assert.equal(diag.podePreparar,false);assert.ok(diag.bloqueios.some(b=>b.codigo===type));const plans=await avaliarImagemMontagem(await pdf.save());assert.deepEqual(plans.map(p=>p.disponivel),[true,true]);assert.equal(diagnosticarMontagem(pdf).necessitaPreparacao,true);}
});
test('300/600 DPI são explícitos; rotação, UserUnit e área CropBox entram no limite de memória',async()=>{
    const pdf=await PDFDocument.create(),p=pdf.addPage([300,200]);p.setCropBox(10,10,280,180);p.setRotation(degrees(90));p.node.set(N('UserUnit'),pdf.context.obj(2));
    const plan=planejarImagemMontagem(pdf,300);assert.equal(plan.paginas[0].largura,1500);assert.equal(plan.paginas[0].altura,2334);assert.throws(()=>planejarImagemMontagem(pdf,150),/300\/600/);assert.equal(p.getRotation().angle,90);
});
test('XFA, scripts, ações e aparências dinâmicas não podem ser contornados pela opção imagem',async()=>{
    for(const type of ['XFA','NeedAppearances','AA','JavaScript']){const pdf=await fixture();if(type==='AA')pdf.catalog.set(N('AA'),pdf.context.obj({}));else if(type==='JavaScript')pdf.catalog.set(N('Names'),pdf.context.obj({JavaScript:{}}));else pdf.catalog.set(N('AcroForm'),pdf.context.obj(type==='XFA'?{XFA:'unsupported'}:{NeedAppearances:true}));const plans=await avaliarImagemMontagem(await pdf.save());assert.ok(plans.every(p=>!p.disponivel&&p.motivo));}
});
test('limites por página e agregado bloqueiam antes de iniciar PDF.js; nenhuma redução silenciosa',async()=>{
    const pdf=await PDFDocument.create();pdf.addPage([10000,10000]);assert.throws(()=>planejarImagemMontagem(pdf,300),/limite/);
    const many=await PDFDocument.create();for(let i=0;i<20;i++)many.addPage([595,842]);assert.throws(()=>planejarImagemMontagem(many,300),/256 MiB/);
    let started=false;await assert.rejects(prepararImagemMontagem(await pdf.save(),{dpi:600,confirmado:true,load:async()=>{started=true;}}),/limite/);assert.equal(started,false);
});
test('cancelamento e falha de carregamento preservam bytes; tarefa/worker têm dono único',async()=>{
    const bytes=await(await fixture()).save(),original=bytes.slice(),a=new AbortController();a.abort();
    await assert.rejects(prepararImagemMontagem(bytes,{dpi:300,confirmado:true,signal:a.signal}),{name:'AbortError'});
    let destroyed=0,workers=0;
    await assert.rejects(prepararImagemMontagem(bytes,{dpi:300,confirmado:true,load:async()=>({GlobalWorkerOptions:{},PDFWorker:class{destroy(){workers++;}},getDocument:()=>({promise:Promise.reject(Error('Falha de leitura simulada')),destroy:async()=>{destroyed++;}})})}),/Falha de leitura/);
    assert.equal(destroyed,1);assert.equal(workers,1);assert.deepEqual(bytes,original);
});
