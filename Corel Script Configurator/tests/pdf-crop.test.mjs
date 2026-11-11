import test from 'node:test';
import assert from 'node:assert/strict';
import { PDFDocument, PDFName, degrees } from '../public/vendor/pdf-lib/pdf-lib.js';
import { PDFWorkspace, mmParaPontos as pt } from './helpers/pdf-resize-modules.mjs';
import { gerarMontagem } from './helpers/pdf-imposition-modules.mjs';
const perto=(a,b)=>assert.ok(Math.abs(a-b)<1e-6,`${a} != ${b}`);
async function criar(rotacao=0,origem=0) {
 const pdf=await PDFDocument.create(), p=pdf.addPage([pt(200),pt(100)]);
 p.setMediaBox(origem,-origem,pt(200),pt(100));p.setRotation(degrees(rotacao));
 p.drawText('Vetor preservado',{x:origem+40,y:-origem+40});pdf.addPage([100,100]);
 const bytes=await pdf.save(), ws=new PDFWorkspace();
 await ws.abrir({name:'sintetico.pdf',arrayBuffer:async()=>bytes.slice().buffer});return {ws,bytes};
}
for(const rotacao of [0,90,180,270]) for(const origem of [0,37]) test(`recorte vetorial transacional: rotação ${rotacao}, origem ${origem}`,async()=>{
 const {ws,bytes}=await criar(rotacao,origem),estado=ws.capturarEstado();
 const streams=ws.pdf.getPage(0).node.Contents().toString();
 const prep=await ws.prepararRecorte(1,{x:origem+pt(10),y:-origem+pt(10),width:pt(180),height:pt(80)});
 assert.deepEqual(ws.capturarEstado(),estado); // Cancelar: basta descartar a preparação.
 ws.aplicarRecorte(prep);
 const pdf=await PDFDocument.load(await ws.obterBytes()),p=pdf.getPage(0);
 perto(p.getWidth(),pt(180));perto(p.getHeight(),pt(80));assert.equal(p.getRotation().angle,rotacao);
 assert.equal(p.node.Contents().toString(),streams);assert.equal(pdf.getPageCount(),2);assert.equal(pdf.getPage(1).getWidth(),100);
 assert.deepEqual(await PDFDocument.load(bytes).then(p=>p.getPage(0).getWidth()),pt(200));
 assert.throws(()=>ws.aplicarRecorte(prep),/vencida/);
});
test('recorte → 90×40 → Step & Repeat .37: nove vetores e 16 marcas externas',async()=>{
 const {ws}=await criar();ws.excluirPaginas([2]);
 ws.aplicarRecorte(await ws.prepararRecorte(1,{x:pt(10),y:pt(10),width:pt(180),height:pt(80)}));
 ws.aplicarRedimensionamento(await ws.prepararRedimensionamento({largura:90,altura:40,modo:'ajustar',paginas:[1]}));
 const bytes=await ws.obterBytes(), reopened=await PDFDocument.load(bytes);perto(reopened.getPage(0).getWidth(),pt(90));perto(reopened.getPage(0).getHeight(),pt(40));
 const output=await gerarMontagem(bytes,{modo:'repeat',repetirPorPagina:true,colunas:3,linhas:3,largura:420,altura:297,marcas:true,modoMarcas:'bordas',espacoX:0,espacoY:0});
 const plano=output.plan;assert.equal(plano.folhas.flat().length,9);assert.equal(plano.cortesGrupo[0].length,16);
 for(const piece of plano.folhas.flat()){perto(piece.largura,90);perto(piece.altura,40);}
 const final=await PDFDocument.load(output.bytes);assert.equal(final.getPageCount(),1);
 for(const mark of plano.cortesGrupo[0])for(const [x,y] of [[mark.x1,mark.y1],[mark.x2,mark.y2]])assert.ok(x>=0&&x<=420&&y>=0&&y<=297);
});
test('cancelamento, revisão obsoleta, sangria, assinaturas e estruturas bloqueadas',async()=>{
 const {ws}=await criar();const caixa={x:pt(10),y:pt(10),width:pt(180),height:pt(80)};
 const controller=new AbortController();const prep=await ws.prepararRecorte(1,caixa,{signal:controller.signal});controller.abort();assert.throws(()=>ws.aplicarRecorte(prep),/cancelada/);
 const antiga=await ws.prepararRecorte(1,caixa);ws.girarPagina(2,90);assert.throws(()=>ws.aplicarRecorte(antiga),/vencida/);
 ws.pdf.getPage(0).setBleedBox(0,0,pt(200),pt(100));const sangria=await ws.prepararRecorte(1,caixa);assert.throws(()=>ws.aplicarRecorte(sangria),/sangria/);ws.aplicarRecorte(sangria,{aceitarPerdaSangria:true});assert.equal(ws.pdf.getPage(0).node.has(PDFName.of('BleedBox')),false);
 ws.pdf.catalog.set(PDFName.of('AcroForm'),ws.pdf.context.obj({Fields:[]}));await assert.rejects(ws.prepararRecorte(1,caixa),/bloqueado/);
 const {ws:assinado}=await criar();assinado.pdf.catalog.set(PDFName.of('Perms'),assinado.pdf.context.obj({DocMDP:assinado.pdf.context.obj({Type:'Sig'})}));await assert.rejects(assinado.prepararRecorte(1,caixa),/bloqueado/);
 const {ws:incompativel}=await criar();incompativel.pdf.getPage(0).node.set(PDFName.of('UserUnit'),incompativel.pdf.context.obj(2));await assert.rejects(incompativel.prepararRecorte(1,caixa),/UserUnit/);
 await assert.rejects(incompativel.prepararRecorte(1,{...caixa,width:NaN}));
});
