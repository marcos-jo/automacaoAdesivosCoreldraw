import test from 'node:test';import assert from 'node:assert/strict';
import {PDFDocument,PDFName,PDFStream} from '../public/vendor/pdf-lib/pdf-lib.js';
import {PDFWorkspace,mmParaPontos} from './helpers/pdf-resize-modules.mjs';
import {marcarEstrutura} from './helpers/pdf-print-fixture.mjs';
const N=PDFName.of,config={largura:200,altura:150,modo:'ajustar',paginas:[1]};
async function abrir({atributo=false,essencial=true}={}){
    const p=await PDFDocument.create(),page=p.addPage([720,540]);page.drawText('Synthetic marked vector',{x:20,y:100,size:24});
    const c=p.context,old=page.node.Contents().asArray();page.node.set(N('Contents'),c.obj([c.register(c.stream('/P << /MCID 0 >> BDC\n')),...old,c.register(c.stream('\nEMC'))]));
    marcarEstrutura(p,page,essencial?'Texto substituto essencial':'Synthetic marked vector');
    const element=p.catalog.lookup(N('StructTreeRoot')).lookup(N('K')).lookup(0);
    if(atributo)element.set(N('A'),c.obj({O:'Layout',BBox:[0,0,100,100]}));
    const bytes=await p.save(),ws=new PDFWorkspace();await ws.abrir({name:'synthetic.pdf',arrayBuffer:async()=>bytes.slice().buffer});return ws;
}
const streams=pdf=>new Map(pdf.context.enumerateIndirectObjects().filter(([,o])=>o instanceof PDFStream).map(([r,o])=>[r.toString(),Buffer.from(o.getContents()).toString('hex')]));
test('720×540pt para200×150mm mantém ActualText essencial e streams; descarte segue protegido',async()=>{
    const ws=await abrir(),state=ws.capturarEstado(),sourceStreams=streams(ws.pdf),root=ws.pdf.catalog.get(N('StructTreeRoot')).toString();
    const preview=await ws.prepararRedimensionamento(config);assert.equal(preview.marcacaoPreservada,true);assert.equal(preview.preparadoParaImpressao,false);assert.deepEqual(preview.removidos,[]);assert.deepEqual(ws.capturarEstado(),state);
    for(const[r,bytes]of sourceStreams)assert.equal(streams(preview.pdf).get(r),bytes,'Fontes/imagens/streams originais não rasterizados');
    ws.aplicarRedimensionamento(preview);assert.equal(ws.alterado,true);const output=await PDFDocument.load(await ws.obterBytes());assert.ok(Math.abs(output.getPage(0).getWidth()-mmParaPontos(200))<1e-7);assert.ok(Math.abs(output.getPage(0).getHeight()-mmParaPontos(150))<1e-7);assert.equal(output.catalog.get(N('StructTreeRoot')).toString(),root);assert.equal(output.getPage(0).node.get(N('StructParents')).toString(),'0');
    await assert.rejects((await abrir()).prepararRedimensionamento({...config,modoImpressao:true}),/ActualText/);
});
test('atributos de layout e conteúdo externo não são liberados para preservação sem tratamento',async()=>{
    const ws=await abrir({atributo:true}),state=ws.capturarEstado();await assert.rejects(ws.prepararRedimensionamento(config),/ActualText/);assert.deepEqual(ws.capturarEstado(),state);
});
test('prévia/cancelamento e edições posteriores continuam transacionais com marcação preservada',async()=>{
    const ws=await abrir();ws.girarPagina(1,90);const state=ws.capturarEstado(),before=await ws.obterBytes();const preview=await ws.prepararRedimensionamento(config);assert.deepEqual(ws.capturarEstado(),state);assert.deepEqual(await ws.obterBytes(),before);ws.girarPagina(1,90);assert.throws(()=>ws.aplicarRedimensionamento(preview),/vencida/);assert.equal(ws.alterado,true);
});
test('formulário, ação adicional e assinaturas não são contornados pelo caminho marcado',async()=>{
    for(const key of ['AA','AcroForm']){const ws=await abrir();ws.pdf.catalog.set(N(key),ws.pdf.context.obj(key==='AA'?{O:{S:'JavaScript'}}:{Fields:[ws.pdf.context.register(ws.pdf.context.obj({FT:'Tx',T:'synthetic'}))]}));await assert.rejects(ws.prepararRedimensionamento(config));}
});
