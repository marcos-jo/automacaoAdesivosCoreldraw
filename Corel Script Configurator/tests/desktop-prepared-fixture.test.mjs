import test from 'node:test';
import assert from 'node:assert/strict';
import {PDFDocument,PDFName} from '../public/vendor/pdf-lib/pdf-lib.js';
import {PDFWorkspace,assinaturasUrl} from './helpers/pdf-resize-modules.mjs';
import {criarPDFPreparadoSintetico} from './helpers/pdf-print-prepared-fixture.mjs';
const {verificarSemAssinaturas}=await import(assinaturasUrl);

test('fixture desktop preparada abre sem assinatura, mantém vetor/imagem e permite edição/exportação',async()=>{
    const bytes=await criarPDFPreparadoSintetico(),pdf=await PDFDocument.load(bytes);
    verificarSemAssinaturas(pdf);
    assert.equal(pdf.catalog.has(PDFName.of('AcroForm')),false);
    assert.equal(pdf.getPageCount(),2);
    assert.ok(pdf.getPage(0).node.Resources().has(PDFName.of('Font')));
    const imagens=pdf.getPage(1).node.Resources().lookup(PDFName.of('XObject'));
    assert.equal(imagens.keys().length,1);
    assert.equal(imagens.lookup(imagens.keys()[0]).dict.lookup(PDFName.of('Subtype')).asString(),'/Image');
    const ws=new PDFWorkspace();await ws.abrir({name:'preparada.pdf',arrayBuffer:async()=>bytes.slice().buffer});
    assert.equal(ws.obterResumoAssinaturas().presentes,0);
    assert.equal(ws.obterResumoAssinaturas().preparada,false,'Não inventa histórico de preparo após importação');
    ws.girarPagina(1,90);
    const exportado=await PDFDocument.load(await ws.obterBytes());
    verificarSemAssinaturas(exportado);
    assert.equal(exportado.getPage(0).getRotation().angle,90);
    assert.equal(exportado.getPageCount(),2);
});
