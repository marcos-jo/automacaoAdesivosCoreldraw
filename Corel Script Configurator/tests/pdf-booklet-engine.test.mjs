import test from 'node:test';
import assert from 'node:assert/strict';
import {PDFDocument,PDFName,decodePDFRawStream} from '../public/vendor/pdf-lib/pdf-lib.js';
import {gerarMontagem} from './helpers/pdf-imposition-modules.mjs';
import {criarOrigemLivreto} from './helpers/pdf-booklet-fixtures.mjs';
import {conferirTransformacoes} from './helpers/pdf-imposition-proof.mjs';
for(const count of [4,8,12,16,5,7,10])test(`PDF vetorial ${count}: faces, brancos, caixas e matrizes serializadas`,async()=>{
    const bytes=await criarOrigemLivreto(count),original=bytes.slice(),r=await gerarMontagem(bytes,{modo:'livreto'},{previsualizarExcesso:true});
    assert.deepEqual(bytes,original);assert.deepEqual(r.plan.bloqueios,[]);
    assert.equal(await conferirTransformacoes(r.bytes,r.plan),count);
    assert.equal(await conferirTransformacoes(r.bytesLeitura,r.plan.leitura),count);
    const pdf=await PDFDocument.load(r.bytes),preview=await PDFDocument.load(r.bytesPrevia);
    assert.equal(pdf.getPageCount(),Math.ceil(count/4)*2);
    for(const [i,p]of pdf.getPages().entries()){
        const contents=(p.node.Contents()?.asArray()||[]).map(ref=>Buffer.from(decodePDFRawStream(pdf.context.lookup(ref)).decode()).toString()).join('\n');
        assert.equal((contents.match(/\bW\b/g)||[]).length,r.plan.folhas[i].length,'clip vetorial, sem imagem de página');
        assert.ok(preview.getPage(i).getWidth()>p.getWidth());
        for(const ref of p.node.Resources()?.lookup(PDFName.of('XObject'))?.values()||[]){const form=pdf.context.lookup(ref);assert.equal(form.dict.get(PDFName.of('Subtype')).toString(),'/Form');assert.ok(form.dict.lookup(PDFName.of('Resources')).lookup(PDFName.of('Font')).keys().length);}
    }
});
for(const orientacao of ['paisagem','retrato'])for(const verso of ['normal','girar180'])test(`PDF com origem girada: ${orientacao}/${verso}`,async()=>{
    const bytes=await criarOrigemLivreto(8,{rotation:90,largura:210,altura:148});
    const r=await gerarMontagem(bytes,{modo:'livreto',orientacao,verso});
    assert.equal(await conferirTransformacoes(r.bytes,r.plan),8);assert.equal(await conferirTransformacoes(r.bytesLeitura,r.plan.leitura),8);
});
test('geração direta bloqueia corte de conteúdo e aceita ajuste solicitado',async()=>{
    const bytes=await criarOrigemLivreto(4,{largura:210,altura:297});
    await assert.rejects(gerarMontagem(bytes,{modo:'livreto'}),/conteúdo excede/);
    const r=await gerarMontagem(bytes,{modo:'livreto',escalaModo:'ajustar'});assert.deepEqual(r.plan.bloqueios,[]);assert.equal(await conferirTransformacoes(r.bytes,r.plan),4);
});
test('cancelamento continua efetivo no gerador comum',async()=>{const a=new AbortController();a.abort();await assert.rejects(gerarMontagem(await criarOrigemLivreto(4),{modo:'livreto'},{signal:a.signal}),{name:'AbortError'});});
