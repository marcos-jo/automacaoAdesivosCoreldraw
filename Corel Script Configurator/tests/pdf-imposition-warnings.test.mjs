import test from 'node:test';import assert from 'node:assert/strict';
import {agruparAvisosMontagem,intervalosFolhas,textoAvisosMontagem} from '../public/js/tools/pdf-imposition-warnings.js';
test('intervalos preservam todas as folhas, sem truncar ou duplicar',()=>{assert.equal(intervalosFolhas([5,1,3,2,3,8,9]),'1–3, 5, 8–9');assert.equal(intervalosFolhas([]),'');});
for(const count of [1,3,24])test(`avisos de ${count} folhas: agrupa papel/margens sem perder ocorrências`,()=>{
    const geometry={sheets:Array.from({length:count},()=>({foraFolha:['esquerda','direita'],foraMargens:['esquerda','direita'],foraImprimivel:null}))};
    const before=structuredClone(geometry),m=agruparAvisosMontagem(geometry);assert.equal(m.afetadas,count);assert.equal(m.grupos.length,2);assert.ok(m.resumo.length<120);assert.deepEqual(geometry,before);
    for(const group of m.grupos){assert.equal(group.ocorrencias.length,count);assert.equal(group.padroes.length,1);assert.deepEqual(group.padroes[0].folhas,Array.from({length:count},(_,i)=>i+1));}
    assert.match(textoAvisosMontagem(m),/bordas esquerda, direita/);assert.ok(textoAvisosMontagem(m).includes(count===1?'folhas 1 —':`folhas 1–${count} —`));
});
test('tipos e bordas diferentes conservam o diagnóstico de cada folha',()=>{
    const m=agruparAvisosMontagem({sheets:[{foraFolha:['esquerda'],foraMargens:['superior'],foraImprimivel:['direita']},{foraFolha:['direita'],foraMargens:[],foraImprimivel:null},{foraFolha:[],foraMargens:['superior'],foraImprimivel:['direita']}]});
    assert.equal(m.afetadas,3);assert.equal(m.grupos.length,3);assert.deepEqual(m.grupos[0].ocorrencias,[{folha:1,bordas:['esquerda']},{folha:2,bordas:['direita']}]);assert.deepEqual(m.grupos[1].padroes[0].folhas,[1,3]);
});
test('corrigir conflitos remove o resumo sem inventar área imprimível',()=>{assert.equal(agruparAvisosMontagem({sheets:[{foraFolha:[],foraMargens:[],foraImprimivel:null}]}).afetadas,0);assert.equal(agruparAvisosMontagem().resumo,'');});
