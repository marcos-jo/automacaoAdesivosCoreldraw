import test from "node:test";
import assert from "node:assert/strict";
import { agruparDimensoes, intervalosPaginas } from "../public/js/tools/pdf-dimension-groups.js";
test("350 páginas equivalentes produzem um grupo e intervalos compactos", () => {
    const paginas = Array.from({ length:350 }, (_, i) => ({numero:i+1, largura:210 + i % 2 * .01, altura:297}));
    const r = agruparDimensoes(paginas); assert.equal(r.grupos.length, 1); assert.equal(r.grupos[0].nome, 'A4');
    assert.equal(intervalosPaginas(r.grupos[0].paginas), '1–350'); assert.equal(r.variados, false);
});
test("orientação e giro são distinguidos sem classificar A4 girado como novo formato", () => {
    const r = agruparDimensoes([
        {numero:1,largura:210,altura:297,rotacao:0},
        {numero:2,largura:297,altura:210,larguraBase:210,alturaBase:297,rotacao:90},
        {numero:3,largura:297,altura:420,rotacao:0}
    ]);
    assert.equal(r.grupos.length, 3); assert.equal(r.variados, true); assert.equal(r.predominante.total, 2);
    assert.equal(r.grupos[0].familia, r.grupos[1].familia); assert.deepEqual(r.grupos[1].giradas, [2]);
    assert.notEqual(r.grupos[2].familia, r.predominante);
    assert.equal(r.orientacoesDiferentes, true);
    assert.deepEqual(r.orientacoes, { retrato:[1,3], paisagem:[2], quadrada:[] });
});

for (const [retratos, paisagens] of [[10,0], [0,10], [9,1], [1,9]]) {
    test(`${retratos} retrato / ${paisagens} paisagem: alerta independe da maioria e não cria outro formato`, () => {
        const r = agruparDimensoes(Array.from({length:retratos+paisagens}, (_, i) => ({
            numero:i+1, largura:i<retratos?210:297, altura:i<retratos?297:210
        })));
        assert.equal(r.variados, false);
        assert.equal(r.orientacoesDiferentes, !!(retratos && paisagens));
        assert.equal(r.avisos.includes("Orientações diferentes detectadas"), !!(retratos && paisagens));
        assert.equal(r.orientacoes.retrato.length, retratos);
        assert.equal(r.orientacoes.paisagem.length, paisagens);
        assert.match(r.detalhesOrientacoes, new RegExp(`${retratos || paisagens} página\\(s\\)`));
    });
}
test("tamanhos e orientações são avisos independentes, com números compactos no documento atual", () => {
    const r = agruparDimensoes([
        {numero:1,largura:210,altura:297}, {numero:2,largura:297,altura:210,rotacao:270},
        {numero:3,largura:297,altura:420}, {numero:4,largura:297,altura:420}
    ]);
    assert.equal(r.avisos.length, 2);
    assert.equal(r.detalhesOrientacoes, "Retrato: 3 página(s) — páginas 1, 3–4. Paisagem: 1 página(s) — páginas 2.");
});
test("páginas quadradas dentro da tolerância não inventam diferença retrato/paisagem", () => {
    const r = agruparDimensoes([{numero:1,largura:210,altura:297}, {numero:2,largura:210,altura:210.01}]);
    assert.equal(r.orientacoesDiferentes, false);
    assert.deepEqual(r.orientacoes.quadrada, [2]);
    assert.equal(r.grupos[1].orientacao, "quadrada");
});
