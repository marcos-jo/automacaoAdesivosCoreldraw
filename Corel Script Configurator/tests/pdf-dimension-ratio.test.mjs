import test from "node:test";
import assert from "node:assert/strict";
import { Elemento } from "./helpers/pdf-dom.mjs";
import { PDFProporcaoMedidas, formatarProporcao } from "../public/js/tools/pdf-dimension-ratio.js";
function iniciar() {
    const campos = Object.fromEntries(['largura', 'altura', 'proporcao', 'vincular'].map(id => [id,new Elemento('input')]));
    campos.largura.value = '210'; campos.altura.value = '297';
    const proporcao = new PDFProporcaoMedidas(id => campos[id.replace('#pdf-redim-', '')]);
    return { campos, proporcao };
}

test("indicador normaliza a menor dimensão em 1, com arredondamento estável e inteiros sem zeros extras", () => {
    for (const [a, b, esperado] of [
        [210, 297, '210 × 297 mm · ≈ 1:1,41'],
        [297, 210, '297 × 210 mm · ≈ 1,41:1'],
        [400, 200, '400 × 200 mm · ≈ 2:1'],
        [200, 200, '200 × 200 mm · ≈ 1:1'],
        [100.5, 150.25, '100,5 × 150,25 mm · ≈ 1:1,5'],
        [.1, 5000, '0,1 × 5000 mm · ≈ 1:50000']
    ]) assert.equal(formatarProporcao(a, b), esperado);
});

test("cadeado expõe estados distintos sem substituir os ícones nem alterar as medidas ao vincular", () => {
    const { campos:c, proporcao:p } = iniciar();
    try {
        assert.equal(c.vincular.getAttribute('aria-label'), 'Proporção livre');
        assert.equal(c.vincular.getAttribute('aria-pressed'), 'false');
        c.vincular.textContent = 'ícones preservados';
        c.largura.value = '123,4567890123'; c.altura.value = '234,5678901234';
        c.vincular.click();
        assert.equal(c.vincular.getAttribute('aria-label'), 'Proporção vinculada');
        assert.equal(c.vincular.getAttribute('aria-pressed'), 'true');
        assert.equal(c.vincular.textContent, 'ícones preservados');
        assert.equal(c.largura.value, '123,4567890123'); assert.equal(c.altura.value, '234,5678901234');
        assert.equal(c.proporcao.textContent, '123,46 × 234,57 mm · ≈ 1:1,9');
        c.vincular.click(); assert.equal(c.vincular.getAttribute('aria-label'), 'Proporção livre');
    } finally { p.destruir(); }
});
test("cadeado recalcula ambas as medidas com vírgula sem arredondar formatos salvos", () => {
    const { campos:c, proporcao:p } = iniciar(); c.vincular.click(); c.largura.value='420'; p.alterar('largura');
    assert.equal(Number(c.altura.value.replace(',','.')), 594);
    c.altura.value='148,5'; p.alterar('altura'); assert.equal(Number(c.largura.value.replace(',','.')), 105);
    c.vincular.click(); c.largura.value='123,4567890123'; c.altura.value='234,5678901234'; p.sincronizar();
    c.vincular.click(); assert.equal(c.largura.value,'123,4567890123'); assert.equal(c.altura.value,'234,5678901234');
    p.destruir(); assert.equal(c.vincular.listeners.get('click').length, 0);
});
test("orientação/formato atualizam vínculo; medida inválida bloqueia configuração sem corromper a outra", () => {
    const { campos:c, proporcao:p } = iniciar(); c.vincular.click();
    c.largura.value='297'; c.altura.value='210'; p.sincronizar(); c.largura.value='594'; p.alterar('largura'); assert.equal(c.altura.value,'420');
    c.largura.value='5000'; p.alterar('largura'); p.validar();
    c.altura.value='5000'; p.alterar('altura'); assert.throws(() => p.validar());
    c.altura.value='210'; p.alterar('altura'); p.validar(); assert.equal(c.largura.value,'297'); p.destruir();
});
