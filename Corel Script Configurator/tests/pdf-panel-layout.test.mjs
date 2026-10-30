import test from "node:test";
import assert from "node:assert/strict";
import { PDFPreferenciasPaineis } from "../public/js/tools/pdf-panel-layout.js";
test("preferências versionadas persistem, limitam valores e ignoram dados inválidos", () => {
    let dado = null; const armazenamento = { ler: () => dado, gravar: (_, valor) => { dado = valor; } };
    const p = new PDFPreferenciasPaineis(armazenamento); p.salvar('esquerda', 200); p.salvar('direita', 2000);
    assert.deepEqual(new PDFPreferenciasPaineis(armazenamento).valor, { esquerda:200, direita:420 });
    for (const invalido of ['{', '{"versao":2,"esquerda":200}', '{"versao":1,"esquerda":"200","direita":-10}']) {
        dado = invalido; assert.deepEqual(new PDFPreferenciasPaineis(armazenamento).valor, { esquerda:150, direita:230 });
    }
});
test("armazenamento negado conserva as preferências da sessão", () => {
    const p = new PDFPreferenciasPaineis({ ler() {throw new Error();}, gravar() {throw new Error();} });
    p.salvar('esquerda', 220); assert.equal(p.valor.esquerda, 220); p.salvar('esquerda', NaN); assert.equal(p.valor.esquerda, 220);
});
