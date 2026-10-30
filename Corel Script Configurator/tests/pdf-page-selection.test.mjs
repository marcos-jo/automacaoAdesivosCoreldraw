import test from "node:test";
import assert from "node:assert/strict";
import { PDFSelecaoPaginas, ordemMoverPaginas } from "../public/js/tools/pdf-page-selection.js";

test("clique/Shift/Ctrl usam identidades e sobrevivem à janela virtual e reordenação", () => {
    const s = new PDFSelecaoPaginas(), ids = Array.from({ length: 350 }, (_, i) => 'id' + i);
    s.atualizar(ids); s.clicar(5); s.clicar(50, { shiftKey: true });
    assert.equal(s.numeros().length, 46);
    s.clicar(7, { ctrlKey: true }); assert.equal(s.ids.has(ids[6]), false);
    s.clicar(300, { ctrlKey: true }); assert.equal(s.ids.has(ids[299]), true);
    const ordem = ordemMoverPaginas(350, s.numeros(), 200);
    const antes = new Set(s.ids); s.atualizar(ordem.map(n => ids[n - 1]));
    assert.deepEqual(s.ids, antes); assert.equal(s.numeros()[0], 200);
});
test("arraste de grupo mantém ordem relativa e rejeita destino/seleção inválidos", () => {
    assert.deepEqual(ordemMoverPaginas(7, [2, 5], 4), [1, 3, 4, 2, 5, 6, 7]);
    assert.deepEqual(ordemMoverPaginas(7, [2, 5], null, 6), [1, 3, 4, 6, 2, 5, 7]);
    assert.deepEqual(ordemMoverPaginas(7, [2, 3], null, 3), [1, 2, 3, 4, 5, 6, 7]);
    assert.throws(() => ordemMoverPaginas(7, [2, 8], 1));
    assert.throws(() => ordemMoverPaginas(7, [2, 5], 7));
});
test("exclusão filtra identidades, união mantém seleção e troca de documento reinicia", () => {
    const s = new PDFSelecaoPaginas(); s.atualizar(['a', 'b', 'c']); s.clicar(2); s.clicar(3, { ctrlKey: true });
    s.atualizar(['b', 'c', 'd']); assert.deepEqual(s.numeros(), [1, 2]);
    s.atualizar(['d']); assert.deepEqual(s.numeros(), [1]);
    s.limpar(); s.atualizar(['a', 'b']); assert.deepEqual(s.numeros(), [1]);
});
