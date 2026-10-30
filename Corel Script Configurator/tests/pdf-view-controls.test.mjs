import test from "node:test";
import assert from "node:assert/strict";
import { setTimeout as aguardar } from "node:timers/promises";
import { Elemento } from "./helpers/pdf-dom.mjs";
import { PDFControlesVisualizacao } from "../public/js/tools/pdf-view-controls.js";

function iniciar() {
    const elementos = new Map(["visualizacao", "zoom-valor", "zoom-pagina", "zoom-largura", "zoom-real", "zoom-mais", "zoom-menos"].map(id => ["#pdf-" + id, new Elemento("div")]));
    const container = elementos.get("#pdf-visualizacao");
    container.clientWidth = 632; container.clientHeight = 432;
    const previa = {}, estado = { chamadas: 0, permitido: true };
    const controles = new PDFControlesVisualizacao(id => elementos.get(id), previa, async () => { estado.chamadas++; }, () => estado.permitido);
    return { elementos, container, previa, estado, controles };
}
test("ajuste usa largura e altura disponíveis; 100% e zoom manual não alteram o PDF", () => {
    const { previa, controles } = iniciar();
    try {
        assert.equal(previa.calcularEscala({ width: 300, height: 800 }), .5);
        assert.equal(previa.calcularEscala({ width: 1200, height: 300 }), .5);
        controles.modo = "largura"; assert.equal(previa.calcularEscala({ width: 300, height: 800 }), 2);
        controles.modo = "real"; controles.zoom = 1; assert.equal(previa.calcularEscala({ width: 300, height: 800 }), 1);
        controles.modo = "manual"; controles.zoom = 6;
        assert.ok(previa.calcularEscala({ width: 5000, height: 5000 }) ** 2 * 25000000 <= 16000000 + 1e-6);
    } finally { controles.destruir(); }
});
test("roda somente no centro, agrupamento de eventos e liberação de listeners/timer", async () => {
    const { container, previa, controles, estado } = iniciar();
    for (let i = 0; i < 10; i++) assert.equal(container.emitir("wheel", container, { deltaY: -30, deltaMode: 0, clientX: 20, clientY: 20 }).defaultPrevented, true);
    await aguardar(100); assert.equal(estado.chamadas, 1);
    assert.equal(controles.modo, "manual");
    estado.permitido = false;
    assert.equal(container.emitir("wheel", container, { deltaY: 50 }).defaultPrevented, false);
    estado.permitido = true; controles.ampliar(2); controles.destruir(); await aguardar(100);
    assert.equal(estado.chamadas, 1); assert.equal(previa.calcularEscala, undefined);
    assert.ok([...container.listeners.values()].every(lista => lista.length === 0));
});
