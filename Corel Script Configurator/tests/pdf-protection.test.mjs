import assert from "node:assert/strict";
import test from "node:test";
import { Elemento, JanelaTeste } from "./helpers/pdf-dom.mjs";
import { renderPDFInterface } from "../public/js/tools/pdf-interface.js";
import { PDFProtecao } from "../public/js/tools/pdf-protection.js";

function montar(pendente = false, janela = new JanelaTeste()) {
    const app = new Elemento("main");
    const raiz = renderPDFInterface(app);
    const workspace = { alterado: pendente };
    const protecao = new PDFProtecao(raiz, workspace, janela);
    return { protecao, workspace, janela, buscar: id => raiz.querySelector("#" + id) };
}

test("sem alterações não abre confirmação nem instala beforeunload", async () => {
    const ui = montar();
    try {
        assert.equal(await ui.protecao.confirmarDescarte(), true);
        assert.ok(!ui.buscar("pdf-confirmar-descarte").open);
        assert.equal(ui.buscar("pdf-alteracoes-pendentes").hidden, true);
        assert.equal(ui.janela.listeners.get("beforeunload")?.size || 0, 0);
    } finally { ui.protecao.destruir(); }
});

test("continuar, Escape e fechamento do diálogo preservam as pendências", async () => {
    const ui = montar(true);
    try {
        for (const acao of ["botao", "cancel", "close"]) {
            const resposta = ui.protecao.confirmarDescarte();
            assert.equal(ui.protecao.confirmarDescarte(), resposta, "Apenas uma confirmação por vez");
            assert.equal(ui.buscar("pdf-confirmar-descarte").open, true);
            if (acao === "botao") ui.buscar("pdf-continuar-editando").click();
            else ui.buscar("pdf-confirmar-descarte").emitir(acao);
            assert.equal(await resposta, false);
            assert.equal(ui.workspace.alterado, true);
            assert.equal(ui.buscar("pdf-confirmar-descarte").open, false);
        }
    } finally { ui.protecao.destruir(); }
});

test("descarte autoriza o chamador a sair e encerramento remove os listeners", async () => {
    const ui = montar(true);
    const resposta = ui.protecao.confirmarDescarte();
    ui.buscar("pdf-descartar-sair").click();
    assert.equal(await resposta, true);
    ui.protecao.destruir();
    assert.equal(ui.janela.listeners.get("beforeunload").size, 0);
    assert.equal(ui.buscar("pdf-continuar-editando").listeners.get("click").length, 0);
});

test("beforeunload e indicador seguem o estado sem acumular listeners em reentradas", async () => {
    const janela = new JanelaTeste();
    for (let i = 0; i < 4; i++) {
        const ui = montar(true, janela);
        ui.protecao.atualizar();
        ui.protecao.atualizar();
        assert.equal(janela.listeners.get("beforeunload").size, 1);
        const evento = new Event("beforeunload", { cancelable: true });
        // Node não possui BeforeUnloadEvent; simula seu returnValue gravável.
        Object.defineProperty(evento, "returnValue", { value: undefined, writable: true });
        janela.dispatchEvent(evento);
        assert.equal(evento.defaultPrevented, true);
        assert.equal(evento.returnValue, "");
        ui.workspace.alterado = false;
        ui.protecao.atualizar();
        assert.equal(janela.listeners.get("beforeunload").size, 0);
        assert.equal(ui.buscar("pdf-alteracoes-pendentes").hidden, true);
        assert.equal(await ui.protecao.confirmarDescarte(), true);
        ui.protecao.destruir();
    }
});
