import assert from "node:assert/strict";
import test from "node:test";
import { setImmediate as turno } from "node:timers/promises";
import { Elemento } from "./helpers/pdf-dom.mjs";
import { JanelaExportacao, nomeExportacao } from "../public/js/app/export-dialog.js";
test("nomes mantêm modalidade PDF e sufixo IMPRESSAO, rejeitando caminhos e nomes vazios", () => {
    assert.equal(nomeExportacao('proposta.pdf'), 'proposta.pdf');
    assert.equal(nomeExportacao('proposta',true), 'proposta_IMPRESSAO.pdf');
    assert.equal(nomeExportacao('proposta_COPIA_IMPRESSAO.pdf',true), 'proposta_COPIA_IMPRESSAO.pdf');
    for (const nome of ['', '..', 'x/y', 'x\\y', 'arquivo?']) assert.throws(() => nomeExportacao(nome));
});
test("abrir/cancelar não exporta; confirmação, falha e retry mantêm o dono e o nome corretos", async () => {
    globalThis.document = { createElement: tag => new Elemento(tag) };
    const root = new Elemento('main'), ui = new JanelaExportacao(root), dono = {};
    let chamadas = [];
    try {
        const dados = { nome:'teste.pdf', modalidade:'Documento de trabalho' };
        ui.abrir(dados,async n => chamadas.push(n),dono); assert.equal(chamadas.length,0);
        ui.cancelar(); assert.equal(ui.aberta,false); assert.equal(chamadas.length,0);
        ui.abrir(dados,async () => {throw new Error('Falha controlada');},dono);
        await ui.confirmar(); assert.equal(ui.aberta,true); assert.match(ui.campo('status').textContent,/Falha controlada/);
        assert.equal(ui.dono,dono); ui.exportar=async n=>chamadas.push(n);
        ui.campo('nome').value='final'; await ui.confirmar(); assert.deepEqual(chamadas,['final.pdf']); assert.equal(ui.aberta,false);
    } finally { ui.destruir(); delete globalThis.document; }
});
test("cancelar durante exportação sinaliza o serviço e impede reabertura até liberar a operação", async () => {
    globalThis.document = { createElement: tag => new Elemento(tag) };
    const root = new Elemento('main'), ui = new JanelaExportacao(root); let liberar, sinal;
    try {
        const espera = new Promise(r=>liberar=r);
        ui.abrir({nome:'teste.pdf'},async (_, {signal}) => {sinal=signal;await espera;},root);
        const tarefa = ui.confirmar(); await turno(); ui.cancelar();
        assert.equal(sinal.aborted,true); assert.equal(ui.abrir({nome:'outro.pdf'},async()=>{},{}),false);
        liberar(); await tarefa; assert.equal(ui.dono,null); assert.equal(ui.ocupada,false);
    } finally { ui.destruir(); delete globalThis.document; }
});
