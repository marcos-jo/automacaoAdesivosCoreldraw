import assert from "node:assert/strict";
import test from "node:test";
import { AbasAplicacao } from "../public/js/app/tabs.js";

function iniciar() {
    const criadas = [], removidas = [], erros = [], eventos = [];
    const interfacePrincipal = { inicio: {}, criarPainel: id => ({ id }), removerPainel: painel => removidas.push(painel.id), informar: e => erros.push(e), exibir(aba) { this.ativa = aba.id; } };
    const modulos = ['pdfs','adesivos','qrcodes'].map(id => ({ id, titulo:id, montar() {
        const controle = { estado: { documento:'original', selecao:[2,5], zoom:2 },
            suspender() { eventos.push('pausar '+id); }, retomar() { eventos.push('retomar '+id); },
            destruir() { eventos.push('destruir '+id); }, confirmarFechamento: async () => true };
        criadas.push(controle); return controle;
    } }));
    const abas = new AbasAplicacao(modulos, interfacePrincipal);
    return { abas, modulos, criadas, removidas, erros, eventos, interfacePrincipal };
}
test("Início não fecha; abrir/alternar reutiliza a mesma instância e preserva o estado", async () => {
    const { abas, criadas, eventos } = iniciar();
    assert.equal(await abas.fechar('inicio'), false);
    await abas.ativar('pdfs'); const controle = criadas[0], estado = controle.estado;
    await abas.ativar('inicio'); await abas.ativar('pdfs'); await abas.ativar('pdfs');
    assert.equal(criadas.length, 1); assert.equal(abas.abas.get('pdfs').controle.estado, estado);
    assert.deepEqual(eventos, ['retomar pdfs','pausar pdfs','retomar pdfs']);
    assert.deepEqual([...abas.abas.keys()], ['inicio','pdfs']);
});
test("fechar pede confirmação na aba visível; cancelar preserva e confirmar desmonta uma vez", async () => {
    const { abas, criadas, removidas, eventos } = iniciar();
    await abas.ativar('pdfs'); await abas.ativar('qrcodes');
    criadas[0].confirmarFechamento = async () => { assert.equal(abas.ativa,'pdfs'); return false; };
    assert.equal(await abas.fechar('pdfs'), false); assert.equal(removidas.length, 0);
    criadas[0].confirmarFechamento = async () => true;
    assert.equal(await abas.fechar('pdfs'), true); assert.deepEqual(removidas,['pdfs']);
    assert.equal(eventos.filter(e => e==='destruir pdfs').length,1); assert.equal(abas.ativa,'inicio');
    await abas.ativar('pdfs'); assert.equal(criadas.length,3);
});
test("ações concorrentes são serializadas sem duplicar ferramentas", async () => {
    const { abas, criadas } = iniciar();
    await Promise.all([abas.ativar('pdfs'),abas.ativar('qrcodes'),abas.ativar('pdfs')]);
    assert.equal(criadas.length,2); assert.equal(abas.ativa,'pdfs');
});
test("modal bloqueia alternância e falha de montagem não remove o trabalho anterior", async () => {
    const { abas, criadas, modulos } = iniciar(); await abas.ativar('pdfs');
    criadas[0].podeAlternar = () => false;
    assert.equal(await abas.ativar('inicio'),false); assert.equal(abas.ativa,'pdfs');
    assert.equal(await abas.fechar('pdfs'),false);
    criadas[0].podeAlternar = () => true; modulos[1].montar = () => { throw new Error('Falha controlada'); };
    await assert.rejects(abas.ativar('adesivos'), /Falha controlada/); assert.equal(abas.ativa,'pdfs');
    await abas.ativar('inicio'); assert.equal(abas.ativa,'inicio');
});
