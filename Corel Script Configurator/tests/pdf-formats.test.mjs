import assert from "node:assert/strict";
import test from "node:test";
import { PDFFormatos, medidaMm } from "../public/js/tools/pdf-formats.js";
import { PDFArmazenamentoFormatos } from "../public/js/tools/pdf-format-storage.js";

function armazenamento() {
    let valor = null;
    return { getItem: () => valor, setItem: (_, novo) => { valor = novo; } };
}
test("formatos: cadastrar, recuperar, editar, renomear e excluir sem alterar padrões", async () => {
    const local = armazenamento(), adaptador = new PDFArmazenamentoFormatos(() => local, "teste-formatos");
    const formatos = new PDFFormatos(adaptador);
    const f = await formatos.salvar({ nome: "Formato 4", largura: "123,45", altura: "234.56" });
    assert.equal(new PDFFormatos(adaptador).listar().at(-1).largura, 123.45);
    await formatos.salvar({ nome: "Convite", largura: "140", altura: "180" }, f.id);
    assert.equal(new PDFFormatos(adaptador).listar().at(-1).nome, "Convite");
    await assert.rejects(formatos.salvar({ nome: "convite", largura: 1, altura: 1 }), /existe/);
    await assert.rejects(formatos.salvar({ nome: " a4 ", largura: 1, altura: 1 }), /existe/);
    await assert.rejects(formatos.excluir("A4"), /personalizados/);
    await assert.rejects(formatos.salvar({ nome: "novo", largura: 1, altura: 1 }, "A4"), /personalizados/);
    await formatos.excluir(f.id);
    assert.equal(new PDFFormatos(adaptador).listar().length, 4);
});
test("formatos: validar medidas, nomes e entradas incompletas", async () => {
    const formatos = new PDFFormatos(new PDFArmazenamentoFormatos(() => armazenamento(), "validacao"));
    for (const valor of ["", "0", "-2", "0,01", "5001", "1.234,5", "Infinity", "1e3", "2mm"]) assert.throws(() => medidaMm(valor));
    assert.equal(medidaMm("0,1"), .1);
    assert.equal(medidaMm("5000"), 5000);
    await assert.rejects(formatos.salvar({ nome: "", largura: 10, altura: 20 }));
    await assert.rejects(formatos.salvar({ nome: "Teste", largura: 10 }));
});
test("formatos: dados corrompidos/versão desconhecida e armazenamento indisponível", async () => {
    for (const valor of ["{", '{"versao":2,"formatos":[]}', '{"versao":1,"formatos":[{}]}']) {
        const formatos = new PDFFormatos({ ler: () => valor });
        assert.equal(formatos.listar().length, 4);
        assert.match(formatos.obterAviso(), /inválida/);
    }
    const indisponivel = new PDFArmazenamentoFormatos(() => { throw new Error("Bloqueado"); }, "falha-storage");
    const formatos = new PDFFormatos(indisponivel);
    await formatos.salvar({ nome: "Sessão", largura: 100, altura: 200 });
    assert.match(formatos.obterAviso(), /sessão/);
    assert.equal(new PDFFormatos(indisponivel).listar().at(-1).nome, "Sessão");
});

test("formatos: contrato também aceita adaptador assíncrono sem alterar a ferramenta", async () => {
    let valor = '{"versao":1,"formatos":[]}';
    const adaptador = { ler: async () => valor, gravar: async json => { valor = json; } };
    const formatos = new PDFFormatos(adaptador); await formatos.pronta;
    await formatos.salvar({ nome: "API simulada", largura: 130, altura: 200 });
    const outraInstancia = new PDFFormatos(adaptador); await outraInstancia.pronta;
    assert.equal(outraInstancia.listar().at(-1).nome, "API simulada");
});

test("formatos: quota esgotada não perde cadastros ao reentrar com leitura ainda disponível", async () => {
    const local = { getItem: () => null, setItem: () => { throw new Error("QuotaExceededError"); } };
    const formatos = new PDFFormatos(new PDFArmazenamentoFormatos(() => local, "quota"));
    await formatos.salvar({ nome: "Temporário", largura: 100, altura: 200 });
    const reentrada = new PDFFormatos(new PDFArmazenamentoFormatos(() => local, "quota"));
    assert.equal(reentrada.listar().at(-1).nome, "Temporário");
    assert.match(reentrada.obterAviso(), /sessão/);
});
