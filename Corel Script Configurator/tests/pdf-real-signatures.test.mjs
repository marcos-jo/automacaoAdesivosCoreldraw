import assert from "node:assert/strict";
import test from "node:test";
import { readFile, access } from "node:fs/promises";
import { createHash } from "node:crypto";
import { PDFDocument } from "../public/vendor/pdf-lib/pdf-lib.js";
import { PDFWorkspace, assinaturasUrl } from "./helpers/pdf-resize-modules.mjs";
import { criarConjuntoAssinaturasReais, verificarAssinaturasCMS } from "./helpers/pdf-real-signatures.mjs";
const { diagnosticarAssinaturas } = await import(assinaturasUrl);
const conjunto = await criarConjuntoAssinaturasReais();
const sha = b => createHash("sha256").update(b).digest("hex");

test("seis documentos de teste têm CMS reais; múltiplas assinaturas mantêm todas as revisões verificáveis", async () => {
    assert.equal(conjunto.casos.length, 6);
    let total = 0;
    for (const caso of conjunto.casos) {
        const d = diagnosticarAssinaturas(await PDFDocument.load(caso.bytes), { tamanhoArquivo: caso.bytes.length });
        assert.equal(d.presentes, caso.vazio ? 0 : caso.paginas.length, caso.id);
        assert.equal(caso.verificacoes.length, caso.vazio ? 0 : caso.paginas.length);
        assert.ok(caso.verificacoes.every(v => v.cmsVerificado));
        if (caso.paginas.length > 1) assert.ok(caso.verificacoes.some(v => v.revisaoAte < caso.bytes.length), "A primeira assinatura deve continuar válida sobre sua revisão anterior");
        total += caso.verificacoes.length;
    }
    assert.equal(total, 7);
    await assert.rejects(access(conjunto.chave), { code: "ENOENT" }, "A chave privada efêmera deve ser eliminada após a geração");
    console.log("CMS reais: 7 assinaturas em 5 PDFs + 1 PDF de campo vazio; OpenSSL verificou digest/assinatura, sem validar confiança da cadeia. Artefatos: " + conjunto.pasta);
});

test("alterar um byte protegido do PDF faz a verificação CMS falhar", async () => {
    const caso = conjunto.casos[0], alterado = Buffer.from(caso.bytes);
    alterado[7] = alterado[7] === 0x37 ? 0x36 : 0x37; // versão no header: PDF continua legível
    await PDFDocument.load(alterado);
    await assert.rejects(verificarAssinaturasCMS(alterado, conjunto, "adulterado"));
    assert.equal((await verificarAssinaturasCMS(caso.bytes, conjunto, "original-preservado")).length, 1);
});

test("reescrita do workspace não é confundida com o original assinado; o original continua criptograficamente verificável", async () => {
    const caso = conjunto.casos[0], ws = new PDFWorkspace();
    await ws.abrir({ name: caso.id + ".pdf", arrayBuffer: async () => Uint8Array.from(caso.bytes).buffer });
    const d = ws.obterResumoAssinaturas(); assert.equal(d.presentesPendentes, 1); assert.equal(ws.alterado, false);
    await assert.rejects(verificarAssinaturasCMS(await ws.obterBytes(), conjunto, "reescrito"));
    assert.equal(ws.obterResumoAssinaturas().presentesPendentes, 1, "Diagnóstico da origem permanece disponível");
    assert.equal(sha(await readFile(caso.arquivo)), caso.sha256);
    assert.equal((await verificarAssinaturasCMS(await readFile(caso.arquivo), conjunto, "origem-apos-workspace")).length, 1);
});
