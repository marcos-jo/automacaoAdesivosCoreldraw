import test from "node:test";
import assert from "node:assert/strict";
import { PDFDocument, PDFName, PDFHexString } from "../public/vendor/pdf-lib/pdf-lib.js";
import { assinaturasUrl } from "./helpers/pdf-resize-modules.mjs";
import { criarPDFAssinado, adicionarAssinatura } from "./helpers/pdf-signature-fixture.mjs";
const { diagnosticarAssinaturas, verificarSemAssinaturas } = await import(assinaturasUrl);
const nome = PDFName.of;
test("distingue presença estrutural, campo vazio, assinatura invisível e páginas visuais", async () => {
    for (const opcoes of [{}, { vazia: true }, { invisivel: true }, { paginas: [1, 3] }]) {
        const bytes = await criarPDFAssinado(opcoes), pdf = await PDFDocument.load(bytes);
        const d = diagnosticarAssinaturas(pdf, { tamanhoArquivo: bytes.length });
        assert.equal(d.presentes, opcoes.vazia ? 0 : opcoes.paginas?.length || 1);
        assert.equal(d.vazios, opcoes.vazia ? 1 : 0);
        assert.equal(d.assinaturas[0].invisivel, !!opcoes.invisivel);
        assert.deepEqual(d.assinaturas.flatMap(s => s.paginasVisuais).sort(), opcoes.invisivel ? [] : opcoes.paginas || [3]);
        assert.equal(d.validade, "Validade criptográfica não verificada");
    }
});
test("campo sem V não é assinatura; V parcial/ByteRange inválido e indícios isolados são inconclusivos", async () => {
    for (const tipo of ["sem-v", "contents-zero", "byte-range", "sigflags", "orfao"]) {
        const pdf = await PDFDocument.create(); pdf.addPage([200, 300]);
        if (tipo === "sigflags") pdf.catalog.set(nome("AcroForm"), pdf.context.obj({ SigFlags: 3 }));
        else {
            const { valor } = adicionarAssinatura(pdf, { vazia: tipo === "sem-v" });
            if (tipo === "contents-zero") valor.set(nome("Contents"), PDFHexString.of("000000"));
            if (tipo === "byte-range") valor.set(nome("ByteRange"), pdf.context.obj([0, 20, 10, 20]));
            if (tipo === "orfao") { pdf.catalog.delete(nome("AcroForm")); pdf.getPage(0).node.delete(nome("Annots")); }
        }
        const d = diagnosticarAssinaturas(pdf);
        assert.equal(d.presentes, 0);
        assert.equal(d.estado, tipo === "sem-v" ? "campos_vazios" : "inconclusiva");
    }
});
test("widget órfão com V é detectado; Perms deduplica o mesmo dicionário; assinatura desenhada não é digital", async () => {
    const pdf = await PDFDocument.create(); pdf.addPage([200, 300]);
    adicionarAssinatura(pdf, { orfa: true, permissao: true });
    assert.equal(diagnosticarAssinaturas(pdf).presentes, 1);
    assert.throws(() => verificarSemAssinaturas(pdf));
    const imagem = await PDFDocument.create(); imagem.addPage([200, 300]).drawText("Assinatura / rubrica desenhada");
    assert.equal(diagnosticarAssinaturas(imagem).estado, "nao_identificada");
    verificarSemAssinaturas(imagem);
    imagem.catalog.set(nome("AcroForm"), imagem.context.obj({ Fields: [], SigFlags: 0 }));
    assert.equal(diagnosticarAssinaturas(imagem).estado, "nao_identificada", "SigFlags zero não deve criar um falso alerta");
});
test("ausência de AP e de flag de impressão não são confundidas com assinatura invisível", async () => {
    const d = diagnosticarAssinaturas(await PDFDocument.load(await criarPDFAssinado({ semAparencia: true, imprimivel: false })));
    assert.equal(d.presentes, 1); assert.equal(d.assinaturas[0].invisivel, false);
    assert.equal(d.assinaturas[0].widgets[0].aparencia, false); assert.equal(d.assinaturas[0].widgets[0].imprimivel, false);
});
