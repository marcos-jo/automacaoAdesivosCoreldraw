import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { createHash } from "node:crypto";
import path from "node:path";
import { PDFDocument, PDFName, PDFHexString, PDFStream } from "../../public/vendor/pdf-lib/pdf-lib.js";
import { criarFixtureImpressao, textoImpressao } from "./pdf-print-fixture.mjs";
import { impressaoUrl } from "./pdf-resize-modules.mjs";
const { diagnosticarImpressao, validarActualTextImpressao, prepararParaImpressao } = await import(impressaoUrl);
const nome = PDFName.of, base64 = bytes => Buffer.from(bytes).toString("base64");
const hash = bytes => createHash("sha256").update(bytes).digest("hex");
const streams = pdf => new Map(pdf.context.enumerateIndirectObjects().filter(([, o]) => o instanceof PDFStream).map(([r, o]) => [r.toString(), hash(o.getContents())]));

// Opcional: PDF_ACTUALTEXT_REAL aponta para um arquivo LOCAL, fora do repositório.
// Nenhum conteúdo de cliente é gravado no código, fixtures ou logs.
export async function testarActualTextBrowser({ avaliar, clicar, pronto, esperarTela, configurarResize, resizePronto, confirmarImpressao, baixar, conferirPixelsExportados, imagem, temporario }) {
    const misto = await PDFDocument.load(await criarFixtureImpressao({ actualText: textoImpressao }));
    const p = misto.getPage(0), c = misto.context, raiz = misto.catalog.lookup(nome("StructTreeRoot"));
    p.node.addContentStream(c.register(c.flateStream("/Span << /MCID 42 >> BDC 0 0 30 40 re f EMC")));
    raiz.lookup(nome("K")).push(c.register(c.obj({ Type: "StructElem", S: "Span", P: misto.catalog.get(nome("StructTreeRoot")), Pg: p.ref, K: [42], ActualText: PDFHexString.fromText(textoImpressao) })));
    const vetorial = await PDFDocument.load(await misto.save());
    vetorial.catalog.lookup(nome("StructTreeRoot")).lookup(nome("K")).remove(0);
    const casos = [
        { id: "redundante", bytes: await criarFixtureImpressao({ actualText: textoImpressao }), compativel: true },
        { id: "essencial", bytes: await criarFixtureImpressao({ actualText: "Símbolos que só existem na estrutura" }), compativel: false },
        { id: "vetorial", bytes: await vetorial.save(), compativel: false },
        { id: "misto", bytes: await misto.save(), compativel: false }
    ];
    const caminhoReal = process.env.PDF_ACTUALTEXT_REAL;
    if (caminhoReal) casos.push({ id: "real", bytes: new Uint8Array(await readFile(caminhoReal)), compativel: true });
    else console.log("ActualText real: não executado (informe PDF_ACTUALTEXT_REAL).");

    const validarNoBrowser = dados => avaliar(`(async()=>{
        const {validarCompatibilidadeImpressao}=await import('/js/tools/pdf-print-validation.js');
        const {PDFPreview}=await import('/js/tools/pdf-preview.js');
        const dono=new PDFPreview(document.createElement('div'));
        try{return await validarCompatibilidadeImpressao({
            antes:Uint8Array.from(atob(${JSON.stringify(base64(dados.antes))}),c=>c.charCodeAt(0)),
            depois:Uint8Array.from(atob(${JSON.stringify(base64(dados.depois))}),c=>c.charCodeAt(0)),
            substitutos:${JSON.stringify(dados.substitutos)}
        },{abrir:(b,o)=>dono.preparar(b,o),fechar:c=>dono.descartar(c)});}
        finally{await dono.destruir();}
    })()`);

    for (const caso of casos) {
        const pdf = await PDFDocument.load(caso.bytes, { updateMetadata: false });
        const diagnostico = diagnosticarImpressao(pdf, [1]);
        assert.ok(diagnostico.actualText.length > 0);
        if (!caso.compativel) {
            await assert.rejects(validarActualTextImpressao(pdf, diagnostico, validarNoBrowser), /não é redundante/);
            assert.equal(pdf.catalog.has(nome("StructTreeRoot")), true);
        } else {
            await validarActualTextImpressao(pdf, diagnostico, validarNoBrowser);
            assert.equal(diagnostico.validacaoActualText.compativel, true);
            prepararParaImpressao(pdf, diagnostico);
            const comparacao = await avaliar(`(async()=>{const r=await compararImpressao(${JSON.stringify(base64(caso.bytes))},${JSON.stringify(base64(await pdf.save()))});return {dimensoes:r.dimensoes,diferentes:r.diferentes,total:r.total,textosIguais:r.textoAntes===r.textoDepois,caracteres:r.textoAntes.length};})()`);
            assert.equal(comparacao.dimensoes, true); assert.equal(comparacao.diferentes, 0); assert.equal(comparacao.textosIguais, true);
            console.log("ActualText " + caso.id + ":", JSON.stringify(diagnostico.validacaoActualText), "comparação adicional com bytes de origem:", JSON.stringify(comparacao));
        }

        // Fluxo real de interface: não chama o engine diretamente.
        const arquivo = "actualtext-" + caso.id;
        await avaliar(`selecionarPDFTeste('pdf-arquivo',[new File([Uint8Array.from(atob(${JSON.stringify(base64(caso.bytes))}),c=>c.charCodeAt(0))],${JSON.stringify(arquivo + ".pdf")},{type:'application/pdf'})])`);
        await pronto(); await clicar("#pdf-ferramenta-redimensionar");
        await configurarResize("ajustar", "todas", "210", "297");
        if (caso.id === "real") await imagem("09-actualtext-real-original");
        await clicar("#pdf-redim-prever");
        if (!caso.compativel) {
            await esperarTela(`document.querySelector('#pdf-redim-resultado').textContent.includes('não é redundante')`, "ActualText essencial bloqueado");
            assert.equal(await avaliar(`document.querySelector('#pdf-redim-impressao').hidden`), true);
            assert.equal(await avaliar(`document.querySelector('#pdf-alteracoes-pendentes').hidden`), true);
            await pronto(); continue;
        }
        await esperarTela(`!document.querySelector('#pdf-redim-impressao').hidden`, "ActualText validado, aguardando consentimento");
        assert.match(await avaliar(`document.querySelector('#pdf-redim-resultado').textContent`), /redundantes verificadas/);
        await confirmarImpressao(false, false); await pronto();
        assert.equal(await avaliar(`document.querySelector('#pdf-alteracoes-pendentes').hidden`), true);
        await confirmarImpressao(true, false); await resizePronto();
        await clicar("#pdf-redim-cancelar"); await pronto();
        assert.equal(await avaliar(`document.querySelector('#pdf-alteracoes-pendentes').hidden`), true);
        await clicar("#pdf-redim-prever"); await esperarTela(`!document.querySelector('#pdf-redim-impressao').hidden`, "ActualText exige nova confirmação");
        await confirmarImpressao(true, false); await resizePronto();
        await avaliar(`(()=>{const c=document.querySelector('#pdf-visualizacao canvas');window.pixelsActualText={w:c.width,h:c.height,d:c.getContext('2d').getImageData(0,0,c.width,c.height).data};})()`);
        if (caso.id === "real") await imagem("10-actualtext-real-previa");
        await clicar("#pdf-redim-aplicar"); await pronto();
        assert.equal(await avaliar(`(()=>{const c=document.querySelector('#pdf-visualizacao canvas'),p=pixelsActualText;return c.width===p.w&&c.height===p.h&&c.getContext('2d').getImageData(0,0,c.width,c.height).data.every((v,i)=>v===p.d[i]);})()`), true);
        assert.equal(await avaliar(`document.querySelector('#pdf-alteracoes-pendentes').hidden`), false);
        await clicar("#pdf-salvar-copia");
        const exportado = await baixar(arquivo + "_editado.pdf", 1);
        assert.ok(Math.abs(exportado.getPage(0).getWidth() - 210 * 72 / 25.4) < 1e-6);
        assert.ok(Math.abs(exportado.getPage(0).getHeight() - 297 * 72 / 25.4) < 1e-6);
        assert.equal(exportado.catalog.has(nome("StructTreeRoot")), false);
        const originais = streams(await PDFDocument.load(caso.bytes)), finais = streams(exportado);
        for (const [ref, resumo] of originais) assert.equal(finais.get(ref), resumo, "Stream original foi alterado: " + ref);
        await conferirPixelsExportados(arquivo + "_editado.pdf", 1);
        const bytesExportados = await readFile(path.join(temporario, arquivo + "_editado.pdf"));
        const textoFinal = await avaliar(`(async()=>{
            const {getDocument}=await import('/vendor/pdfjs/pdf.js');
            const texto=async(b64)=>{const t=getDocument({data:Uint8Array.from(atob(b64),c=>c.charCodeAt(0))});try{const p=await(await t.promise).getPage(1),c=await p.getTextContent({disableNormalization:true});return c.items.map(i=>i.str+(i.hasEOL?'\\n':'')).join('');}finally{await t.destroy();}};
            const a=await texto(${JSON.stringify(base64(caso.bytes))}),b=await texto(${JSON.stringify(base64(bytesExportados))});
            const normalizar=s=>s.normalize('NFC').replace(/\\s+/gu,' ').trim();
            return {iguais:normalizar(a)===normalizar(b),caracteresAntes:a.length,caracteresDepois:b.length};
        })()`);
        assert.equal(textoFinal.iguais, true, "Texto extraível depois do redimensionamento/exportação");
        console.log("Texto após A4 " + caso.id + ":", JSON.stringify(textoFinal));
        await avaliar(`selecionarPDFTeste('pdf-arquivo',[new File([Uint8Array.from(atob(${JSON.stringify(base64(bytesExportados))}),c=>c.charCodeAt(0))],'reaberto.pdf',{type:'application/pdf'})])`);
        await pronto();
        assert.equal(await avaliar(`document.querySelector('#pdf-alteracoes-pendentes').hidden`), true);
        await conferirPixelsExportados(arquivo + "_editado.pdf", 1);
        if (caso.id === "real") {
            await imagem("11-actualtext-real-reaberto");
            assert.equal(hash(await readFile(caminhoReal)), hash(caso.bytes), "O arquivo de cliente original deve permanecer intacto");
            console.log("ActualText real: fluxo A4 importar/pré-visualizar/recusar/aceitar/cancelar/aplicar/exportar/reabrir passou; streams originais intactos; SHA-256 de origem preservado.");
        }
    }
}
