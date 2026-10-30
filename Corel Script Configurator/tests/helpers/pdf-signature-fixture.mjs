import { PDFDocument, PDFName, PDFString, PDFHexString, rgb, degrees } from "../../public/vendor/pdf-lib/pdf-lib.js";
const nome = PDFName.of;
export function adicionarAssinatura(pdf, { pagina = 1, vazia = false, invisivel = false, semAparencia = false, imprimivel = true, orfa = false, permissao = false } = {}) {
    const c = pdf.context, p = pdf.getPage(pagina - 1);
    const valor = vazia ? undefined : c.register(c.obj({ Type: "Sig", Filter: "Adobe.PPKLite", SubFilter: "adbe.pkcs7.detached", ByteRange: [0, 10, 20, 10], Contents: PDFHexString.of("0102030405060708") }));
    const campo = c.obj({ FT: "Sig", T: PDFString.of("Campo " + pagina), ...(valor ? { V: valor } : {}) });
    const refCampo = c.register(campo);
    const fonte = c.register(c.obj({ Type: "Font", Subtype: "Type1", BaseFont: "Helvetica" }));
    const ap = c.register(c.flateStream("q 1 1 1 rg 0 0 180 50 re f 0.1 0.3 0.7 RG 2 w 2 2 176 46 re S 0 0 0 rg BT /F1 14 Tf 8 20 Td (AP ASSINATURA VISIVEL) Tj ET Q", { Type: "XObject", Subtype: "Form", BBox: [0, 0, 180, 50], Resources: { Font: { F1: fonte } } }));
    const widget = c.register(c.obj({ Type: "Annot", Subtype: "Widget", Parent: refCampo, P: p.ref, F: imprimivel ? 4 : 0,
        Rect: invisivel ? [0, 0, 0, 0] : [20, 30, 200, 80], ...(!invisivel && !semAparencia ? { AP: { N: ap } } : {}) }));
    campo.set(nome("Kids"), c.obj([widget])); p.node.addAnnot(widget);
    if (!orfa) {
        let form = pdf.catalog.lookup(nome("AcroForm"));
        if (!form) { form = c.obj({ Fields: [], SigFlags: 3 }); pdf.catalog.set(nome("AcroForm"), form); }
        form.lookup(nome("Fields")).push(refCampo);
    }
    if (permissao && valor) pdf.catalog.set(nome("Perms"), c.obj({ DocMDP: valor }));
    return { campo, widget: c.lookup(widget), valor: valor && c.lookup(valor) };
}

export async function criarPDFAssinado({ total = 3, paginas = [total], vazia = false, invisivel = false, semAparencia = false, imprimivel = true, tamanhos, rotacao = 0, anotacao = false } = {}) {
    const pdf = await PDFDocument.create();
    for (let i = 0; i < total; i++) {
        const p = pdf.addPage(tamanhos?.[i] || [300, 400]);
        p.setRotation(degrees(rotacao));
        p.drawText("Documento sintetico pagina " + (i + 1), { x: 10, y: 200, size: 12 });
        p.drawRectangle({ x: 5, y: 5, width: 15, height: 30, color: rgb(.1, .8, .3), opacity: .6 });
    }
    for (const pagina of paginas) adicionarAssinatura(pdf, { pagina, vazia, invisivel, semAparencia, imprimivel });
    if (anotacao) {
        const c = pdf.context, ap = c.register(c.flateStream("1 0 0 rg 0 0 40 20 re f", { Type: "XObject", Subtype: "Form", BBox: [0, 0, 40, 20], Resources: {} }));
        pdf.getPage(0).node.addAnnot(c.register(c.obj({ Type: "Annot", Subtype: "Stamp", Rect: [10, 100, 50, 120], AP: { N: ap }, F: 4 })));
    }
    return pdf.save({ updateFieldAppearances: false });
}
