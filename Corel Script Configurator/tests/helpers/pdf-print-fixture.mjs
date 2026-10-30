import { readFile } from "node:fs/promises";
import { PDFDocument, PDFName, PDFString, PDFHexString, rgb, degrees } from "../../public/vendor/pdf-lib/pdf-lib.js";

export const textoImpressao = "Gráfica: ação, coração • Ω ± € → símbolos";
const nome = PDFName.of;

// Fonte do sistema somente nos PDFs temporários de teste; nenhum binário é distribuído.
async function fonteSistema() {
    for (const caminho of ["C:/Windows/Fonts/arial.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"]) {
        try { return await readFile(caminho); } catch (erro) { if (erro.code !== "ENOENT") throw erro; }
    }
    throw new Error("Teste de fonte incorporada requer Arial ou DejaVu Sans instalada.");
}

// Fixture TrueType/Identity-H independente de fontkit. CMap formato 4 (BMP).
async function incorporarFonte(pdf, texto) {
    const bytes = await fonteSistema(), tabelas = {};
    for (let i = 0; i < bytes.readUInt16BE(4); i++) {
        const o = 12 + i * 16; tabelas[bytes.toString("ascii", o, o + 4)] = bytes.readUInt32BE(o + 8);
    }
    const cmap = tabelas.cmap;
    let mapa;
    for (let i = 0; i < bytes.readUInt16BE(cmap + 2); i++) {
        const o = cmap + bytes.readUInt32BE(cmap + 4 + i * 8 + 4);
        if (bytes.readUInt16BE(o) === 4) { mapa = o; break; }
    }
    if (!mapa) throw new Error("Fonte de teste sem cmap formato 4.");
    const segmentos = bytes.readUInt16BE(mapa + 6) / 2;
    const fim = mapa + 14, inicio = fim + 2 * segmentos + 2, delta = inicio + 2 * segmentos, range = delta + 2 * segmentos;
    const glifo = codigo => {
        for (let i = 0; i < segmentos; i++) if (codigo >= bytes.readUInt16BE(inicio + 2 * i) && codigo <= bytes.readUInt16BE(fim + 2 * i)) {
            const deslocamento = bytes.readUInt16BE(range + 2 * i), d = bytes.readInt16BE(delta + 2 * i);
            if (!deslocamento) return (codigo + d) & 65535;
            const g = bytes.readUInt16BE(range + 2 * i + deslocamento + 2 * (codigo - bytes.readUInt16BE(inicio + 2 * i)));
            return g ? (g + d) & 65535 : 0;
        }
        return 0;
    };
    const caracteres = [...new Set(texto)], unidade = bytes.readUInt16BE(tabelas.head + 18), medida = v => v * 1000 / unidade;
    const ids = caracteres.map(c => glifo(c.charCodeAt(0)));
    if (ids.some(id => !id)) throw new Error("Fonte de teste sem um dos símbolos exigidos.");
    const cidMapa = Buffer.alloc((caracteres.length + 1) * 2);
    ids.forEach((id, i) => cidMapa.writeUInt16BE(id, (i + 1) * 2));
    const metricas = bytes.readUInt16BE(tabelas.hhea + 34);
    const larguras = ids.map(id => medida(bytes.readUInt16BE(tabelas.hmtx + Math.min(id, metricas - 1) * 4)));
    const contexto = pdf.context;
    const arquivo = contexto.register(contexto.flateStream(bytes, { Length1: bytes.length }));
    const descritor = contexto.register(contexto.obj({ Type: "FontDescriptor", FontName: "PrintFixture", Flags: 32,
        FontBBox: [36, 38, 40, 42].map(o => medida(bytes.readInt16BE(tabelas.head + o))), ItalicAngle: 0,
        Ascent: medida(bytes.readInt16BE(tabelas.hhea + 4)), Descent: medida(bytes.readInt16BE(tabelas.hhea + 6)), CapHeight: 700, StemV: 80, FontFile2: arquivo }));
    const descendente = contexto.register(contexto.obj({ Type: "Font", Subtype: "CIDFontType2", BaseFont: "PrintFixture",
        CIDSystemInfo: { Registry: PDFString.of("Adobe"), Ordering: PDFString.of("Identity"), Supplement: 0 },
        FontDescriptor: descritor, W: [1, larguras], CIDToGIDMap: contexto.register(contexto.flateStream(cidMapa)) }));
    const hex = n => n.toString(16).padStart(4, "0");
    const unicode = `/CIDInit /ProcSet findresource begin 12 dict begin begincmap
/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def
/CMapName /PrintFixture-UCS def /CMapType 2 def
1 begincodespacerange <0000> <ffff> endcodespacerange
${caracteres.length} beginbfchar
${caracteres.map((c, i) => `<${hex(i + 1)}> <${hex(c.charCodeAt(0))}>`).join("\n")}
endbfchar endcmap CMapName currentdict /CMap defineresource pop end end`;
    const fonte = contexto.register(contexto.obj({ Type: "Font", Subtype: "Type0", BaseFont: "PrintFixture", Encoding: "Identity-H",
        DescendantFonts: [descendente], ToUnicode: contexto.register(contexto.flateStream(unicode)) }));
    return { fonte, codificado: [...texto].map(c => hex(caracteres.indexOf(c) + 1)).join("") };
}

export function marcarEstrutura(pdf, pagina = pdf.getPage(0), actualText = false) {
    const c = pdf.context, raiz = c.register(c.obj({ Type: "StructTreeRoot" }));
    const elemento = c.register(c.obj({ Type: "StructElem", S: "P", P: raiz, Pg: pagina.ref, K: 0,
        ...(actualText === true || typeof actualText === "string" ? { ActualText: PDFHexString.fromText(typeof actualText === "string" ? actualText : "Texto substituto essencial") } : {}) }));
    const dict = c.lookup(raiz); dict.set(nome("K"), c.obj([elemento]));
    dict.set(nome("ParentTree"), c.register(c.obj({ Nums: [0, [elemento]] })));
    dict.set(nome("ParentTreeNextKey"), c.obj(1));
    pdf.catalog.set(nome("StructTreeRoot"), raiz); pdf.catalog.set(nome("MarkInfo"), c.obj({ Marked: true }));
    pagina.node.set(nome("StructParents"), c.obj(0)); pagina.node.set(nome("Tabs"), nome("S"));
}

export async function criarFixtureImpressao({ marcada = true, navegacao = false, links = false, rotacao = 0, actualText = false } = {}) {
    const pdf = await PDFDocument.create(), p = pdf.addPage([960, 540]), c = pdf.context;
    const { fonte, codificado } = await incorporarFonte(pdf, textoImpressao);
    p.setRotation(degrees(rotacao));
    for (let i = 1; i <= 5; i++) p.node.setFontDictionary(nome("F" + i), fonte);
    // Texto marcado e texto substituto NO CONTEÚDO permanecem intactos.
    p.node.addContentStream(c.register(c.flateStream(`/P << /MCID 0 >> BDC BT /F1 27 Tf 1 0 0 1 70 420 Tm <${codificado}> Tj ET EMC`)));
    p.node.addContentStream(c.register(c.flateStream(`/Span << /ActualText <feff0054006500780074006f00200061006c0074> >> BDC BT /F2 18 Tf 1 0 0 1 70 360 Tm <${codificado}> Tj ET EMC`)));
    p.drawRectangle({ x: 50, y: 70, width: 260, height: 190, color: rgb(0, .4, .8), opacity: .45, borderWidth: 3, borderColor: rgb(.9, 0, 0) });
    p.drawEllipse({ x: 430, y: 180, xScale: 100, yScale: 65, color: rgb(0, .8, .2), opacity: .6 });
    const png = Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aV1cAAAAASUVORK5CYII=", "base64");
    for (let i = 0; i < 6; i++) { const imagem = await pdf.embedPng(png); p.drawImage(imagem, { x: 600 + i * 35, y: 80, width: 24, height: 45 }); }
    const form = c.register(c.flateStream("0.8 0.1 0.6 rg 0 0 35 35 re f", { Type: "XObject", Subtype: "Form", BBox: [0, 0, 35, 35], Resources: {}, ...(marcada ? { StructParents: 1 } : {}) }));
    p.node.setXObject(nome("Vetor"), form);
    p.node.addContentStream(c.register(c.flateStream("q 1 0 0 1 820 180 cm /Vetor Do Q")));
    pdf.catalog.set(nome("Lang"), PDFString.of("pt-BR"));
    if (marcada) marcarEstrutura(pdf, p, actualText);
    if (navegacao) {
        const destino = c.obj([p.ref, "XYZ", 0, 540, null]);
        const outlines = c.register(c.obj({ Type: "Outlines", Count: 1 }));
        const item = c.register(c.obj({ Title: PDFString.of("Slide 1"), Parent: outlines, Dest: destino }));
        c.lookup(outlines).set(nome("First"), item); c.lookup(outlines).set(nome("Last"), item);
        pdf.catalog.set(nome("Outlines"), outlines); pdf.catalog.set(nome("PageMode"), nome("UseOutlines"));
        pdf.catalog.set(nome("Dests"), c.obj({ Slide: destino }));
        pdf.catalog.set(nome("Names"), c.obj({ Dests: { Names: [PDFString.of("slide"), destino] }, EmbeddedFiles: { Names: [] } }));
        pdf.catalog.set(nome("OpenAction"), c.obj({ S: "GoTo", D: destino }));
    }
    if (links) for (const acao of [null, { S: "GoTo", D: [p.ref, "Fit"] }, { S: "URI", URI: PDFString.of("https://example.com") }]) {
        p.node.addAnnot(c.register(c.obj({ Type: "Annot", Subtype: "Link", P: p.ref, Rect: [50, 300, 300, 330], Border: [0, 0, 2], C: [1, 0, 0],
            ...(acao ? { A: acao } : { Dest: [p.ref, "Fit"] }) })));
    }
    return pdf.save();
}
