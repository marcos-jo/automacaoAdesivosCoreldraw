import { PDFDocument, PDFName, PDFDict, PDFArray, PDFNumber, PDFObjectCopier, degrees,
    pushGraphicsState, popGraphicsState, concatTransformationMatrix, drawObject } from "/vendor/pdf-lib/pdf-lib.js";
import { diagnosticarAssinaturas, verificarSemAssinaturas, dicionariosPDF } from "./pdf-signatures.js";
const nome = PDFName.of;
export function cancelarImpressao(signal) { if (signal?.aborted) throw new DOMException("Preparação cancelada", "AbortError"); }

export function planejarCopiaImpressao(pdf, config) {
    if (![300, 600].includes(config.dpi) || !["necessarias", "inteiro"].includes(config.modo)) throw new Error("Escolha 300/600 DPI e um modo de preparação válido.");
    const paginas = pdf.getPages(), diagnostico = diagnosticarAssinaturas(pdf), necessarias = new Set(), avisos = [];
    const form = pdf.catalog.lookupMaybe(nome("AcroForm"), PDFDict), nomes = pdf.catalog.lookupMaybe(nome("Names"), PDFDict);
    if (form?.has(nome("XFA")) || form?.lookup(nome("NeedAppearances"))?.toString() === "true" || pdf.catalog.lookup(nome("NeedsRendering"))?.toString() === "true" || form?.lookupMaybe(nome("CO"), PDFArray)?.size()) throw new Error("XFA/geração dinâmica ou cálculo de aparências não pode ser comprovado por este fluxo. Prepare em outro leitor confiável.");
    const acao = pdf.catalog.lookup(nome("OpenAction"));
    if (dicionariosPDF(pdf).some(d => d.has(nome("AA"))) || nomes?.has(nome("JavaScript")) || acao instanceof PDFDict && (acao.has(nome("Next")) || !["/GoTo", "/URI", "/GoToR"].includes(acao.lookupMaybe(nome("S"), PDFName)?.asString()))) throw new Error("Scripts/ações adicionais do documento ou campos podem alterar a impressão. Preparação bloqueada.");
    if (config.modo === "necessarias" && pdf.catalog.has(nome("OCProperties"))) throw new Error("Camadas/OCProperties dependem do catálogo. Use Rasterizar documento inteiro para preservar sua renderização de impressão.");
    for (const assinatura of diagnostico.assinaturas) for (const widget of assinatura.widgets) {
        if (!widget.retanguloValido) throw new Error("Widget de assinatura com retângulo inválido: não é possível determinar a aparência.");
        if (widget.visivel) {
            if (!widget.aparencia || !widget.imprimivel) throw new Error(`Página ${widget.pagina}: assinatura visual sem aparência AP utilizável ou sem flag de impressão. O PDF.js não permite comprovar sua impressão; não será apresentada como corrigida. Confira em outro leitor.`);
            necessarias.add(widget.pagina);
        }
    }
    const anotacoes = new Map();
    paginas.forEach((p, indice) => {
        if (p.node.has(nome("AA"))) throw new Error(`Página ${indice + 1}: ações adicionais de impressão não suportadas.`);
        const imprimiveis = [];
        for (const ref of p.node.Annots()?.asArray() || []) {
            const d = pdf.context.lookup(ref, PDFDict), flags = d.lookupMaybe(nome("F"), PDFNumber)?.asNumber() || 0;
            if (d.has(nome("AA"))) throw new Error(`Página ${indice + 1}: anotação com ações adicionais não tratadas.`);
            const rect = d.lookupMaybe(nome("Rect"), PDFArray)?.asArray().map(v => pdf.context.lookup(v)?.asNumber?.());
            if (flags & 4 && (!rect || rect.length !== 4 || !rect.every(Number.isFinite))) throw new Error(`Página ${indice + 1}: retângulo de anotação imprimível não verificável.`);
            const area = rect?.length === 4 && rect[0] !== rect[2] && rect[1] !== rect[3];
            if (flags & 4 && !(flags & 3) && area) { imprimiveis.push(ref.toString()); necessarias.add(indice + 1); }
            else avisos.push(`Página ${indice + 1}: anotações não imprimíveis, invisíveis ou sem área não serão transportadas; compare com o modo de impressão do original.`);
        }
        anotacoes.set(indice + 1, imprimiveis);
    });
    const manual = config.paginas || [];
    if (!Array.isArray(manual) || manual.some(n => !Number.isSafeInteger(n) || n < 1 || n > paginas.length)) throw new Error("Páginas manuais inválidas.");
    const rasterizadas = config.modo === "inteiro" ? paginas.map((_, i) => i + 1) : [...new Set([...necessarias, ...manual])].sort((a, b) => a - b);
    if (!rasterizadas.length) throw new Error("Não há páginas visuais determinadas automaticamente. Informe páginas manualmente ou escolha Rasterizar documento inteiro; nenhuma página foi inventada.");
    return { dpi: config.dpi, modo: config.modo, rasterizadas, necessarias: [...necessarias].sort((a, b) => a - b), diagnostico, anotacoes,
        avisos: [...new Set(avisos)], vetoriais: paginas.map((_, i) => i + 1).filter(n => !rasterizadas.includes(n)) };
}

function copiarCaixas(origem, destino) {
    const m = origem.getMediaBox(), c = origem.getCropBox();
    const unidade = origem.node.lookupMaybe(nome("UserUnit"), PDFNumber)?.asNumber() ?? 1;
    const r = ((origem.getRotation().angle % 360) + 360) % 360, w = m.width * unidade, h = m.height * unidade;
    if (![0, 90, 180, 270].includes(r) || !Number.isFinite(unidade) || unidade <= 0) throw new Error("Rotação/UserUnit não suportado para cópia de impressão.");
    const ponto = (x, y) => {
        const a = (x - m.x) * unidade, b = (y - m.y) * unidade;
        return r === 0 ? [a, b] : r === 90 ? [b, w - a] : r === 180 ? [w - a, h - b] : [h - b, a];
    };
    const transformar = box => {
        const a = ponto(box.x, box.y), b = ponto(box.x + box.width, box.y + box.height);
        return [Math.min(a[0], b[0]), Math.min(a[1], b[1]), Math.max(a[0], b[0]), Math.max(a[1], b[1])];
    };
    destino.setMediaBox(0, 0, r % 180 ? h : w, r % 180 ? w : h);
    const crop = transformar(c); destino.setCropBox(crop[0], crop[1], crop[2] - crop[0], crop[3] - crop[1]);
    // A imagem já inclui a orientação REAL de impressão e efeitos NoRotate dos
    // widgets. Não renderizar com Rotate=0 na fonte e girar a assinatura depois.
    destino.setRotation(degrees(0));
    for (const chave of ["TrimBox", "BleedBox", "ArtBox"]) {
        const box = origem.node.lookupMaybe(nome(chave), PDFArray);
        if (box) {
            const [x, y, x2, y2] = box.asArray().map(v => box.context.lookup(v).asNumber());
            destino.node.set(nome(chave), destino.doc.context.obj(transformar({ x, y, width: x2 - x, height: y2 - y })));
        }
    }
}

// Reconstrução em um CONTEXTO NOVO; não clona o catálogo/AcroForm/Perms.
export async function reconstruirCopiaImpressao(origem, config, { signal, rasterizar, progresso = () => {} } = {}) {
    cancelarImpressao(signal);
    const plano = planejarCopiaImpressao(origem, config), novo = await PDFDocument.create(), imagens = new Map(), paginas = origem.getPages();
    // Remover referências interativas só na cópia transacional. Suas aparências
    // imprimíveis estão no plano de rasterização; PDF.js lê bytes imutáveis separados.
    for (const p of paginas) for (const chave of ["Annots", "AA", "StructParents", "Tabs", "Metadata", "PieceInfo", "PresSteps", "B", "AF"]) p.node.delete(nome(chave));
    if (config.modo === "necessarias" && origem.catalog.has(nome("OutputIntents"))) novo.catalog.set(nome("OutputIntents"), PDFObjectCopier.for(origem.context, novo.context).copy(origem.catalog.get(nome("OutputIntents"))));
    novo.catalog.set(nome("ViewerPreferences"), novo.context.obj({ PrintScaling: "None" }));
    const copias = await novo.copyPages(origem, plano.vetoriais.map(n => n - 1));
    const vetoriais = new Map(plano.vetoriais.map((n, i) => [n, copias[i]]));
    let acumulados = 0;
    for (let i = 0; i < paginas.length; i++) {
        cancelarImpressao(signal);
        const numero = i + 1;
        progresso(`Preparando página ${numero} de ${paginas.length}...`);
        if (plano.rasterizadas.includes(numero)) {
            const imagem = await rasterizar(numero, plano);
            cancelarImpressao(signal);
            if (!(imagem.dados instanceof Uint8Array) || !imagem.dados.length || ![imagem.largura, imagem.altura].every(n => Number.isSafeInteger(n) && n > 0)) throw new Error("Rasterização não retornou uma imagem válida.");
            acumulados += imagem.dados.length;
            if (acumulados > 256 * 1024 * 1024) throw new Error("As imagens comprimidas excedem o limite de 256 MiB desta preparação. Divida o trabalho.");
            const pagina = novo.addPage(); copiarCaixas(paginas[i], pagina);
            const ref = novo.context.register(novo.context.stream(imagem.dados, { Type: "XObject", Subtype: "Image", Width: imagem.largura, Height: imagem.altura, ColorSpace: "DeviceRGB", BitsPerComponent: 8, Filter: "FlateDecode", Interpolate: false }));
            pagina.node.setXObject(nome("Impressao"), ref);
            const m = pagina.getMediaBox(), c = pagina.getCropBox();
            const x = Math.max(m.x, c.x), yTopo = Math.min(m.y + m.height, c.y + c.height), escala = imagem.dpi / 72;
            pagina.pushOperators(pushGraphicsState(), concatTransformationMatrix(imagem.largura / escala, 0, 0, imagem.altura / escala, x, yTopo - imagem.altura / escala), drawObject(nome("Impressao")), popGraphicsState());
            imagens.set(numero, imagem);
        } else {
            novo.addPage(vetoriais.get(numero));
        }
        await new Promise(resolve => setTimeout(resolve, 0));
    }
    cancelarImpressao(signal);
    verificarSemAssinaturas(novo);
    return { pdf: novo, plano, imagens };
}
