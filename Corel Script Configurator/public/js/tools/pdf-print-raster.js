// Renderização de impressão, RGB opaco sobre branco, sem JPEG nem PNG gigante
// intermediário: compacta faixas RGB sequencialmente com CompressionStream.
// Multiplicações com UserUnit podem produzir 2375.0000000000005. Não criar
// uma coluna extra por esse erro de ponto flutuante (< 1e-7 pixel).
const pixels = valor => Math.ceil(valor - 1e-7);
export function estimarRaster(larguraPontos, alturaPontos, dpi) {
    if (![300, 600].includes(dpi)) throw new Error("Resolução inválida.");
    const largura = pixels(larguraPontos * dpi / 72), altura = pixels(alturaPontos * dpi / 72), total = largura * altura;
    if (![largura, altura, total].every(Number.isSafeInteger) || largura <= 0 || altura <= 0 || largura > 12000 || altura > 12000 || total > 36000000) throw new Error(`Página excede o limite de 36 milhões de pixels/12000 por lado em ${dpi} DPI. Escolha 300 DPI ou um formato menor; a resolução não será reduzida automaticamente.`);
    return { largura, altura, pixels: total, bytesCanvas: total * 4, picoEstimado: total * 16 };
}
function interromper(signal) { if (signal?.aborted) throw new DOMException("Preparação cancelada", "AbortError"); }
async function aguardar(promise, signal, cancelar = () => {}) {
    let abortar, timer;
    const limite = new Promise((_, reject) => {
        abortar = () => reject(new DOMException("Preparação cancelada", "AbortError"));
        timer = setTimeout(() => reject(new Error("Tempo limite na renderização/preparação para impressão.")), 60000);
    });
    signal?.addEventListener("abort", abortar, { once: true });
    try { interromper(signal); return await Promise.race([promise, limite]); }
    catch (erro) { cancelar(); throw erro; }
    finally { clearTimeout(timer); signal?.removeEventListener("abort", abortar); }
}
async function renderizar(pagina, viewport, signal, pdfjs, annotationMode = pdfjs.AnnotationMode.ENABLE) {
    const canvas = document.createElement("canvas"); canvas.width = pixels(viewport.width); canvas.height = pixels(viewport.height);
    const ctx = canvas.getContext("2d", { alpha: false, willReadFrequently: true });
    if (!ctx) { canvas.width = canvas.height = 0; throw new Error("Canvas de impressão indisponível."); }
    try {
        ctx.fillStyle = "white"; ctx.fillRect(0, 0, canvas.width, canvas.height);
        const t = pagina.render({ canvas, canvasContext: ctx, viewport, intent: "print", annotationMode, background: "rgb(255,255,255)" });
        await aguardar(t.promise, signal, () => t.cancel()); return canvas;
    } catch (erro) { canvas.width = canvas.height = 0; throw erro; }
}

async function conferirAparencias(pagina, numero, plano, signal, pdfjs) {
    const esperados = plano.anotacoes.get(numero);
    if (!esperados?.length) return;
    const anotacoes = await aguardar(pagina.getAnnotations({ intent: "print" }), signal);
    const lista = await aguardar(pagina.getOperatorList({ intent: "print", annotationMode: pdfjs.AnnotationMode.ENABLE }), signal);
    const pinturas = new Set(["stroke", "closeStroke", "fill", "eoFill", "fillStroke", "eoFillStroke", "closeFillStroke", "closeEOFillStroke", "showText", "showSpacedText", "nextLineShowText", "nextLineSetSpacingShowText", "shadingFill", "rawFillPath", ...Object.keys(pdfjs.OPS).filter(k => k.startsWith("paintImage") || k.startsWith("paintInlineImage") || k === "paintSolidColorImageMask")].map(k => pdfjs.OPS[k]).filter(Number.isInteger));
    for (const ref of esperados) {
        const [num, ger] = ref.split(" "), id = `${num}R${ger !== "0" ? ger : ""}`;
        const anotacao = anotacoes.find(a => a.id === id);
        const inicio = lista.fnArray.findIndex((fn, i) => fn === pdfjs.OPS.beginAnnotation && lista.argsArray[i][0] === id);
        const fim = inicio < 0 ? -1 : lista.fnArray.indexOf(pdfjs.OPS.endAnnotation, inicio + 1);
        const temPintura = inicio >= 0 && fim > inicio && lista.fnArray.some((fn, i) => i > inicio && i < fim && pinturas.has(fn === pdfjs.OPS.constructPath ? lista.argsArray[i][0] : fn));
        if (!anotacao || !anotacao.hasAppearance || inicio < 0 || fim <= inicio + 1 || !temPintura) throw new Error(`Página ${numero}: aparência imprimível ${ref} não foi disponibilizada pelo PDF.js. Não é possível afirmar que a assinatura/anotação foi preservada.`);
    }
    const widgets = plano.diagnostico.assinaturas.flatMap(s => s.widgets).filter(w => w.pagina === numero && w.visivel);
    if (widgets.length) {
        let com, sem;
        try {
            const viewport = pagina.getViewport({ scale: 1 });
            com = await renderizar(pagina, viewport, signal, pdfjs);
            sem = await renderizar(pagina, viewport, signal, pdfjs, pdfjs.AnnotationMode.DISABLE);
            for (const w of widgets) {
                const v = [...viewport.convertToViewportPoint(w.rect[0], w.rect[1]), ...viewport.convertToViewportPoint(w.rect[2], w.rect[3])];
                const x = Math.max(0, Math.floor(Math.min(v[0], v[2]))), y = Math.max(0, Math.floor(Math.min(v[1], v[3])));
                const largura = Math.min(com.width - x, Math.ceil(Math.max(v[0], v[2])) - x), altura = Math.min(com.height - y, Math.ceil(Math.max(v[1], v[3])) - y);
                if (largura <= 0 || altura <= 0) throw new Error("A região da assinatura ficou fora da página renderizada.");
                const a = com.getContext("2d").getImageData(x, y, largura, altura).data, b = sem.getContext("2d").getImageData(x, y, largura, altura).data;
                if (!a.some((valor, i) => valor !== b[i])) throw new Error(`Página ${numero}: o PDF.js não produziu pixels de anotação na região da assinatura. Confira em outro leitor; a página não será apresentada como corrigida.`);
            }
        } finally { if (com) com.width = com.height = 0; if (sem) sem.width = sem.height = 0; }
    }
}

export async function rasterizarPaginaImpressao(documento, numero, plano, { signal, progresso = () => {} } = {}) {
    const pdfjs = await import("/vendor/pdfjs/pdf.js");
    if (typeof CompressionStream === "undefined") throw new Error("O navegador não oferece compressão local para este preparo.");
    const pagina = await aguardar(documento.getPage(numero), signal), base = pagina.getViewport({ scale: 1 });
    const limites = estimarRaster(base.width, base.height, plano.dpi);
    await conferirAparencias(pagina, numero, plano, signal, pdfjs);
    const viewport = pagina.getViewport({ scale: plano.dpi / 72 });
    let canvas, writer, leitura, finalizada = false;
    try {
        progresso(`Renderizando página ${numero} em ${plano.dpi} DPI (${limites.largura} × ${limites.altura})...`);
        canvas = await renderizar(pagina, viewport, signal, pdfjs);
        const compressor = new CompressionStream("deflate");
        leitura = new Response(compressor.readable).arrayBuffer(); leitura.catch(() => {}); writer = compressor.writable.getWriter();
        for (let y = 0; y < canvas.height; y += 32) {
            interromper(signal);
            const h = Math.min(32, canvas.height - y), rgba = canvas.getContext("2d").getImageData(0, y, canvas.width, h).data;
            const rgb = new Uint8Array(canvas.width * h * 3);
            for (let a = 0, b = 0; a < rgba.length; a += 4) { rgb[b++] = rgba[a]; rgb[b++] = rgba[a + 1]; rgb[b++] = rgba[a + 2]; }
            await aguardar(writer.write(rgb), signal, () => { writer.abort().catch(() => {}); });
            if (y % 256 === 0) await new Promise(resolve => setTimeout(resolve, 0));
        }
        await aguardar(writer.close(), signal); finalizada = true;
        return { dados: new Uint8Array(await aguardar(leitura, signal)), ...limites, area: viewport.viewBox,
            escala: viewport.width / (viewport.viewBox[2] - viewport.viewBox[0]), dpi: plano.dpi };
    } finally {
        if (writer && !finalizada) await writer.abort().catch(() => {});
        if (canvas) canvas.width = canvas.height = 0;
        pagina.cleanup();
    }
}

async function compararImagem(pagina, imagem, signal, pdfjs) {
    const viewport = pagina.getViewport({ scale: imagem.dpi / 72 });
    if (pixels(viewport.width) !== imagem.largura || pixels(viewport.height) !== imagem.altura) throw new Error("A imagem preparada mudou as dimensões físicas.");
    let canvas, leitor;
    try {
        canvas = await renderizar(pagina, viewport, signal, pdfjs);
        leitor = new Blob([imagem.dados]).stream().pipeThrough(new DecompressionStream("deflate")).getReader();
        let bloco = new Uint8Array(), posicao = 0, diferentes = 0, total = 0;
        for (let y = 0; y < canvas.height; y += 32) {
            const h = Math.min(32, canvas.height - y), rgba = canvas.getContext("2d").getImageData(0, y, canvas.width, h).data;
            for (let i = 0; i < rgba.length; i++) {
                if (i % 4 === 3) continue;
                if (posicao === bloco.length) { const l = await aguardar(leitor.read(), signal); if (l.done) throw new Error("Imagem RGB truncada."); bloco = l.value; posicao = 0; }
                if (Math.abs(rgba[i] - bloco[posicao++]) > 2) ++diferentes;
                ++total;
            }
            interromper(signal); await new Promise(resolve => setTimeout(resolve, 0));
        }
        if (posicao !== bloco.length || !(await aguardar(leitor.read(), signal)).done) throw new Error("Dados RGB excedentes na imagem.");
        if (diferentes / total >= .001) throw new Error("A imagem reaberta difere da renderização de impressão; preparo não comprovado.");
        return { diferentes, componentes: total };
    } finally { await leitor?.cancel().catch(() => {}); if (canvas) canvas.width = canvas.height = 0; pagina.cleanup(); }
}

export async function conferirCopiaImpressao(fonte, candidato, preparacao, { signal, progresso = () => {} } = {}) {
    const pdfjs = await import("/vendor/pdfjs/pdf.js"), amostras = [];
    if (fonte.numPages !== candidato.numPages) throw new Error("A cópia perdeu páginas.");
    for (let n = 1; n <= fonte.numPages; n++) {
        interromper(signal); progresso(`Conferindo página preparada ${n} de ${fonte.numPages}...`);
        const a = await aguardar(fonte.getPage(n), signal), b = await aguardar(candidato.getPage(n), signal);
        const va = a.getViewport({ scale: 1 }), vb = b.getViewport({ scale: 1 });
        if (Math.abs(va.width - vb.width) > 1e-6 || Math.abs(va.height - vb.height) > 1e-6) throw new Error(`Página ${n}: dimensões físicas não preservadas.`);
        const imagem = preparacao.imagens.get(n);
        if (imagem) amostras.push({ pagina: n, rasterizada: true, ...await compararImagem(b, imagem, signal, pdfjs) });
        else {
            if (va.width * va.height > 8000000) throw new Error(`Página vetorial ${n}: excede o limite de comparação a 72 DPI. Reveja a estratégia.`);
            let ca, cb;
            try {
                ca = await renderizar(a, va, signal, pdfjs); cb = await renderizar(b, vb, signal, pdfjs);
                const x = ca.getContext("2d").getImageData(0, 0, ca.width, ca.height).data, y = cb.getContext("2d").getImageData(0, 0, cb.width, cb.height).data;
                if (x.length !== y.length || x.some((v, i) => v !== y[i])) throw new Error(`Página vetorial ${n}: aparência alterada na reconstrução; tente Rasterizar documento inteiro.`);
                amostras.push({ pagina: n, rasterizada: false, diferentes: 0, componentes: x.length });
            } finally { if (ca) ca.width = ca.height = 0; if (cb) cb.width = cb.height = 0; a.cleanup(); b.cleanup(); }
        }
    }
    return { compativel: true, amostras, motor: "PDF.js; conferência pelo operador ainda obrigatória" };
}

// A prévia da proposta usa o MESMO intent e modo de anotações do preparo.
export function documentoDeImpressao(fonte) {
    return { fonte, numPages: fonte.numPages, async getPage(n) {
        const p = await fonte.getPage(n);
        return { rotate: p.rotate, getViewport: o => p.getViewport(o), render: o => p.render({ ...o, intent: "print", annotationMode: 1, background: "rgb(255,255,255)" }) };
    } };
}
