// Validação excepcional de ActualText, antes de qualquer transformação geométrica.
// Os canvases servem apenas à comparação; nunca entram no PDF exportado.
const normalizar = texto => texto.normalize("NFC").replace(/\s+/gu, " ").trim();
const cancelarSeNecessario = signal => { if (signal?.aborted) throw new DOMException("Validação cancelada", "AbortError"); };

export function indexarTextoMarcado(conteudo) {
    const pilha = [], marcas = new Map();
    for (const item of conteudo.items) {
        if (item.type === "beginMarkedContent" || item.type === "beginMarkedContentProps") {
            const id = item.id?.match(/_mc(\d+)$/)?.[1] ?? null;
            pilha.push(id);
            if (id !== null) {
                const entrada = marcas.get(Number(id)) || { texto: "", ocorrencias: 0 };
                ++entrada.ocorrencias; marcas.set(Number(id), entrada);
            }
        } else if (item.type === "endMarkedContent") pilha.pop();
        else if (typeof item.str === "string") {
            for (const id of new Set(pilha.filter(id => id !== null))) {
                marcas.get(Number(id)).texto += item.str + (item.hasEOL ? "\n" : "");
            }
        }
    }
    return marcas;
}

export function verificarRedundanciaActualText(substitutos, marcasPorPagina) {
    for (const [indice, substituto] of substitutos.entries()) {
        if (substituto.erro || !substituto.conteudos.length) return { compativel: false, motivo: `ActualText ${indice + 1}: ${substituto.erro || "sem conteúdo marcado verificável"} Operação bloqueada.` };
        const partes = substituto.conteudos.map(c => marcasPorPagina.get(c.pagina)?.get(c.mcid));
        if (partes.some(p => !p || p.ocorrencias !== 1) ||
            normalizar(substituto.texto) !== normalizar(partes.map(p => p.texto).join(""))) {
            return { compativel: false, motivo: `ActualText ${indice + 1}: texto substituto não é redundante no conteúdo marcado correspondente, ou seu MCID é ambíguo. Pode ser essencial à extração de caracteres; operação bloqueada.` };
        }
    }
    return { compativel: true };
}

async function aguardar(promise, signal, cancelar = () => {}) {
    let interromper, timer;
    const interrupcao = new Promise((_, reject) => {
        interromper = () => reject(new DOMException("Validação cancelada", "AbortError"));
        timer = setTimeout(() => reject(new Error("Tempo limite na comparação de texto/renderização.")), 30000);
    });
    signal?.addEventListener("abort", interromper, { once: true });
    try { cancelarSeNecessario(signal); return await Promise.race([promise, interrupcao]); }
    catch (erro) { cancelar(); throw erro; }
    finally { clearTimeout(timer); signal?.removeEventListener("abort", interromper); }
}

export async function validarCompatibilidadeImpressao({ antes, depois, substitutos, signal }, {
    abrir, fechar, criarCanvas = () => document.createElement("canvas")
}) {
    const candidatos = [], marcas = new Map();
    let componentes = 0, caracteres = 0;
    try {
        cancelarSeNecessario(signal);
        const original = await abrir(antes, { signal }); candidatos.push(original);
        const preparado = await abrir(depois, { signal }); candidatos.push(preparado);
        if (!original?.documento || !preparado?.documento) throw new Error("Comparação sem documento PDF disponível.");
        if (original.documento.numPages !== preparado.documento.numPages) throw new Error("A preparação alterou a quantidade de páginas.");
        for (let numero = 1; numero <= original.documento.numPages; numero++) {
            cancelarSeNecessario(signal);
            const a = await aguardar(original.documento.getPage(numero), signal);
            const b = await aguardar(preparado.documento.getPage(numero), signal);
            const canvases = [];
            try {
                const textoA = await aguardar(a.getTextContent({ includeMarkedContent: true, disableNormalization: true }), signal);
                const textoB = await aguardar(b.getTextContent({ includeMarkedContent: true, disableNormalization: true }), signal);
                const textos = conteudo => conteudo.items.filter(i => typeof i.str === "string").map(i => [i.str, !!i.hasEOL]);
                if (JSON.stringify(textos(textoA)) !== JSON.stringify(textos(textoB))) return { compativel: false, motivo: `Página ${numero}: a preparação alterou o texto extraível. Operação bloqueada.` };
                caracteres += textos(textoA).reduce((total, [texto]) => total + texto.length, 0);
                if (substitutos.some(s => s.conteudos.some(c => c.pagina === numero))) marcas.set(numero, indexarTextoMarcado(textoA));
                const verificaveis = substitutos.filter(s => s.erro || s.conteudos.every(c => c.pagina <= numero));
                const redundanciaParcial = verificarRedundanciaActualText(verificaveis, marcas);
                if (!redundanciaParcial.compativel) return redundanciaParcial;
                // Canonical 72 dpi, sem amostragem de páginas ou redução silenciosa.
                const va = a.getViewport({ scale: 1 }), vb = b.getViewport({ scale: 1 });
                if (va.width !== vb.width || va.height !== vb.height || va.rotation !== vb.rotation) throw new Error(`Página ${numero}: geometria alterada durante a preparação.`);
                const largura = Math.ceil(va.width), altura = Math.ceil(va.height);
                if (!Number.isFinite(largura * altura) || largura <= 0 || altura <= 0 || largura > 8192 || altura > 8192 || largura * altura > 8000000) throw new Error(`Página ${numero}: tamanho excede o limite de comparação visual integral (8 milhões de pixels/8192 por lado).`);
                for (const [pagina, viewport] of [[a, va], [b, vb]]) {
                    const canvas = criarCanvas(); canvases.push(canvas); canvas.width = largura; canvas.height = altura;
                    const contexto = canvas.getContext("2d", { willReadFrequently: true });
                    if (!contexto) throw new Error("Canvas indisponível para validar aparência.");
                    const tarefa = pagina.render({ canvas, canvasContext: contexto, viewport, intent: "print" });
                    await aguardar(tarefa.promise, signal, () => tarefa.cancel());
                }
                const pixelsA = canvases[0].getContext("2d").getImageData(0, 0, largura, altura).data;
                const pixelsB = canvases[1].getContext("2d").getImageData(0, 0, largura, altura).data;
                for (let i = 0; i < pixelsA.length; i++) if (pixelsA[i] !== pixelsB[i]) return { compativel: false, motivo: `Página ${numero}: a preparação alterou a renderização. Operação bloqueada.` };
                componentes += pixelsA.length;
            } finally {
                for (const canvas of canvases) canvas.width = canvas.height = 0;
                a.cleanup?.(); b.cleanup?.();
            }
            await new Promise(resolve => setTimeout(resolve, 0));
        }
        cancelarSeNecessario(signal);
        const redundancia = verificarRedundanciaActualText(substitutos, marcas);
        if (!redundancia.compativel) return redundancia;
        return { compativel: true, paginas: original.documento.numPages, substituicoes: substitutos.length, componentes, caracteres, dpi: 72 };
    } finally {
        const resultados = await Promise.allSettled(candidatos.filter(Boolean).map(fechar));
        const falha = resultados.find(r => r.status === "rejected");
        if (falha) throw falha.reason;
    }
}
