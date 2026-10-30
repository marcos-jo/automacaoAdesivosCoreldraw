import { PDFName, PDFArray, PDFDict, PDFNumber, PDFContentStream, degrees,
    pushGraphicsState, popGraphicsState, concatTransformationMatrix, rectangle, clip, endPath
} from "/vendor/pdf-lib/pdf-lib.js";
import { medidaMm } from "./pdf-formats.js";
import { diagnosticarImpressao, prepararParaImpressao, validarActualTextImpressao, PDFCompatibilidadeErro } from "./pdf-print-compatibility.js";

export const mmParaPontos = mm => mm * 72 / 25.4;
const nome = PDFName.of;
export function verificarCancelamento(signal) { if (signal?.aborted) throw new DOMException("Prévia cancelada", "AbortError"); }
export function validarRedimensionamento(config, total) {
    const largura = medidaMm(config.largura), altura = medidaMm(config.altura);
    if (!["ajustar", "centralizar", "preencher"].includes(config.modo)) throw new Error("Modo de redimensionamento inválido.");
    const paginas = [...new Set(config.paginas || [])];
    if (!paginas.length || paginas.some(n => !Number.isInteger(n) || n < 1 || n > total)) throw new Error("Seleção de páginas inválida.");
    return { largura, altura, modo: config.modo, paginas, modoImpressao: config.modoImpressao === true };
}
function intersecao(a, b) {
    const x = Math.max(a.x, b.x), y = Math.max(a.y, b.y);
    const largura = Math.min(a.x + a.width, b.x + b.width) - x;
    const altura = Math.min(a.y + a.height, b.y + b.height) - y;
    return largura > 0 && altura > 0 ? { x, y, width: largura, height: altura } : null;
}
function ponto(m, x, y) { return [m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]]; }
function transformarRetangulo(m, r) {
    const pontos = [ponto(m, r.x, r.y), ponto(m, r.x + r.width, r.y), ponto(m, r.x, r.y + r.height), ponto(m, r.x + r.width, r.y + r.height)];
    const xs = pontos.map(p => p[0]), ys = pontos.map(p => p[1]);
    return { x: Math.min(...xs), y: Math.min(...ys), width: Math.max(...xs) - Math.min(...xs), height: Math.max(...ys) - Math.min(...ys) };
}
function lerRetangulo(array, permitirVazio = false) {
    if (!(array instanceof PDFArray) || array.size() !== 4) throw new Error("Retângulo PDF inválido.");
    const [x, y, x2, y2] = array.asArray().map(v => array.context.lookup(v).asNumber());
    if (![x, y, x2, y2].every(Number.isFinite) || !permitirVazio && (x2 === x || y2 === y)) throw new Error("Caixa PDF sem área válida.");
    return { x: Math.min(x, x2), y: Math.min(y, y2), width: Math.abs(x2 - x), height: Math.abs(y2 - y) };
}

export function calcularTransformacao(crop, rotacao, unidade, largura, altura, modo) {
    if (![crop.x, crop.y, crop.width, crop.height, unidade].every(Number.isFinite) || crop.width <= 0 || crop.height <= 0 || unidade <= 0 || unidade > 75000) throw new Error("Geometria ou UserUnit inválido.");
    const r = ((rotacao % 360) + 360) % 360;
    if (![0, 90, 180, 270].includes(r)) throw new Error("Rotação da página não suportada.");
    const w = (r % 180 ? crop.height : crop.width) * unidade;
    const h = (r % 180 ? crop.width : crop.height) * unidade;
    const W = mmParaPontos(largura), H = mmParaPontos(altura);
    const escala = modo === "centralizar" ? 1 : Math[modo === "preencher" ? "max" : "min"](W / w, H / h);
    const dx = (W - w * escala) / 2, dy = (H - h * escala) / 2;
    const { x, y, width, height } = crop;
    const orientacao = r === 0 ? [1, 0, 0, 1, -x, -y]
        : r === 90 ? [0, -1, 1, 0, -y, x + width]
        : r === 180 ? [-1, 0, 0, -1, x + width, y + height]
        : [0, 1, -1, 0, y + height, -x];
    const matriz = orientacao.map(v => v * escala * unidade);
    matriz[4] += dx; matriz[5] += dy;
    return { matriz, escala: escala * unidade, destino: { x: 0, y: 0, width: W, height: H },
        original: { largura: w * 25.4 / 72, altura: h * 25.4 / 72 },
        corte: w * escala > W + 1e-6 || h * escala > H + 1e-6,
        semAlteracao: Math.abs(W - w) < 1e-6 && Math.abs(H - h) < 1e-6 };
}

function planoPagina(pagina, numero, config) {
    const media = lerRetangulo(pagina.node.MediaBox());
    const crop = intersecao(media, lerRetangulo(pagina.node.CropBox() || pagina.node.MediaBox()));
    if (!crop) throw new Error(`Página ${numero}: CropBox fora da MediaBox.`);
    const unidade = pagina.node.lookupMaybe(nome("UserUnit"), PDFNumber)?.asNumber() ?? 1;
    const plano = { ...calcularTransformacao(crop, pagina.getRotation().angle, unidade, config.largura, config.altura, config.modo), pagina, numero, crop, caixas: [], anotacoes: [] };
    for (const chave of ["TrimBox", "BleedBox", "ArtBox"]) {
        const valor = pagina.node.lookupMaybe(nome(chave), PDFArray);
        if (!valor) continue; // Não fabricar caixas gráficas ausentes.
        const visivel = intersecao(lerRetangulo(valor), crop);
        const caixa = visivel && intersecao(transformarRetangulo(plano.matriz, visivel), plano.destino);
        if (!caixa) throw new Error(`Página ${numero}: ${chave} ficaria sem área visível. Revise o formato, o modo e as caixas do PDF de origem.`);
        plano.caixas.push({ chave, caixa });
    }
    for (const ref of pagina.node.Annots()?.asArray() || []) {
        const anotacao = pagina.doc.context.lookup(ref, PDFDict);
        const tipo = anotacao.lookup(nome("Subtype"), PDFName).asString();
        const acao = anotacao.lookupMaybe(nome("A"), PDFDict);
        if (!["/Link", "/Text", "/Popup"].includes(tipo) || anotacao.has(nome("AP")) || anotacao.has(nome("AA")) ||
            acao && !["/URI", "/GoToR"].includes(acao.lookup(nome("S"), PDFName).asString())) {
            throw new Error(`Página ${numero}: ${tipo} ou aparência interativa complexa não suportada. Nenhuma alteração aplicada.`);
        }
        plano.anotacoes.push(anotacao);
    }
    return plano;
}
function transformarAnotacao(plano, anotacao) {
    const contexto = plano.pagina.doc.context;
    const original = lerRetangulo(anotacao.lookup(nome("Rect"), PDFArray), true);
    const recortado = intersecao(original, plano.crop);
    const rect = recortado && intersecao(transformarRetangulo(plano.matriz, recortado), plano.destino);
    anotacao.set(nome("Rect"), contexto.obj(rect ? [rect.x, rect.y, rect.x + rect.width, rect.y + rect.height] : [0, 0, 0, 0]));
    const quads = anotacao.lookupMaybe(nome("QuadPoints"), PDFArray);
    if (quads) {
        if (quads.size() % 8) throw new Error("QuadPoints inválidos.");
        const valores = quads.asArray().map(v => contexto.lookup(v).asNumber());
        const novos = [];
        for (let i = 0; i < valores.length; i += 2) novos.push(...ponto(plano.matriz, valores[i], valores[i + 1]));
        anotacao.set(nome("QuadPoints"), contexto.obj(novos));
    }
    const border = anotacao.lookupMaybe(nome("Border"), PDFArray);
    if (border) {
        const valores = border.asArray().map((v, i) => i < 3 ? contexto.obj(contexto.lookup(v).asNumber() * plano.escala)
            : contexto.obj(contexto.lookup(v, PDFArray).asArray().map(n => contexto.lookup(n).asNumber() * plano.escala)));
        if (plano.matriz[0] === 0) [valores[0], valores[1]] = [valores[1], valores[0]];
        anotacao.set(nome("Border"), contexto.obj(valores));
    }
    const bs = anotacao.lookupMaybe(nome("BS"), PDFDict);
    if (bs?.has(nome("W"))) {
        const copia = bs.clone(contexto);
        copia.set(nome("W"), contexto.obj(bs.lookup(nome("W")).asNumber() * plano.escala));
        const dash = bs.lookupMaybe(nome("D"), PDFArray);
        if (dash) copia.set(nome("D"), contexto.obj(dash.asArray().map(n => contexto.lookup(n).asNumber() * plano.escala)));
        anotacao.set(nome("BS"), copia);
    }
}
function aplicarPlano(plano) {
    const { pagina, matriz, crop, destino } = plano, contexto = pagina.doc.context;
    pagina.pushOperators(); // Normaliza streams existentes e inclui páginas vazias.
    const inicio = PDFContentStream.of(contexto.obj({}), [pushGraphicsState(), concatTransformationMatrix(...matriz), rectangle(crop.x, crop.y, crop.width, crop.height), clip(), endPath()]);
    const fim = PDFContentStream.of(contexto.obj({}), [popGraphicsState()]);
    if (!pagina.node.wrapContentStreams(contexto.register(inicio), contexto.register(fim))) throw new Error("Não foi possível envolver o conteúdo PDF.");
    pagina.setMediaBox(0, 0, destino.width, destino.height);
    pagina.setCropBox(0, 0, destino.width, destino.height);
    pagina.setRotation(degrees(0));
    pagina.node.set(nome("UserUnit"), contexto.obj(1));
    for (const { chave, caixa } of plano.caixas) pagina.node.set(nome(chave), contexto.obj([caixa.x, caixa.y, caixa.x + caixa.width, caixa.y + caixa.height]));
    for (const anotacao of plano.anotacoes) transformarAnotacao(plano, anotacao);
}

// Recebe apenas uma cópia transacional, nunca o PDF definitivo do workspace.
export async function redimensionarDocumento(pdf, configuracao, { signal, validarImpressao } = {}) {
    verificarCancelamento(signal);
    const config = validarRedimensionamento(configuracao, pdf.getPageCount());
    const diagnostico = diagnosticarImpressao(pdf, config.paginas);
    await validarActualTextImpressao(pdf, diagnostico, validarImpressao, { signal });
    if (diagnostico.bloqueios.length || diagnostico.removiveis.length && !config.modoImpressao) throw new PDFCompatibilidadeErro(diagnostico);
    const removidos = config.modoImpressao ? prepararParaImpressao(pdf, diagnostico) : [];
    const planos = config.paginas.map(n => planoPagina(pdf.getPage(n - 1), n, config));
    const avisos = new Set();
    if (removidos.length) avisos.add(`Modo de impressão: descartados no documento inteiro: ${removidos.join("; ")}. O conteúdo gráfico não foi rasterizado.`);
    for (let i = 0; i < planos.length; i++) {
        verificarCancelamento(signal);
        const plano = planos[i];
        if (plano.corte) avisos.add("Há risco de corte: parte da área originalmente visível ficará fora do novo formato.");
        if (plano.caixas.length) avisos.add("Caixas gráficas existentes acompanham o conteúdo e são limitadas à área visível. Sangria não foi criada nem validada.");
        if (plano.anotacoes.length) avisos.add("Links externos e notas simples foram reposicionados. Aparência dos ícones de notas depende do leitor PDF.");
        if (!plano.semAlteracao) aplicarPlano(plano);
        if (i % 10 === 9) await new Promise(resolve => setTimeout(resolve, 0));
    }
    verificarCancelamento(signal);
    return { config, avisos: [...avisos], removidos, validacaoActualText: diagnostico.validacaoActualText, preparadoParaImpressao: removidos.length > 0, alteradas: planos.filter(p => !p.semAlteracao).map(p => p.pagina.ref.toString()),
        detalhes: planos.map(p => ({ numero: p.numero, original: p.original, nova: { largura: config.largura, altura: config.altura }, corte: p.corte })) };
}
