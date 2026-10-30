import { PDFName, PDFDict, PDFArray, PDFNumber, PDFString, PDFHexString, PDFStream, PDFPageLeaf, PDFNull } from "/vendor/pdf-lib/pdf-lib.js";

const nome = PDFName.of;
const tipo = (dict, chave) => dict.lookupMaybe(nome(chave), PDFName)?.asString();
const numero = (dict, chave, padrao = 0) => dict.lookupMaybe(nome(chave), PDFNumber)?.asNumber() ?? padrao;
function retangulo(array) {
    if (!(array instanceof PDFArray) || array.size() !== 4) return null;
    const v = array.asArray().map(x => array.context.lookup(x)?.asNumber?.());
    return v.every(Number.isFinite) ? [Math.min(v[0], v[2]), Math.min(v[1], v[3]), Math.max(v[0], v[2]), Math.max(v[1], v[3])] : null;
}
const temArea = rect => !!rect && rect[2] > rect[0] && rect[3] > rect[1];

export function aparenciaNormal(anotacao) {
    const ap = anotacao.lookupMaybe(nome("AP"), PDFDict);
    let normal = ap?.lookup(nome("N"));
    if (normal instanceof PDFDict) {
        const estado = anotacao.lookupMaybe(nome("AS"), PDFName);
        normal = estado ? normal.lookup(estado) : null;
    }
    return normal instanceof PDFStream && normal.getContents().length > 0 && temArea(retangulo(normal.dict.lookup(nome("BBox"))));
}

function evidenciaValor(valor, tamanhoArquivo) {
    if (!valor || valor === PDFNull) return "vazio";
    if (!(valor instanceof PDFDict)) return "inconclusiva";
    const conteudo = valor.lookup(nome("Contents")), faixa = valor.lookup(nome("ByteRange"));
    if (!(conteudo instanceof PDFHexString || conteudo instanceof PDFString) || !conteudo.asBytes().some(b => b !== 0) || !(faixa instanceof PDFArray)) return "inconclusiva";
    const v = faixa.asArray().map(x => faixa.context.lookup(x)?.asNumber?.());
    if (v.length < 4 || v.length % 2 || v[0] !== 0 || v.some(n => !Number.isSafeInteger(n) || n < 0)) return "inconclusiva";
    let fim = 0, lacuna = false;
    for (let i = 0; i < v.length; i += 2) {
        if (v[i] < fim) return "inconclusiva";
        if (v[i] > fim) lacuna = true;
        fim = v[i] + v[i + 1];
        if (!Number.isSafeInteger(fim) || tamanhoArquivo !== undefined && fim > tamanhoArquivo) return "inconclusiva";
    }
    return lacuna && v[1] > 0 ? "presente" : "inconclusiva";
}

export function dicionariosPDF(pdf) {
    const fila = pdf.context.enumerateIndirectObjects().map(([, objeto]) => objeto), vistos = new Set(), encontrados = [];
    while (fila.length) {
        let objeto = fila.pop();
        if (objeto instanceof PDFStream) objeto = objeto.dict;
        if (!objeto || vistos.has(objeto)) continue;
        vistos.add(objeto);
        if (objeto instanceof PDFDict) {
            encontrados.push(objeto);
            for (const [, v] of objeto.entries()) if (v instanceof PDFDict || v instanceof PDFArray || v instanceof PDFStream) fila.push(v);
        } else if (objeto instanceof PDFArray) for (const v of objeto.asArray()) fila.push(v);
    }
    return encontrados;
}

function herancaCampo(dict) {
    const cadeia = [], vistos = new Set();
    let atual = dict;
    while (atual instanceof PDFDict && !vistos.has(atual) && cadeia.length < 128) {
        vistos.add(atual); cadeia.push(atual); atual = atual.lookup(nome("Parent"));
    }
    if (atual instanceof PDFDict) throw new Error("Ciclo/profundidade excessiva na árvore de campos.");
    const ft = cadeia.find(d => d.has(nome("FT")));
    const valor = cadeia.find(d => d.has(nome("V")));
    const dono = valor || cadeia.find(d => d.has(nome("T"))) || ft || dict;
    return { ft: ft && tipo(ft, "FT"), valor: valor?.lookup(nome("V")), dono };
}

// Leitura estrutural, SEM save(), execução de scripts ou validação de certificado.
// Chamar na origem, antes da primeira reserialização, e no resultado reconstruído.
export function diagnosticarAssinaturas(pdf, { tamanhoArquivo } = {}) {
    try {
        const objetos = pdf.context.enumerateIndirectObjects(), refs = new Map(objetos.map(([r, o]) => [o, r.toString()]));
        const entradas = new Map(), indicios = [], widgets = new Map(), paginas = [];
        pdf.catalog.Pages().traverse((node, ref) => { if (node instanceof PDFPageLeaf) paginas.push({ node, id: ref.toString(), numero: paginas.length + 1 }); });
        const registrar = (dono, valor, origem, orfa = false) => {
            const chave = valor instanceof PDFDict ? valor : dono;
            if (!entradas.has(chave)) entradas.set(chave, {
                id: refs.get(chave) || `direto-${entradas.size + 1}`, estado: orfa ? "inconclusiva" : evidenciaValor(valor, tamanhoArquivo),
                origem, widgets: [], objetoOrfao: orfa
            });
            return entradas.get(chave);
        };
        const campo = (dict, orfa = false) => {
            const herdado = herancaCampo(dict);
            return herdado.ft === "/Sig" ? registrar(herdado.dono, herdado.valor, orfa ? "campo /FT /Sig sem associação atual" : "campo /FT /Sig", orfa) : null;
        };
        const formulario = pdf.catalog.lookupMaybe(nome("AcroForm"), PDFDict);
        if (formulario?.has(nome("SigFlags")) && numero(formulario, "SigFlags") !== 0) indicios.push("AcroForm/SigFlags (indício; sozinho não comprova assinatura)");
        if (formulario?.has(nome("XFA"))) indicios.push("XFA: árvore dinâmica não analisada integralmente");
        const fila = [...(formulario?.lookupMaybe(nome("Fields"), PDFArray)?.asArray() || [])], vistos = new Set();
        while (fila.length) {
            const dict = pdf.context.lookup(fila.pop(), PDFDict);
            if (vistos.has(dict)) { indicios.push("Árvore de campos compartilhada/cíclica"); continue; }
            vistos.add(dict);
            const filhos = dict.lookupMaybe(nome("Kids"), PDFArray)?.asArray() || [];
            if (!filhos.length || filhos.every(r => { const d = pdf.context.lookup(r, PDFDict); return tipo(d, "Subtype") === "/Widget" && !d.has(nome("T")); })) campo(dict);
            if (tipo(dict, "Subtype") === "/Widget") widgets.set(dict, campo(dict));
            for (const filho of filhos) fila.push(filho);
        }
        const permissoes = pdf.catalog.lookupMaybe(nome("Perms"), PDFDict);
        if (permissoes) for (const [chave, valor] of permissoes.entries()) {
            const d = pdf.context.lookup(valor);
            if (d instanceof PDFDict) registrar(d, d, `catálogo /Perms${chave.asString()}`);
            else indicios.push("Referência não reconhecida em Perms");
        }
        for (const pagina of paginas) for (const ref of pagina.node.Annots()?.asArray() || []) {
            const dict = pdf.context.lookup(ref, PDFDict);
            if (tipo(dict, "Subtype") !== "/Widget") continue;
            const entrada = campo(dict); widgets.set(dict, entrada);
            if (!entrada) continue;
            const rect = retangulo(dict.lookup(nome("Rect"))), flags = numero(dict, "F");
            const crop = retangulo(pagina.node.CropBox() || pagina.node.MediaBox());
            const visivel = temArea(rect) && !(flags & 3) && (!crop || rect[2] > crop[0] && rect[0] < crop[2] && rect[3] > crop[1] && rect[1] < crop[3]);
            entrada.widgets.push({ id: refs.get(dict) || ref.toString(), pagina: pagina.numero, paginaId: pagina.id,
                rect, visivel, retanguloValido: !!rect, imprimivel: !!(flags & 4) && !(flags & 3), aparencia: aparenciaNormal(dict), flags });
        }
        // Campos/widget órfãos e dicionários de revisões antigas são indícios,
        // separados das assinaturas associadas ao formulário/catálogo atual.
        for (const objeto of dicionariosPDF(pdf)) {
            if (tipo(objeto, "FT") === "/Sig" && !vistos.has(objeto) && !widgets.has(objeto)) campo(objeto, true);
            if (["/Sig", "/DocTimeStamp"].includes(tipo(objeto, "Type")) || objeto.has(nome("ByteRange")) && objeto.has(nome("Contents"))) {
                if (!entradas.has(objeto)) registrar(objeto, objeto, "dicionário sem associação atual ao formulário/catálogo", true);
            }
        }
        const assinaturas = [...entradas.values()].map(e => ({ ...e,
            invisivel: e.estado !== "vazio" && !e.widgets.some(w => w.visivel),
            paginasVisuais: [...new Set(e.widgets.filter(w => w.visivel).map(w => w.pagina))]
        }));
        const presentes = assinaturas.filter(s => s.estado === "presente").length;
        const vazios = assinaturas.filter(s => s.estado === "vazio").length;
        const inconclusivas = assinaturas.filter(s => s.estado === "inconclusiva").length;
        return { estado: presentes ? "detectada" : inconclusivas || indicios.length && !vazios ? "inconclusiva" : vazios ? "campos_vazios" : "nao_identificada",
            presentes, vazios, inconclusivas, assinaturas, indicios, validade: "Validade criptográfica não verificada" };
    } catch (erro) {
        return { estado: "inconclusiva", presentes: 0, vazios: 0, inconclusivas: 1, assinaturas: [], indicios: [`Detecção estrutural incompleta: ${erro.message}`], validade: "Validade criptográfica não verificada" };
    }
}

export function verificarSemAssinaturas(pdf) {
    const diagnostico = diagnosticarAssinaturas(pdf);
    if (diagnostico.estado !== "nao_identificada" || pdf.catalog.has(nome("Perms")) || pdf.catalog.has(nome("AcroForm"))) throw new Error("A cópia reconstruída ainda contém campos, estruturas de assinatura ou referências inconclusivas.");
    if (dicionariosPDF(pdf).some(d => tipo(d, "FT") === "/Sig" || ["/Sig", "/DocTimeStamp"].includes(tipo(d, "Type")) || d.has(nome("ByteRange")) || d.has(nome("SigFlags")))) throw new Error("Referência de assinatura remanescente nos objetos da cópia.");
    return diagnostico;
}
