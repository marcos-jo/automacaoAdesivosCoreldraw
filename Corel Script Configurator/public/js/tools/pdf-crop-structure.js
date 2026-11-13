import { PDFName, PDFDict, PDFArray, PDFNumber, PDFRef, PDFString, PDFHexString, PDFBool, PDFNull, PDFRawStream, decodePDFRawStream } from "/vendor/pdf-lib/pdf-lib.js";

// Perfil inicial: uma página, MCIDs no próprio Contents e ParentTree /Nums.
// Não repara árvores nem prepara cópias sem marcação. Não certifica PDF/UA.
const N = PDFName.of;
const CAIXAS = new Set(["MediaBox", "CropBox", "TrimBox", "BleedBox", "ArtBox"]);
const PAPEIS = new Set("Document Part Art Sect Div BlockQuote Caption TOC TOCI Index NonStruct Private P H H1 H2 H3 H4 H5 H6 L LI Lbl LBody Table TR TH TD THead TBody TFoot Span Quote Note Reference BibEntry Code Link Figure Formula".split(" "));
const LIMITES = { profundidade: 64, objetos: 20000, elementos: 5000, mcids: 10000, bytes: 8 * 1024 * 1024, tokens: 500000, tempo: 1500 };
const OPERADORES = new Set("q Q cm w J j M d ri i gs m l c v y h re S s f F f* B B* b b* n W W* BT ET Tc Tw Tz TL Tf Tr Ts Td TD Tm T* Tj TJ ' \" d0 d1 CS cs SC SCN sc scn G g RG rg K k sh Do MP BMC BDC EMC BX EX".split(" "));
const falhar = motivo => { throw Object.assign(new Error(`Recorte bloqueado: estrutura marcada não suportada (${motivo}).`), { codigo: 'PDF_RECORTE_ESTRUTURA' }); };

// Lexer limitado de conteúdo: strings, comentários e imagens não podem ser
// confundidos com operadores BDC/MCID. Inline images e propriedades nomeadas
// ficam fora do perfil. Nenhum texto ou operador é reescrito.
function mcidsDoConteudo(bytes, verificar) {
    let i = 0, tokens = 0;
    const branco = c => /[\x00\t\n\f\r ]/.test(c);
    const delim = c => !c || branco(c) || "()<>[]{}/%".includes(c);
    let texto = "";
    for (let j = 0; j < bytes.length; j += 8192) texto += String.fromCharCode(...bytes.subarray(j, j + 8192));
    function token(depth = 0) {
        verificar();
        if (++tokens > LIMITES.tokens || depth > 32) falhar("limite do conteúdo");
        while (i < texto.length) {
            if (branco(texto[i])) { i++; continue; }
            if (texto[i] !== "%") break;
            while (i < texto.length && !"\r\n".includes(texto[i])) i++;
        }
        if (i === texto.length) return null;
        const c = texto[i++];
        if (c === "(") {
            let nivel = 1;
            while (i < texto.length && nivel) {
                const ch = texto[i++];
                if (ch === "\\") i++;
                else if (ch === "(") { if (++nivel > 32) falhar("string profunda"); }
                else if (ch === ")") nivel--;
            }
            if (nivel) falhar("string incompleta");
            return { tipo: "string" };
        }
        if (c === "<" && texto[i] !== "<") {
            while (i < texto.length && texto[i] !== ">") { if (!/[0-9a-fA-F\s]/.test(texto[i++])) falhar("string hexadecimal inválida"); }
            if (texto[i++] !== ">") falhar("string hexadecimal incompleta");
            return { tipo: "string" };
        }
        if (c === "[" || (c === "<" && texto[i] === "<")) {
            const dict = c === "<"; if (dict) i++;
            const valores = [];
            while (true) {
                const t = token(depth + 1);
                if (!t) falhar("objeto de conteúdo incompleto");
                if (t.tipo === "fim" && t.valor === (dict ? ">>" : "]")) break;
                if (t.tipo === "fim" || t.tipo === "operador") falhar("objeto de conteúdo inválido");
                valores.push(t);
            }
            if (!dict) return { tipo: "array", valores };
            const mapa = new Map();
            for (let n = 0; n < valores.length; n += 2) {
                if (valores[n].tipo !== "nome" || !valores[n + 1] || mapa.has(valores[n].valor)) falhar("propriedade duplicada ou inválida");
                mapa.set(valores[n].valor, valores[n + 1]);
            }
            return { tipo: "dict", mapa };
        }
        if (c === "]") return { tipo: "fim", valor: c };
        if (c === ">" && texto[i++] === ">") return { tipo: "fim", valor: ">>" };
        if ("()>{}".includes(c)) falhar("delimitador de conteúdo inesperado");
        const inicio = c === "/" ? i : i - 1;
        while (i < texto.length && !delim(texto[i])) i++;
        let valor = texto.slice(inicio, i);
        if (c === "/") {
            if (/#(?![0-9a-fA-F]{2})/.test(valor)) falhar("nome PDF inválido");
            valor = valor.replace(/#([0-9a-fA-F]{2})/g, (_, h) => String.fromCharCode(parseInt(h, 16)));
            return { tipo: "nome", valor };
        }
        if (/^[+-]?(?:\d+\.?\d*|\.\d+)$/.test(valor)) return { tipo: "numero", lexema: valor };
        if (["true", "false", "null"].includes(valor)) return { tipo: "literal", valor };
        return { tipo: "operador", valor };
    }
    const ids = new Set(), pilha = []; let args = [];
    for (let t; (t = token());) {
        if (t.tipo !== "operador") { if (t.tipo === "fim" || args.length > 256) falhar("operandos inválidos"); args.push(t); continue; }
        if (["BI", "ID", "EI", "DP"].includes(t.valor)) falhar("imagem inline ou ponto marcado fora do perfil");
        if (!OPERADORES.has(t.valor)) falhar("operador de conteúdo desconhecido");
        if (t.valor === "BDC") {
            if (args.length !== 2 || args[0].tipo !== "nome" || args[1].tipo !== "dict") falhar("BDC fora do perfil");
            const m = args[1].mapa;
            if ([...m.keys()].some(k => !["MCID", "Lang"].includes(k)) || (m.has("Lang") && m.get("Lang").tipo !== "string")) falhar("propriedades BDC fora do perfil");
            const id = m.get("MCID");
            // Inteiro PDF é uma forma lexical, não apenas um Number integral.
            // Sinal opcional e zeros à esquerda são legítimos; reais não são.
            if (!id || id.tipo !== "numero" || !/^[+-]?\d+$/.test(id.lexema)) falhar("MCID exige inteiro PDF");
            const mcid = Number(id.lexema);
            if (!Number.isSafeInteger(mcid) || mcid < 0 || mcid >= LIMITES.mcids || ids.has(mcid)) falhar("MCID inválido ou duplicado");
            ids.add(mcid); pilha.push(true);
        } else if (t.valor === "BMC") {
            if (args.length !== 1 || args[0].tipo !== "nome") falhar("BMC inválido");
            pilha.push(false);
        } else if (t.valor === "EMC") {
            if (args.length || !pilha.length) falhar("EMC sem abertura");
            pilha.pop();
        }
        if (pilha.length > LIMITES.profundidade) falhar("marcação profunda");
        args = [];
    }
    if (pilha.length || args.length) falhar("conteúdo marcado incompleto");
    return ids;
}

export function validarEstruturaRecorte(pdf) {
    if (!pdf.catalog.has(N("StructTreeRoot"))) return null;
    const inicio = Date.now();
    const verificar = () => { if (Date.now() - inicio > LIMITES.tempo) falhar("limite de processamento"); };
    const resolver = valor => {
        verificar();
        if (valor instanceof PDFRef) { const alvo = pdf.context.lookup(valor); if (!alvo || alvo instanceof PDFRef) falhar("referência ausente ou encadeada"); return alvo; }
        return valor;
    };
    const get = (dict, chave) => resolver(dict.get(N(chave)));
    const dict = valor => { const o = resolver(valor); if (!(o instanceof PDFDict) || o instanceof PDFRawStream) falhar("dicionário inválido"); return o; };
    const inteiro = valor => { const o = resolver(valor); if (!(o instanceof PDFNumber) || !Number.isSafeInteger(o.asNumber()) || o.asNumber() < 0 || o.asNumber() >= LIMITES.mcids) falhar("índice inválido"); return o.asNumber(); };
    const chaves = (o, aceitas) => { if (o.keys().some(k => !aceitas.includes(k.decodeText()))) falhar("atributo estrutural fora do perfil"); };
    const texto = o => o instanceof PDFString || o instanceof PDFHexString;
    const paginas = pdf.getPages();
    if (paginas.length !== 1) falhar("somente página única");
    const pagina = paginas[0], refPagina = pagina.ref.toString();
    const objetos = pdf.context.enumerateIndirectObjects();
    if (objetos.length > LIMITES.objetos) falhar("quantidade de objetos");
    if (pagina.node.has(N("Annots")) && pagina.node.Annots()?.size() !== 0) falhar("anotações");
    for (const k of ["AcroForm", "OCProperties", "Perms", "Threads", "OpenAction", "AA"]) if (pdf.catalog.has(N(k))) falhar(k);
    const rootRef = pdf.catalog.get(N("StructTreeRoot"));
    if (!(rootRef instanceof PDFRef)) falhar("raiz deve ser indireta");
    const root = dict(rootRef);
    chaves(root, ["Type", "K", "ParentTree", "ParentTreeNextKey", "RoleMap"]);
    if (get(root, "Type") !== N("StructTreeRoot")) falhar("tipo da raiz");
    const mark = dict(pdf.catalog.get(N("MarkInfo")));
    chaves(mark, ["Marked"]);
    if (get(mark, "Marked") !== PDFBool.True) falhar("MarkInfo");
    if (pdf.catalog.has(N("Lang")) && !texto(get(pdf.catalog, "Lang"))) falhar("idioma inválido");
    const roleMap = root.has(N("RoleMap")) ? dict(root.get(N("RoleMap"))) : null;
    if (roleMap) for (const [, v] of roleMap.entries()) { const papel = resolver(v); if (!(papel instanceof PDFName) || !PAPEIS.has(papel.decodeText())) falhar("RoleMap fora do perfil"); }
    const key = inteiro(pagina.node.get(N("StructParents")));
    const pt = dict(root.get(N("ParentTree"))); chaves(pt, ["Nums"]);
    const nums = get(pt, "Nums");
    if (!(nums instanceof PDFArray) || nums.size() !== 2 || inteiro(nums.get(0)) !== key) falhar("ParentTree deve conter a chave da página");
    const pais = resolver(nums.get(1));
    if (!(pais instanceof PDFArray) || !pais.size() || pais.size() > LIMITES.mcids) falhar("ParentTree inválida");
    if (root.has(N("ParentTreeNextKey")) && inteiro(root.get(N("ParentTreeNextKey"))) <= key) falhar("ParentTreeNextKey inválida");
    const visitados = new Set(), donos = new Map();
    let elementos = 0;
    function caminhar(valor, pai, pg, nivel, raiz = false) {
        verificar(); if (nivel > LIMITES.profundidade) falhar("profundidade da árvore");
        const o = resolver(valor);
        if (o instanceof PDFArray) {
            if (visitados.has(o)) falhar("array cíclico ou compartilhado"); visitados.add(o);
            if (o.size() > LIMITES.elementos) falhar("quantidade de filhos");
            for (const filho of o.asArray()) caminhar(filho, pai, pg, nivel + 1, raiz);
        } else if (o instanceof PDFNumber) {
            const id = inteiro(o);
            if (raiz || pg !== refPagina || donos.has(id) || id >= pais.size() || !(pais.get(id) instanceof PDFRef) || pais.get(id).toString() !== pai.toString()) falhar("vínculo MCID/ParentTree inválido");
            donos.set(id, pai.toString());
        } else {
            if (!(valor instanceof PDFRef) || !(o instanceof PDFDict) || visitados.has(o)) falhar("elemento direto, repetido, cíclico ou desconhecido");
            visitados.add(o); if (++elementos > LIMITES.elementos) falhar("quantidade de elementos");
            chaves(o, ["Type", "S", "P", "Pg", "K", "T", "Lang", "Alt", "ActualText", "E"]);
            if (o.has(N("Type")) && get(o, "Type") !== N("StructElem")) falhar("tipo de elemento");
            const p = o.get(N("P"));
            if (!(p instanceof PDFRef) || p.toString() !== pai.toString() || resolver(p) !== resolver(pai)) falhar("pai inconsistente");
            let papel = get(o, "S");
            if (!(papel instanceof PDFName)) falhar("papel inválido");
            if (roleMap?.has(papel)) papel = resolver(roleMap.get(papel));
            if (!(papel instanceof PDFName) || !PAPEIS.has(papel.decodeText())) falhar("papel desconhecido");
            if (o.has(N("Pg"))) { const ref = o.get(N("Pg")); if (!(ref instanceof PDFRef) || ref.toString() !== refPagina || resolver(ref) !== pagina.node) falhar("página inconsistente"); pg = ref.toString(); }
            for (const k of ["T", "Lang", "Alt", "ActualText", "E"]) if (o.has(N(k)) && !texto(get(o, k))) falhar("texto estrutural inválido");
            if (!o.has(N("K"))) falhar("elemento sem filhos verificáveis");
            caminhar(o.get(N("K")), valor, pg, nivel + 1);
        }
    }
    caminhar(root.get(N("K")), rootRef, null, 0, true);
    if (donos.size !== pais.size()) falhar("ParentTree com entradas sem dono");
    for (let i = 0; i < pais.size(); i++) { resolver(pais.get(i)); if (!donos.has(i)) falhar("lacuna em ParentTree"); }
    let conteudos = get(pagina.node, "Contents");
    conteudos = conteudos instanceof PDFArray ? conteudos.asArray() : [conteudos];
    if (!conteudos.length || conteudos.length > 64) falhar("quantidade de streams");
    const partes = []; let tamanho = 0;
    for (const v of conteudos) {
        const s = resolver(v);
        if (!(s instanceof PDFRawStream) || s.contents.length > LIMITES.bytes) falhar("stream fora do perfil");
        const filtro = get(s.dict, "Filter");
        if (filtro && filtro !== N("FlateDecode")) falhar("filtro de conteúdo fora do perfil");
        if (s.dict.has(N("DecodeParms")) && get(s.dict, "DecodeParms") !== PDFNull) falhar("parâmetros de conteúdo fora do perfil");
        const dados = decodePDFRawStream(s).getBytes(LIMITES.bytes - tamanho + 1);
        tamanho += dados.length; if (tamanho > LIMITES.bytes) falhar("tamanho do conteúdo");
        partes.push(dados);
    }
    const juntos = new Uint8Array(tamanho); let pos = 0;
    for (const dados of partes) { juntos.set(dados, pos); pos += dados.length; }
    const ids = mcidsDoConteudo(juntos, verificar);
    if (ids.size !== donos.size || [...ids].some(id => !donos.has(id))) falhar("MCIDs do conteúdo não correspondem à árvore");

    // Manifesto dos objetos existentes: ignora apenas caixas da página alvo e
    // contêineres de serialização. Compara bytes dos streams sem rasterização.
    let passos = 0, bytesStreams = 0;
    function representar(o, nivel = 0, ignoradas = null) {
        verificar(); if (++passos > 200000 || nivel > LIMITES.profundidade) falhar("limite do manifesto");
        if (o instanceof PDFRef) { resolver(o); return o.toString(); }
        if (o instanceof PDFArray) return "[" + o.asArray().map(v => representar(v, nivel + 1)).join(" ") + "]";
        if (o instanceof PDFDict) {
            if ((get(o, "Type") === N("StructElem") && !visitados.has(o)) || (get(o, "Type") === N("StructTreeRoot") && o !== root)) falhar("objeto estrutural fora da árvore");
            if (get(o, "Subtype") === N("Form") || (o !== pagina.node && (o.has(N("StructParents")) || o.has(N("StructParent"))))) falhar("XObject ou vínculo estrutural fora da página");
            return "<<" + o.keys().filter(k => !ignoradas?.has(k.decodeText())).sort((a,b) => a.toString().localeCompare(b.toString())).map(k => k.toString() + ":" + representar(o.get(k), nivel + 1)).join("|") + ">>";
        }
        if (!(o instanceof PDFName || o instanceof PDFNumber || texto(o) || o instanceof PDFBool || o === PDFNull)) falhar("objeto desconhecido");
        return o.toString();
    }
    const manifesto = objetos.filter(([, o]) => !(o instanceof PDFRawStream && [N("ObjStm"), N("XRef")].includes(get(o.dict, "Type")))).map(([ref, o]) => {
        if (o instanceof PDFRawStream) {
            bytesStreams += o.contents.length; if (bytesStreams > 32 * 1024 * 1024) falhar("tamanho dos recursos");
            // Length é recalculado pelo gravador; os bytes são comparados abaixo.
            return { ref: ref.toString(), objeto: representar(o.dict, 0, new Set(["Length"])), bytes: o.contents.slice() };
        }
        return { ref: ref.toString(), objeto: representar(o, 0, o === pagina.node ? CAIXAS : null) };
    });
    return { elementos, mcids: ids.size, manifesto };
}

export function conferirEstruturaRecorte(antes, pdf) {
    const depois = validarEstruturaRecorte(pdf);
    if (!antes || !depois || antes.elementos !== depois.elementos || antes.mcids !== depois.mcids || antes.manifesto.length !== depois.manifesto.length) falhar("estrutura alterada na serialização");
    for (let i = 0; i < antes.manifesto.length; i++) {
        const a = antes.manifesto[i], b = depois.manifesto[i];
        if (a.ref !== b.ref || a.objeto !== b.objeto || !!a.bytes !== !!b.bytes || (a.bytes && (a.bytes.length !== b.bytes.length || a.bytes.some((v, n) => v !== b.bytes[n])))) falhar("objeto ou stream alterado na serialização");
    }
    return depois;
}
