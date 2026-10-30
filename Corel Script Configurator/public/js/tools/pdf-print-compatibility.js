import { PDFName, PDFArray, PDFDict, PDFStream, PDFNumber, PDFDocument } from "/vendor/pdf-lib/pdf-lib.js";

const nome = PDFName.of;
const tipoAcao = acao => acao?.lookupMaybe(nome("S"), PDFName)?.asString();

// Percorre dicionários diretos e de streams sem ler/modificar os bytes de conteúdo.
function dicionarios(pdf) {
    const vistos = new Set(), encontrados = [], fila = pdf.context.enumerateIndirectObjects().map(([, objeto]) => objeto);
    while (fila.length) {
        let objeto = fila.pop();
        if (objeto instanceof PDFStream) objeto = objeto.dict;
        if (!objeto || vistos.has(objeto)) continue;
        vistos.add(objeto);
        if (objeto instanceof PDFDict) {
            encontrados.push(objeto);
            for (const [, valor] of objeto.entries()) if (valor instanceof PDFDict || valor instanceof PDFArray || valor instanceof PDFStream) fila.push(valor);
        } else if (objeto instanceof PDFArray) for (const valor of objeto.asArray()) fila.push(valor);
    }
    return encontrados;
}

// A comparação precisa do MCID e da página correspondentes, não de procurar a
// mesma palavra em outro ponto do PDF. Referências de XObjects ficam bloqueadas.
function textosSubstitutos(pdf) {
    const paginas = new Map(pdf.getPages().map((p, i) => [p.ref.toString(), i + 1]));
    const fila = [{ valor: pdf.catalog.get(nome("StructTreeRoot")), pagina: null }], vistos = new Set(), textos = [];
    while (fila.length) {
        let { valor, pagina } = fila.pop();
        const objeto = pdf.context.lookup(valor);
        if (!objeto) continue;
        if (vistos.has(objeto)) {
            if (objeto instanceof PDFDict || objeto instanceof PDFArray) textos.push({ texto: "", conteudos: [], erro: "Árvore lógica cíclica/compartilhada não suportada." });
            continue;
        }
        vistos.add(objeto);
        if (objeto instanceof PDFArray) for (const valor of [...objeto.asArray()].reverse()) fila.push({ valor, pagina });
        else if (objeto instanceof PDFDict) {
            pagina = objeto.get(nome("Pg"))?.toString() || pagina;
            if (objeto.has(nome("ActualText"))) {
                const registro = { texto: objeto.lookup(nome("ActualText")).decodeText(), conteudos: [] };
                const pendentes = [{ valor: objeto.get(nome("K")), pagina }], visitados = new Set();
                while (pendentes.length) {
                    let { valor, pagina: refPagina } = pendentes.pop();
                    const item = pdf.context.lookup(valor);
                    if (item instanceof PDFArray) {
                        if (visitados.has(item)) { registro.erro = "Estrutura cíclica ou compartilhada não suportada."; break; }
                        visitados.add(item);
                        for (const valor of [...item.asArray()].reverse()) pendentes.push({ valor, pagina: refPagina });
                    } else if (item instanceof PDFNumber || item instanceof PDFDict) {
                        if (item instanceof PDFDict) {
                            if (visitados.has(item) || item.has(nome("Stm")) || item.has(nome("StmOwn"))) { registro.erro = "ActualText em estrutura compartilhada/XObject não suportado."; break; }
                            visitados.add(item); refPagina = item.get(nome("Pg"))?.toString() || refPagina;
                            const mcid = item.lookupMaybe(nome("MCID"), PDFNumber);
                            if (!mcid) { pendentes.push({ valor: item.get(nome("K")), pagina: refPagina }); continue; }
                            valor = mcid;
                        } else valor = item;
                        const numero = paginas.get(refPagina);
                        if (!refPagina) { registro.erro = "ActualText sem página identificável."; break; }
                        // Árvores antigas podem ainda mencionar uma página excluída pelo operador.
                        if (!numero) continue;
                        const mcid = valor.asNumber();
                        if (!Number.isInteger(mcid) || mcid < 0) { registro.erro = "MCID inválido na estrutura."; break; }
                        registro.conteudos.push({ pagina: numero, mcid });
                    } else { registro.erro = "ActualText sem conteúdo marcado identificável."; break; }
                }
                if (registro.conteudos.length || registro.erro) textos.push(registro);
            }
            fila.push({ valor: objeto.get(nome("K")), pagina });
        }
    }
    return textos;
}

// Provas efêmeras vinculadas à própria cópia e ao diagnóstico. Nenhuma opção da
// interface ou booleano persistido pode dispensar a comparação.
const provasActualText = new WeakMap();

export function diagnosticarImpressao(pdf, paginasAlvo) {
    const removiveis = new Map(), bloqueios = [], alvos = new Set(paginasAlvo);
    const registrar = (codigo, descricao) => removiveis.set(codigo, { codigo, descricao });
    const catalogo = pdf.catalog, nomes = catalogo.lookupMaybe(nome("Names"), PDFDict);
    const todosDicionarios = dicionarios(pdf), actualText = textosSubstitutos(pdf);
    if (catalogo.has(nome("StructTreeRoot")) || catalogo.has(nome("MarkInfo")) || todosDicionarios.some(d => d.has(nome("StructParent")) || d.has(nome("StructParents")))) {
        registrar("acessibilidade", "Árvore de acessibilidade, marcação lógica e referências StructParent/StructParents");
        if (actualText.length) bloqueios.push({ codigo: "actualtext", descricao: "ActualText na árvore exige comprovar redundância no conteúdo marcado, igualdade de texto e de renderização antes de permitir seu descarte." });
    }
    if (catalogo.has(nome("Outlines"))) registrar("marcadores", "Marcadores de navegação (Outlines)");
    if (catalogo.has(nome("Dests")) || nomes?.has(nome("Dests"))) registrar("destinos", "Destinos internos/nomeados");
    const abertura = catalogo.lookup(nome("OpenAction"));
    if (abertura instanceof PDFArray || abertura instanceof PDFDict && tipoAcao(abertura) === "/GoTo" && !abertura.has(nome("Next"))) {
        registrar("abertura", "Destino interno usado na abertura do documento");
    } else if (abertura && (!(abertura instanceof PDFDict) || !["/URI", "/GoToR"].includes(tipoAcao(abertura)) || abertura.has(nome("Next")))) {
        bloqueios.push({ codigo: "abertura-visual", descricao: "Ação de abertura não suportada; pode modificar a aparência ou o conteúdo." });
    }
    const formulario = catalogo.lookupMaybe(nome("AcroForm"), PDFDict);
    if (formulario?.lookupMaybe(nome("Fields"), PDFArray)?.size()) bloqueios.push({ codigo: "formulario", descricao: "Campos de formulário: suas aparências e cálculos não são tratados na preparação para impressão." });
    if (formulario?.has(nome("XFA"))) bloqueios.push({ codigo: "xfa", descricao: "Formulário XFA: conteúdo visual não tratado." });
    if (formulario?.has(nome("SigFlags"))) bloqueios.push({ codigo: "assinatura", descricao: "Assinatura digital: a preparação não preserva sua validade." });
    if (catalogo.has(nome("AA"))) bloqueios.push({ codigo: "acoes-documento", descricao: "Ações adicionais do documento podem alterar sua aparência na impressão." });
    if (nomes?.has(nome("JavaScript"))) bloqueios.push({ codigo: "scripts", descricao: "Scripts do documento podem alterar conteúdo visual; preparação não suportada." });
    const anotacoes = [];
    pdf.getPages().forEach((pagina, indice) => {
        const numero = indice + 1;
        if (pagina.node.has(nome("AA"))) bloqueios.push({ codigo: "acoes-pagina", descricao: `Página ${numero}: ações adicionais podem alterar conteúdo visível.` });
        for (const ref of pagina.node.Annots()?.asArray() || []) {
            const anotacao = pdf.context.lookup(ref, PDFDict);
            const tipo = anotacao.lookupMaybe(nome("Subtype"), PDFName)?.asString() || "sem Subtype";
            const acao = anotacao.lookupMaybe(nome("A"), PDFDict), interna = anotacao.has(nome("Dest")) || tipoAcao(acao) === "/GoTo";
            anotacoes.push({ pagina: numero, tipo, destinoInterno: interna });
            if (interna) {
                if (tipo !== "/Link" || acao?.has(nome("Next"))) bloqueios.push({ codigo: "destino-complexo", descricao: `Página ${numero}: destino interno em anotação/ação complexa não suportada.` });
                else registrar("links-internos", "Destinos/ações de links internos (os retângulos e bordas dos links são preservados)");
            }
            if (alvos.has(numero) && (!["/Link", "/Text", "/Popup"].includes(tipo) || anotacao.has(nome("AP")) || anotacao.has(nome("AA")) ||
                acao && (!["/URI", "/GoToR", "/GoTo"].includes(tipoAcao(acao)) || acao.has(nome("Next"))))) {
                bloqueios.push({ codigo: "aparencia-anotacao", descricao: `Página ${numero}: ${tipo}, aparência AP ou ação complexa pode conter conteúdo impresso não tratado. Nenhuma anotação será apagada para contornar esse risco.` });
            }
        }
    });
    return { removiveis: [...removiveis.values()], bloqueios, anotacoes, actualText,
        podePreparar: removiveis.size > 0 && bloqueios.length === 0 };
}

export class PDFCompatibilidadeErro extends Error {
    constructor(diagnostico) {
        const mensagem = diagnostico.bloqueios.length
            ? diagnostico.bloqueios.map(b => b.descricao).join(" ")
            : `${diagnostico.validacaoActualText ? `ActualText: ${diagnostico.actualText.length} substituições redundantes verificadas; texto extraível e renderização preservados antes de redimensionar. ` : ""}Foram identificados: ${diagnostico.removiveis.map(r => r.descricao).join("; ")}. Use Continuar em modo de impressão para revisar e confirmar o descarte desses recursos na cópia.`;
        super(mensagem);
        this.name = "PDFCompatibilidadeErro";
        this.diagnostico = diagnostico;
    }
}

// Apenas a cópia transacional. Streams, fontes, imagens, OCProperties e metadados
// de produção não são apagados. Objetos órfãos podem permanecer serializados.
export function prepararParaImpressao(pdf, diagnostico) {
    if (diagnostico.bloqueios.length) throw new PDFCompatibilidadeErro(diagnostico);
    if (diagnostico.actualText?.length && provasActualText.get(pdf) !== diagnostico) throw new Error("ActualText sem prova de compatibilidade para esta cópia.");
    return removerRecursos(pdf, diagnostico);
}

export async function validarActualTextImpressao(pdf, diagnostico, validar, { signal } = {}) {
    if (!diagnostico.actualText.length || !validar || diagnostico.bloqueios.some(b => b.codigo !== "actualtext")) return;
    const verificar = () => { if (signal?.aborted) throw new DOMException("Validação cancelada", "AbortError"); };
    verificar();
    const antes = await pdf.save(), copia = await PDFDocument.load(antes, { updateMetadata: false });
    // A exceção só existe para construir outra cópia descartável a ser comparada.
    // Nunca autoriza preparar/aplicar o documento recebido sem uma prova.
    removerRecursos(copia, diagnostico);
    verificar();
    let relatorio;
    try { relatorio = await validar({ antes, depois: await copia.save(), substitutos: diagnostico.actualText, signal }); }
    catch (erro) {
        if (erro.name === "AbortError") throw erro;
        throw new PDFCompatibilidadeErro({ ...diagnostico, bloqueios: [{ codigo: "actualtext", descricao: `Não foi possível validar ActualText: ${erro.message} Nenhuma alteração aplicada.` }] });
    }
    verificar();
    if (!relatorio?.compativel) throw new PDFCompatibilidadeErro({ ...diagnostico, bloqueios: [{ codigo: "actualtext", descricao: relatorio?.motivo || "Não foi comprovada a preservação de ActualText; operação bloqueada." }] });
    diagnostico.bloqueios = diagnostico.bloqueios.filter(b => b.codigo !== "actualtext");
    diagnostico.podePreparar = diagnostico.removiveis.length > 0;
    diagnostico.validacaoActualText = relatorio;
    provasActualText.set(pdf, diagnostico);
}

function removerRecursos(pdf, diagnostico) {
    const codigos = new Set(diagnostico.removiveis.map(r => r.codigo)), catalogo = pdf.catalog;
    if (codigos.has("acessibilidade")) {
        catalogo.delete(nome("StructTreeRoot")); catalogo.delete(nome("MarkInfo"));
        for (const dict of dicionarios(pdf)) { dict.delete(nome("StructParent")); dict.delete(nome("StructParents")); }
        for (const p of pdf.getPages()) if (p.node.lookupMaybe(nome("Tabs"), PDFName)?.asString() === "/S") p.node.delete(nome("Tabs"));
    }
    if (codigos.has("marcadores")) {
        catalogo.delete(nome("Outlines"));
        if (catalogo.lookupMaybe(nome("PageMode"), PDFName)?.asString() === "/UseOutlines") catalogo.delete(nome("PageMode"));
    }
    if (codigos.has("destinos")) {
        catalogo.delete(nome("Dests"));
        const nomes = catalogo.lookupMaybe(nome("Names"), PDFDict);
        if (nomes?.has(nome("Dests"))) {
            const copia = nomes.clone(pdf.context); copia.delete(nome("Dests"));
            if (copia.keys().length) catalogo.set(nome("Names"), copia); else catalogo.delete(nome("Names"));
        }
    }
    if (codigos.has("abertura")) catalogo.delete(nome("OpenAction"));
    if (codigos.has("links-internos")) for (const p of pdf.getPages()) for (const ref of p.node.Annots()?.asArray() || []) {
        const anotacao = pdf.context.lookup(ref, PDFDict);
        if (anotacao.lookupMaybe(nome("Subtype"), PDFName)?.asString() !== "/Link") continue;
        anotacao.delete(nome("Dest"));
        if (tipoAcao(anotacao.lookupMaybe(nome("A"), PDFDict)) === "/GoTo") anotacao.delete(nome("A"));
    }
    return diagnostico.removiveis.map(r => r.descricao);
}
