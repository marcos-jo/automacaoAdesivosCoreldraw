import { FORMATOS_PADRAO } from "./pdf-formats.js";
const igual = (a, b) => Math.abs(a - b) <= .1;
export function intervalosPaginas(numeros) {
    const partes = []; let inicio, fim;
    for (const n of numeros) {
        if (inicio === undefined) inicio = fim = n;
        else if (n === fim + 1) fim = n;
        else { partes.push(inicio === fim ? String(inicio) : `${inicio}–${fim}`); inicio = fim = n; }
    }
    if (inicio !== undefined) partes.push(inicio === fim ? String(inicio) : `${inicio}–${fim}`);
    return partes.join(", ");
}
export function agruparDimensoes(paginas) {
    const familias = [], grupos = [];
    const orientacoes = { retrato: [], paisagem: [], quadrada: [] };
    for (const p of paginas) {
        // largura/altura já são visuais (incluem 90°/270°); não reaplicar o giro.
        const orientacao = igual(p.largura, p.altura) ? "quadrada" : p.largura > p.altura ? "paisagem" : "retrato";
        orientacoes[orientacao].push(p.numero);
        const menor = Math.min(p.larguraBase ?? p.largura, p.alturaBase ?? p.altura), maior = Math.max(p.larguraBase ?? p.largura, p.alturaBase ?? p.altura);
        let familia = familias.find(f => igual(f.menor, menor) && igual(f.maior, maior));
        if (!familia) { familia = { menor, maior, total: 0 }; familias.push(familia); }
        familia.total++;
        let grupo = grupos.find(g => g.familia === familia && g.orientacao === orientacao && igual(g.largura, p.largura) && igual(g.altura, p.altura));
        if (!grupo) {
            const formato = FORMATOS_PADRAO.find(f => igual(f.largura, menor) && igual(f.altura, maior));
            grupo = { familia, nome: formato?.nome || "Personalizado", orientacao, largura:p.largura, altura:p.altura, paginas:[], giradas:[] }; grupos.push(grupo);
        }
        grupo.paginas.push(p.numero); if (p.rotacao % 360) grupo.giradas.push(p.numero);
    }
    const predominante = familias.reduce((a, b) => !a || b.total > a.total ? b : a, null);
    const variados = familias.length > 1;
    const orientacoesDiferentes = orientacoes.retrato.length > 0 && orientacoes.paisagem.length > 0;
    const avisos = [];
    if (variados) avisos.push("Atenção: há tamanhos diferentes do predominante.");
    if (orientacoesDiferentes) avisos.push("Orientações diferentes detectadas");
    const detalhesOrientacoes = Object.entries(orientacoes).filter(([, numeros]) => numeros.length).map(([nome, numeros]) =>
        `${nome === "quadrada" ? "Quadradas" : nome === "retrato" ? "Retrato" : "Paisagem"}: ${numeros.length} página(s) — páginas ${intervalosPaginas(numeros)}.`).join(" ");
    return { grupos, variados, predominante, orientacoes, orientacoesDiferentes, avisos, detalhesOrientacoes };
}
