import { agruparDimensoes, intervalosPaginas } from "./pdf-dimension-groups.js";
// Agrupa metadados já conhecidos; não relê páginas nem serializa o documento.
export class PDFInformacoesDocumento {
    constructor(buscar) {
        this.buscar = buscar;
        this.textos = [];
        this.paginas = [];
    }

    atualizar(paginas, nome, assinaturas) {
        const buscar = this.buscar;
        const primeira = paginas[0];
        const agrupamento = agruparDimensoes(paginas);
        const variados = paginas.some(p => Math.abs(p.largura - primeira.largura) > .1 || Math.abs(p.altura - primeira.altura) > .1);
        const formatar = n => n.toLocaleString("pt-BR", { maximumFractionDigits: 2 });
        buscar("#pdf-nome").textContent = nome;
        buscar("#pdf-conteudo").hidden = false;
        buscar("#pdf-total").textContent = paginas.length;
        buscar("#pdf-tamanho").textContent = variados ? "Variados" : `${formatar(primeira.largura)} × ${formatar(primeira.altura)} mm`;
        const informacoes = buscar("#pdf-informacoes");
        const alertaAssinatura = !!assinaturas?.fontes.length && !assinaturas.preparada;
        const alertas = agrupamento.avisos.length + Number(alertaAssinatura);
        informacoes.classList.toggle("pdf-alerta", alertas > 0);
        informacoes.textContent = alertas ? `⚠ Informações · ${alertas} alerta${alertas > 1 ? "s" : ""}` : "Informações";
        buscar("#pdf-alertas").textContent = agrupamento.avisos.join(" ") || "Mesmo tamanho físico. Giro não altera o formato da página.";
        const orientacoes = buscar("#pdf-orientacoes-alerta");
        orientacoes.hidden = !agrupamento.orientacoesDiferentes;
        orientacoes.textContent = agrupamento.orientacoesDiferentes ? "Orientações diferentes detectadas" : "";
        buscar("#pdf-orientacoes-detalhes").textContent = agrupamento.detalhesOrientacoes;
        const aviso = buscar("#pdf-assinaturas-alerta");
        aviso.hidden = !assinaturas || !assinaturas.fontes.length;
        const fontes = assinaturas?.fontes || [];
        // Recolhe ao mudar a origem ou o estado, mantendo a escolha durante navegação/giro.
        if (this.preparada !== assinaturas?.preparada || fontes.length !== this.fontes?.length || fontes.some((o, i) => o !== this.fontes[i])) {
            buscar("#pdf-assinaturas-expandido").hidden = true;
            buscar("#pdf-assinaturas-expandir").setAttribute("aria-expanded", "false");
            buscar("#pdf-assinaturas-expandir").textContent = "Ver detalhes";
        }
        this.fontes = fontes;
        this.preparada = assinaturas?.preparada;
        aviso.classList.toggle("pdf-assinaturas-preparada", !!assinaturas?.preparada);
        if (assinaturas) {
            const bloqueada = !assinaturas.preparada && assinaturas.pendentes.some(o => ['detectada', 'inconclusiva'].includes(o.diagnostico.estado) || o.diagnostico.presentes > 0 || o.diagnostico.inconclusivas > 0);
            buscar("#pdf-assinaturas-titulo").textContent = assinaturas.preparada ? "Cópia para impressão preparada" : assinaturas.presentesPendentes ? "PDF assinado. Alterações bloqueadas." : bloqueada ? "Possível assinatura digital. Alterações bloqueadas por segurança." : "Campo de assinatura vazio";
            buscar("#pdf-assinaturas-resumo").textContent = assinaturas.preparada ? "Confira a cópia antes de imprimir." : bloqueada ? "Visualização disponível. Validade da assinatura não verificada." : "Campo vazio não comprova assinatura aplicada.";
            buscar("#pdf-assinaturas-resolver").textContent = bloqueada ? "Ver restrições" : assinaturas.preparada ? "Refazer preparo" : "Preparar cópia";
            buscar("#pdf-assinaturas-resolver").dataset.assinaturaBloqueada = String(bloqueada);
            const atual = !fontes.length ? "Documento atual: nenhuma assinatura estrutural identificada."
                : assinaturas.preparada
                ? "Documento atual: cópia para impressão preparada, verificada estruturalmente sem assinatura ativa. Confira a aparência no seu leitor/RIP."
                : `Documento atual: ${bloqueada ? 'alterações bloqueadas; código PDF_ASSINATURA_PROTEGIDA' : 'campo vazio; sem assinatura aplicada identificada'}. Páginas visuais atuais: ${assinaturas.paginasVisuais.join(", ") || "não determinadas"}; invisíveis nas origens pendentes: ${assinaturas.invisiveisPendentes}.`;
            const detalhes = atual + "\n\nHistórico e diagnóstico das origens\n" + assinaturas.validade + "\n" + (assinaturas.fontes.length ? assinaturas.fontes.map(o => {
                const paginasOriginais = [...new Set(o.diagnostico.assinaturas.flatMap(s => s.paginasVisuais))].sort((a, b) => a - b);
                const widgets = o.diagnostico.assinaturas.flatMap(s => s.widgets);
                return `${o.arquivo}: ${o.diagnostico.estado}; presentes ${o.diagnostico.presentes}; vazios ${o.diagnostico.vazios}; inconclusivos ${o.diagnostico.inconclusivas}. Páginas visuais na origem: ${paginasOriginais.join(", ") || "não determinadas"}; aparências AP reconhecidas: ${widgets.filter(w => w.visivel && w.aparencia).length}; invisíveis: ${o.diagnostico.assinaturas.filter(s => s.invisivel).length}. ${o.resolvida ? "Origem da cópia de impressão; não descreve uma assinatura válida na versão atual." : ""} ${o.diagnostico.indicios.join("; ")}`;
            }).join("\n") : "Nenhuma assinatura estrutural identificada nas origens analisadas. Isso não comprova ausência de assinaturas ou validade de certificados.");
            buscar("#pdf-assinaturas-detalhes").textContent = detalhes;
            buscar("#pdf-assinaturas-expandido").textContent = detalhes;
        }
        const lista = buscar("#pdf-lista-paginas");
        while (lista.children.length > agrupamento.grupos.length) lista.children[lista.children.length - 1].remove();
        agrupamento.grupos.forEach((g, indice) => {
            let item = lista.children[indice];
            if (!item) {
                item = document.createElement("li"); const detalhes = document.createElement("details");
                detalhes.append(document.createElement("summary"), document.createElement("p")); item.append(detalhes); lista.append(item);
            }
            const diferente = g.familia !== agrupamento.predominante;
            item.classList.toggle("pdf-dimensao-diferente", diferente);
            const titulo = `${g.nome} — ${formatar(g.largura)} × ${formatar(g.altura)} mm — ${g.paginas.length} página(s) · ${g.orientacao}${diferente ? " · diferente do predominante" : ""}`;
            const detalhes = item.children[0], chave = `${g.largura}:${g.altura}`;
            if (detalhes.dataset.dimensao !== chave) detalhes.open = false;
            detalhes.dataset.dimensao = chave;
            if (detalhes.children[0].textContent !== titulo) detalhes.children[0].textContent = titulo;
            detalhes.children[1].textContent = `Páginas: ${intervalosPaginas(g.paginas)}.${g.giradas.length ? ` Giro aplicado nas páginas: ${intervalosPaginas(g.giradas)}; isso não significa redimensionamento.` : ""}`;
        });
        this.paginas = paginas;
    }
}
