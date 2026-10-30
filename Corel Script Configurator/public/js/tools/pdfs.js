import { renderPDFInterface } from "./pdf-interface.js";
import { PDFPreview } from "./pdf-preview.js";
import { PDFMiniaturas } from "./pdf-thumbnails.js";
import { PDFProtecao } from "./pdf-protection.js";
import { PDFDocumentoVisual } from "./pdf-visual-document.js";
import { PDFInformacoesDocumento } from "./pdf-document-info.js";
import { PDFRedimensionarUI } from "./pdf-resize-ui.js";
import { PDFRedimensionarSessao } from "./pdf-resize-session.js";
import { PDFImpressaoSessao } from "./pdf-print-session.js";
import { PDFImpressaoUI } from "./pdf-print-ui.js";
import { PDFControlesVisualizacao } from "./pdf-view-controls.js";
import { ordemMoverPaginas } from "./pdf-page-selection.js";
import { PDFAcoesPaginas } from "./pdf-page-actions.js";
import { PDFPaineis } from "./pdf-panel-layout.js";
import { JanelaExportacao } from "../app/export-dialog.js";
import { agruparDimensoes } from "./pdf-dimension-groups.js";
import { salvarArquivoPDF } from "../app/save-file.js";
import { informarDocumento, instalarImportacaoDesktop } from "../app/desktop-client.js";

import { salvarPaginasSelecionadas } from "./pdf-editor.js";

import { renderUnirPDFs } from "./pdf-merge.js";

import { PDFWorkspace } from "./pdf-workspace.js";

export function renderPDFs(app = document.getElementById("app"), { exportacao = null } = {}) {

    const raiz = renderPDFInterface(app, () => {
        if (!ocupado && documento && ferramentaAtiva()) {
            if (sessaoResize?.ativa) sessaoResize.reexibir().catch(erro => { status.textContent = erro.message; });
            else mostrarPagina(paginaDesejada);
        }
    });
    const removerImportacaoDesktop = instalarImportacaoDesktop(raiz);

    // Cada instância guarda seus elementos, inclusive durante tarefas assíncronas.
    const elementos = new Map(
        [...app.querySelectorAll("[id]")].map(elemento => ["#" + elemento.id, elemento])
    );
    const buscar = seletor => elementos.get(seletor);

    const arquivoInput = buscar("#pdf-arquivo");
    const status = buscar("#pdf-status");
    const resumo = new PDFInformacoesDocumento(buscar);
    const visualizacao = buscar("#pdf-visualizacao");
    const informacoes = buscar("#pdf-informacoes");
    const detalhes = buscar("#pdf-detalhes");
    const janelaExportacao = exportacao || new JanelaExportacao(raiz);

    //---------------------------------
    // Controle dos botões de páginas
    //---------------------------------

    // A raiz é própria desta instância; #app permanece na página.
    raiz.addEventListener("click", (evento) => {

        const botao = evento.target.closest("button");

        if (!botao) return;

        if (botao.id?.startsWith("pdf-ferramenta-") && botao.id !== "pdf-ferramenta-redimensionar" && sessaoResize?.ativa) {
            interfaceResize.invalidar();
        }

        if (
            botao.id !== "pdf-extrair" &&
            botao.id !== "pdf-excluir"
        ) {
            return;
        }

        evento.preventDefault();

        console.log("Botão PDF clicado:", botao.id);

        if (!documento) {

            status.textContent =
                "Abra um PDF antes de utilizar esta ferramenta.";

            return;

        }

        const operacao = botao.id === "pdf-extrair"
            ? "extrair"
            : "excluir";

        processarPaginas(operacao);

    });

    let documento = null;
    let paginaAtual = 1;
    let paginaDesejada = 1;
    let rotacoes = new Map();
    let encerrada = false;
    let suspensa = false;

    //---------------------------------
    // Documento de trabalho
    //---------------------------------

    let protecao;
    const workspace = new PDFWorkspace(() => protecao?.atualizar());
    protecao = new PDFProtecao(raiz, workspace);
    let ocupado = false;
    let fimOperacao = Promise.resolve();
    let concluirOperacao;
    let sessaoResize, interfaceResize, sessaoImpressao, interfaceImpressao;
    const previa = new PDFPreview(visualizacao);
    const controlesVisualizacao = new PDFControlesVisualizacao(buscar, previa,
        () => sessaoResize?.ativa ? sessaoResize.reexibir() : mostrarPagina(paginaDesejada),
        () => !ocupado && !!documento && !sessaoImpressao?.ativa && ferramentaAtiva());
    const paineis = new PDFPaineis(buscar);
    const modalPrevia = new PDFPreview(buscar("#pdf-impressao-previa"));
    const miniaturas = new PDFMiniaturas(buscar("#pdf-miniaturas"), numero => {
        if (!ocupado && ferramentaAtiva()) mostrarPagina(numero);
    }, (origem, destino, grupo) => moverPagina(origem, destino, null, grupo));
    const acoesPaginas = new PDFAcoesPaginas(raiz, miniaturas, {
        duplicar: duplicarSelecionadas,
        excluir: () => processarPaginas("excluir", miniaturas.selecao.numeros())
    });

    const uniao = renderUnirPDFs(app, adicionarAoDocumento);
    sessaoResize = new PDFRedimensionarSessao({
        workspace, previa, ativa: ferramentaAtiva,
        documento: () => documento, mostrarOriginal: () => mostrarPagina(paginaAtual, { propagarErro: true }),
        rotulo: texto => { buscar("#pdf-pagina-atual").textContent = texto; },
        iniciar: () => iniciarOperacao(true), liberar: () => definirOcupado(false),
        atualizar: () => definirOcupado(ocupado),
        aplicar: async (preparacao, candidato) => {
            const anterior = workspace.capturarEstado();
            try {
                if (!workspace.aplicarRedimensionamento(preparacao)) return;
                await atualizarVisualizacaoWorkspace({ preservarPaginas: true, candidatoPreparado: candidato,
                    invalidarIds: new Set(preparacao.alteradas) });
                status.textContent = "Páginas redimensionadas. Alterações ainda não exportadas.";
            } catch (erro) { workspace.restaurarEstado(anterior); throw erro; }
        }
    });
    interfaceResize = new PDFRedimensionarUI(buscar, {
        contexto: () => ({ total: documento?.numPages || 0, atual: paginaAtual, revisao: workspace.revisao, dimensoes: documento?.obterDimensoes?.() || [] }),
        interpretar: interpretarPaginas,
        prever: config => sessaoResize.prever(config),
        cancelar: () => sessaoResize.cancelar(), aplicar: () => sessaoResize.aplicar()
    });
    sessaoImpressao = new PDFImpressaoSessao({
        workspace, previa, modalPrevia, ativa: ferramentaAtiva,
        podeAbrir: () => !ocupado && !!documento && !sessaoResize.ativa && !sessaoImpressao.ativa,
        iniciar: () => iniciarOperacao(true), liberar: () => definirOcupado(false), atualizar: () => definirOcupado(ocupado),
        progresso: texto => interfaceImpressao?.mensagem(texto),
        aplicar: async (preparacao, candidato) => {
            const anterior = workspace.capturarEstado();
            try {
                workspace.aplicarCopiaImpressao(preparacao);
                if (!await atualizarVisualizacaoWorkspace({ candidatoPreparado: candidato, reiniciarSelecao: true })) throw new Error("Não foi possível adotar a visualização da cópia.");
                status.textContent = "Cópia proposta para impressão aplicada. Confira no leitor/RIP e use Salvar cópia (_IMPRESSAO). Validade da assinatura não preservada.";
            } catch (erro) {
                workspace.restaurarEstado(anterior);
                try { await atualizarVisualizacaoWorkspace(); }
                catch { previa.limparCanvas(); }
                throw erro;
            }
        }
    });
    interfaceImpressao = new PDFImpressaoUI(buscar, sessaoImpressao, interpretarPaginas);

    function ferramentaAtiva() {
        return !encerrada && !suspensa && app.querySelector("#pdf-ferramenta") === raiz;
    }

    function definirOcupado(valor) {
        ocupado = valor;
        const bloqueado = valor || !!sessaoResize?.ativa || !!sessaoImpressao?.ativa;
        arquivoInput.disabled = bloqueado;
        informacoes.disabled = bloqueado || !documento;

        for (const id of ["pdf-girar-esquerda", "pdf-girar-direita", "pdf-extrair", "pdf-excluir"]) {
            buscar("#" + id).disabled = bloqueado || !documento;
        }

        buscar("#pdf-anterior").disabled = bloqueado || !documento || paginaAtual <= 1;
        buscar("#pdf-proxima").disabled = bloqueado || !documento || paginaAtual >= documento.numPages;
        uniao.definirOcupado(bloqueado);
        miniaturas.definirOcupado(bloqueado);
        buscar("#pdf-paginas-acoes").disabled = bloqueado || !documento;
        if (bloqueado) acoesPaginas.fechar(false);
        interfaceResize?.atualizarEstado({ ocupado: valor || !!sessaoImpressao?.ativa, ativa: !!sessaoResize?.ativa, aplicando: !!sessaoResize?.aplicando });
        buscar("#pdf-assinaturas-resolver").disabled = bloqueado || !documento;
        interfaceImpressao?.atualizar();
        atualizarControlesOrdem();
        atualizarBotaoSalvar();
        informarDocumento({ pending: workspace.alterado, busy: valor });
        if (!valor) {
            const concluir = concluirOperacao;
            concluirOperacao = null;
            concluir?.();
        }
    }

    function iniciarOperacao(permitirPrevia = false) {
        if (ocupado || !permitirPrevia && (sessaoResize?.ativa || sessaoImpressao?.ativa) || protecao.confirmacao || !ferramentaAtiva()) return false;
        fimOperacao = new Promise(resolve => { concluirOperacao = resolve; });
        // Uma navegação pendente não deve trocar a página durante a edição.
        previa.invalidar();
        definirOcupado(true);
        return true;
    }

    function atualizarControlesOrdem() {
        const bloqueado = ocupado || !!sessaoResize?.ativa || !!sessaoImpressao?.ativa;
        const selecionadas = miniaturas.selecao.numeros(), inicio = selecionadas.length > 1 ? selecionadas[0] : paginaAtual;
        const maximo = Math.max(1, (documento?.numPages || 1) - Math.max(0, selecionadas.length - 1));
        buscar("#pdf-mover-cima").disabled = bloqueado || !documento || inicio <= 1;
        buscar("#pdf-mover-baixo").disabled = bloqueado || !documento || inicio >= maximo;
        buscar("#pdf-mover-destino").disabled = bloqueado || !documento || documento.numPages < 2;
        buscar("#pdf-mover-aplicar").disabled = bloqueado || !documento || documento.numPages < 2;
        buscar("#pdf-mover-destino").max = maximo;
    }

    function moverParaPosicao() {
        if (ocupado || !documento || !ferramentaAtiva()) return;
        const entrada = buscar("#pdf-mover-destino");
        const texto = entrada.value.trim();
        const destino = Number(texto);
        if (!texto || !Number.isSafeInteger(destino) || destino < 1 || destino > documento.numPages) {
            status.textContent = `Informe uma posição inteira entre 1 e ${documento.numPages}.`;
            entrada.setAttribute("aria-invalid", "true");
            entrada.focus();
            return;
        }
        entrada.setAttribute("aria-invalid", "false");
        moverPagina(paginaAtual, destino, buscar("#pdf-mover-aplicar"));
    }

    buscar("#pdf-mover-aplicar").addEventListener("click", moverParaPosicao);
    buscar("#pdf-mover-destino").addEventListener("keydown", evento => {
        if (evento.key === "Enter") { evento.preventDefault(); moverParaPosicao(); }
    });

    async function moverPagina(origem, destino, controle = null, grupo = null) {
        if (!documento || !iniciarOperacao()) return;
        const anterior = workspace.capturarEstado();
        const paginaAnterior = paginaAtual;
        try {
            const escolhidas = grupo?.paginas || miniaturas.selecao.numeros();
            const numeros = escolhidas.length > 1 || escolhidas.includes(origem) ? escolhidas : [origem];
            const ordem = ordemMoverPaginas(workspace.getTotalPaginas(), numeros, destino, grupo?.divisoria);
            const resultado = { alterada: await workspace.reordenarPaginas(ordem), ordem };
            if (!resultado.alterada) return;
            // A permutação conserva a identidade da página exibida, inclusive se
            // outra página tiver sido arrastada para antes/depois dela.
            paginaAtual = resultado.ordem.indexOf(paginaAnterior) + 1;
            if (await atualizarVisualizacaoWorkspace({ reutilizar: true })) {
                buscar("#pdf-selecao-paginas").value = "";
                status.textContent = numeros.length > 1 ? `${numeros.length} páginas movidas em grupo. Alterações ainda não exportadas.` : `Página ${origem} movida para ${destino}. Alterações ainda não exportadas.`;
            }
        } catch (erro) {
            workspace.restaurarEstado(anterior);
            paginaAtual = paginaAnterior;
            console.error(erro);
            status.textContent = "Não foi possível mover a página: " + erro.message;
        } finally {
            definirOcupado(false);
            if (controle && ferramentaAtiva()) {
                if (!controle.disabled) controle.focus();
                else miniaturas.focarSelecionada();
            }
        }
    }

    buscar("#pdf-mover-cima").addEventListener("click", () => {
        const selecionadas = miniaturas.selecao.numeros();
        moverPagina(paginaAtual, (selecionadas.length > 1 ? selecionadas[0] : paginaAtual) - 1, buscar("#pdf-mover-cima"));
    });
    buscar("#pdf-mover-baixo").addEventListener("click", () => {
        const selecionadas = miniaturas.selecao.numeros();
        moverPagina(paginaAtual, (selecionadas.length > 1 ? selecionadas[0] : paginaAtual) + 1, buscar("#pdf-mover-baixo"));
    });

    async function adicionarAoDocumento(arquivos, nomeArquivo) {
        if (!iniciarOperacao()) {
            throw new Error("Aguarde a operação atual terminar.");
        }

        const anterior = workspace.capturarEstado();

        status.textContent = "Adicionando PDFs ao documento...";

        try {
            const adicionadas = await workspace.adicionarArquivos(arquivos, nomeArquivo);
            if (await atualizarVisualizacaoWorkspace()) {
                status.textContent = `${adicionadas} página(s) adicionada(s). Alterações ainda não exportadas.`;
            }
            return adicionadas;
        } catch (erro) {
            workspace.restaurarEstado(anterior);
            status.textContent = "Não foi possível adicionar os PDFs: " + erro.message;
            throw erro;
        } finally {
            definirOcupado(false);
        }
    }

    arquivoInput.addEventListener("change", async () => {

        const arquivo = arquivoInput.files?.[0];
        arquivoInput.value = "";

        if (!arquivo || !iniciarOperacao()) return;

        const anterior = workspace.capturarEstado();
        const paginaAnterior = paginaAtual;

        try {
            if (!await protecao.confirmarDescarte() || !ferramentaAtiva()) return;
            status.textContent = "Analisando PDF...";
            await workspace.abrir(arquivo);
            paginaAtual = 1;

            if (await atualizarVisualizacaoWorkspace({ reiniciarSelecao: true })) {
                detalhes.hidden = true;
                informacoes.setAttribute("aria-expanded", "false");
                status.textContent = "PDF carregado.";
            }
        } catch (erro) {
            workspace.restaurarEstado(anterior);
            paginaAtual = paginaAnterior;
            console.error(erro);
            status.textContent = "Não foi possível abrir o PDF. Verifique o arquivo ou o console.";
        } finally {
            definirOcupado(false);
        }
    });

    async function mostrarPagina(numero, { propagarErro = false } = {}) {
        if (!documento || !ferramentaAtiva()) return;
        paginaDesejada = numero;
        const pdf = documento;
        const versaoRolagem = miniaturas.versaoRolagem;
        try {
            const exibida = await previa.mostrar(numero, rotacoes.get(numero) || 0);
            if (!exibida || !ferramentaAtiva() || pdf !== documento) return;
            controlesVisualizacao.paginaExibida(pdf.obterChaveMiniatura?.(numero) ?? numero, pdf.fonte || pdf);
            paginaAtual = numero;
            atualizarControlesOrdem();
            buscar("#pdf-pagina-atual").textContent = numero + " / " + pdf.numPages;
            buscar("#pdf-anterior").disabled = ocupado || !!sessaoImpressao?.ativa || numero === 1;
            buscar("#pdf-proxima").disabled = ocupado || !!sessaoImpressao?.ativa || numero === pdf.numPages;
            miniaturas.selecionar(numero, { versaoRolagem });
        } catch (erro) {
            if (!ferramentaAtiva()) return;
            if (propagarErro) throw erro;
            console.error(erro);
            status.textContent = "Erro ao visualizar a página.";
        }
    }

    //---------------------------------
    // Atualizar PDF após alterações
    //---------------------------------

    async function atualizarVisualizacaoWorkspace({ reutilizar = false, preservarPaginas = reutilizar, candidatoPreparado = null, invalidarIds = new Set(), reiniciarSelecao = false } = {}) {

        const anterior = documento;
        if (reutilizar && anterior instanceof PDFDocumentoVisual) {
            try {
                const novo = anterior.atualizar(workspace);
                if (!await previa.atualizarDocumento(novo) || !ferramentaAtiva()) return false;
                documento = novo;
                rotacoes.clear();
                await mostrarPagina(paginaAtual, { propagarErro: true });
                await miniaturas.atualizarDocumento(novo, paginaAtual);
                resumo.atualizar(novo.obterDimensoes(), workspace.nomeArquivo, workspace.obterResumoAssinaturas());
                atualizarBotaoSalvar();
                return true;
            } catch {
                // Fonte incompatível ou falha de render: mantém o caminho completo
                // como alternativa, sem publicar índices antigos sobre bytes novos.
                documento = anterior;
                await previa.atualizarDocumento(anterior);
            }
        }

        if (!ferramentaAtiva()) return false;
        const candidato = candidatoPreparado || await previa.preparar(await workspace.obterBytes());
        if (!candidato) return false;
        let adotada = false;
        try {
            const novoDocumento = await PDFDocumentoVisual.criar(candidato.documento, workspace,
                preservarPaginas ? anterior : null, invalidarIds);
            candidato.documento = novoDocumento;

            if (!ferramentaAtiva()) return false;

            const selecionada = Math.min(paginaAtual, novoDocumento.numPages);
            // Só destrói a fonte antiga depois de validar a renderização nova.
            if (!await previa.previsualizar(novoDocumento, selecionada) || !ferramentaAtiva()) return false;
            await miniaturas.limpar({ preservarCache: preservarPaginas });
            miniaturas.invalidarIdentidades(invalidarIds);
            if (!ferramentaAtiva() || !await previa.adotar(candidato, { preservarCanvas: true })) return false;
            documento = novoDocumento;
            adotada = true;
            if (reiniciarSelecao) miniaturas.selecao.limpar();

            // As rotações já estão no workspace; não aplicá-las novamente.
            rotacoes.clear();

            resumo.atualizar(novoDocumento.obterDimensoes(), workspace.nomeArquivo, workspace.obterResumoAssinaturas());

            //---------------------------------
            // Atualiza a página selecionada
            //---------------------------------

            paginaAtual = Math.min(
                paginaAtual,
                documento.numPages
            );

            raiz.classList.toggle("pdf-com-documento", true);
            buscar("#pdf-painel-miniaturas").hidden = false;
            if (preservarPaginas) await miniaturas.atualizarDocumento(documento, paginaAtual);
            else miniaturas.definirDocumento(documento, paginaAtual);
            await mostrarPagina(paginaAtual, { propagarErro: true });

            atualizarBotaoSalvar();

            return true;

        } finally {
            if (!adotada) await previa.descartar(candidato);
        }

    }

    buscar("#pdf-anterior").addEventListener("click", () => {
        if (!ocupado && documento && paginaAtual > 1) {
            miniaturas.selecao.clicar(paginaAtual - 1);
            mostrarPagina(paginaAtual - 1);
        }
    });

    buscar("#pdf-proxima").addEventListener("click", () => {
        if (!ocupado && documento && paginaAtual < documento.numPages) {
            miniaturas.selecao.clicar(paginaAtual + 1);
            mostrarPagina(paginaAtual + 1);
        }
    });

    informacoes.addEventListener("click", () => {
        detalhes.hidden = !detalhes.hidden;
        informacoes.setAttribute("aria-expanded", String(!detalhes.hidden));
    });

    //---------------------------------
    // Atualiza botão Salvar cópia
    //---------------------------------

    function atualizarBotaoSalvar() {

        buscar("#pdf-salvar-copia").disabled =
            ocupado || !!sessaoResize?.ativa || !!sessaoImpressao?.ativa || !workspace.pdf;

    }


    //---------------------------------
    // Girar página no workspace
    //---------------------------------

    async function girarPagina(graus) {

        if (!documento || !iniciarOperacao()) return;
        const anterior = workspace.capturarEstado();
        let aplicada = false;

        try {

            //---------------------------------
            // Aplica no documento de trabalho
            //---------------------------------

            workspace.girarPagina(
                paginaAtual,
                graus
            );
            aplicada = true;

            status.textContent =
                "Atualizando rotação...";

            //---------------------------------
            // Atualiza a pré-visualização
            //---------------------------------

            await atualizarVisualizacaoWorkspace({ reutilizar: true });

            status.textContent =
                "Página girada. Alterações ainda não exportadas.";

        } catch (erro) {

            if (aplicada) {
                workspace.girarPagina(paginaAtual, -graus);
                workspace.restaurarEstado(anterior);
            }
            console.error(erro);

            status.textContent =
                "Erro ao girar página: " + erro.message;

        } finally {

            definirOcupado(false);

        }

    }

    //---------------------------------
    // Botões de rotação
    //---------------------------------

    buscar("#pdf-girar-esquerda")
        .addEventListener("click", () => {

            girarPagina(-90);

        });


    buscar("#pdf-girar-direita")
        .addEventListener("click", () => {

            girarPagina(90);

        });


    //---------------------------------
    // Exportar PDF modificado
    //---------------------------------

    buscar("#pdf-salvar-copia").addEventListener("click", () => {
        if (!documento || buscar("#pdf-salvar-copia").disabled || !ferramentaAtiva()) return;
        const assinaturas = workspace.obterResumoAssinaturas(), preparo = workspace.preparoImpressao;
        const fmt = n => n.toLocaleString("pt-BR", { maximumFractionDigits:2 });
        const diagnostico = agruparDimensoes(documento.obterDimensoes()), grupos = diagnostico.grupos;
        const dimensoes = grupos.slice(0,4).map(g => `${fmt(g.largura)} × ${fmt(g.altura)} mm (${g.paginas.length} páginas)`).join("; ") + (grupos.length > 4 ? `; e mais ${grupos.length - 4} grupos. Consulte Informações.` : "");
        const avisos = [...diagnostico.avisos];
        if (diagnostico.orientacoesDiferentes) avisos.push(diagnostico.detalhesOrientacoes);
        avisos.push("Cores, sangrias e resolução efetiva das imagens não verificadas.");
        if (preparo) avisos.push("Nas páginas rasterizadas do preparo, texto e vetores tornaram-se imagem RGB sobre branco, sem texto pesquisável; cores/transparências podem diferir no RIP. A cópia não preserva a validade da assinatura. Confira com o original em leitor confiável e no RIP.");
        else if (assinaturas.fontes.length) avisos.push("As origens possuem campos/estruturas de assinatura. Reescrever ou editar o PDF pode invalidar assinaturas; validade criptográfica não verificada.");
        janelaExportacao.abrir({ nome:workspace.obterNomeExportacao(), impressao:assinaturas.preparada,
            modalidade: assinaturas.preparada ? "Cópia preparada para impressão" : "Documento de trabalho",
            resumo:`${documento.numPages} página(s)`, dimensoes,
            estado:workspace.alterado ? "Alterações ainda não exportadas" : "Sem alterações pendentes", avisos:avisos.join("\n")
        }, exportarDocumento, raiz);
    });

    async function exportarDocumento(nomeExportado, { signal }) {
        if (!workspace.pdf || !iniciarOperacao()) throw new Error("Conclua a operação atual antes de exportar.");

        const pdfAtual = documento;
        const versaoExportada = workspace.capturarVersao();

        status.textContent =
            "Gerando cópia PDF...";

        try {

            //---------------------------------
            // Obtém o PDF atualizado
            //---------------------------------

            const bytes = await workspace.obterBytes();
            if (signal.aborted) throw new DOMException("Exportação cancelada", "AbortError");

            if (!ferramentaAtiva()) return;

            //---------------------------------
            // Prepara o download
            //---------------------------------

            await salvarArquivoPDF(bytes, nomeExportado, { signal });
            workspace.marcarExportado(versaoExportada);

            if (pdfAtual === documento) {

                status.textContent =
                    "PDF exportado com sucesso.";

            }

        } catch (erro) {

            if (pdfAtual === documento) {

                status.textContent =
                    erro.name === "AbortError" ? "Exportação cancelada. Documento preservado." : "Não foi possível exportar o PDF: " + erro.message;

            }
            throw erro;

        } finally {

            definirOcupado(false);

        }

    }

    //---------------------------------
    // Interpretar seleção de páginas
    //---------------------------------

    function interpretarPaginas(texto, total) {

        if (!texto.trim()) {
            throw new Error(
                "Informe as páginas desejadas."
            );
        }

        const selecionadas = new Set();

        const partes = texto.split(",");

        for (const parte of partes) {

            const trecho = parte.trim();

            const correspondencia = trecho.match(
                /^(\d+)(?:\s*-\s*(\d+))?$/
            );

            if (!correspondencia) {
                throw new Error(
                    `Seleção inválida: ${trecho}`
                );
            }

            const inicio = Number(correspondencia[1]);

            const fim = correspondencia[2]
                ? Number(correspondencia[2])
                : inicio;

            if (
                !Number.isSafeInteger(inicio) ||
                !Number.isSafeInteger(fim) ||
                inicio < 1 ||
                fim < inicio ||
                fim > total
            ) {
                throw new Error(
                    `Intervalo inválido: ${trecho}. ` +
                    `Este PDF possui ${total} páginas.`
                );
            }

            for (let numero = inicio; numero <= fim; numero++) {
                selecionadas.add(numero);
            }
        }

        return [...selecionadas].sort((a, b) => a - b);

    }


    //---------------------------------
    // Processar extração e exclusão
    //---------------------------------

    async function duplicarSelecionadas() {
        if (!documento || !iniciarOperacao()) return;
        const anterior = workspace.capturarEstado(), selecionadas = new Set(miniaturas.selecao.ids), paginaAnterior = paginaAtual;
        try {
            const novas = await workspace.duplicarPaginas(miniaturas.selecao.numeros());
            paginaAtual = novas[0];
            miniaturas.selecao.ids = new Set(novas.map(n => workspace.obterPaginas()[n - 1].id));
            if (!await atualizarVisualizacaoWorkspace({ preservarPaginas: true })) throw new Error("Não foi possível atualizar a prévia.");
            status.textContent = `${novas.length} página(s) duplicada(s). Alterações ainda não exportadas.`;
        } catch (erro) {
            workspace.restaurarEstado(anterior); paginaAtual = paginaAnterior; miniaturas.selecao.ids = selecionadas;
            try {
                documento = documento.atualizar(workspace);
                await previa.atualizarDocumento(documento);
                await mostrarPagina(paginaAtual, { propagarErro: true });
                await miniaturas.atualizarDocumento(documento, paginaAtual);
            } catch { previa.limparCanvas(); }
            status.textContent = erro.message;
        } finally { definirOcupado(false); if (ferramentaAtiva()) miniaturas.focarSelecionada(); }
    }

    async function processarPaginas(operacao, selecao = null) {

        if (!documento || ocupado || !ferramentaAtiva()) {
            return;
        }

        let paginas;

        try {

            paginas = selecao || interpretarPaginas(
                buscar("#pdf-selecao-paginas").value,
                documento.numPages
            );

            if (
                operacao === "excluir" &&
                paginas.length === documento.numPages
            ) {
                throw new Error(
                    "Não é possível excluir todas as páginas."
                );
            }
            if (selecao && (!paginas.length || !globalThis.confirm(`Excluir ${paginas.length} página(s) selecionada(s) do documento de trabalho?`))) return;

        } catch (erro) {

            status.textContent = erro.message;
            return;

        }

        //---------------------------------
        // Preparar exportação
        //---------------------------------

        const pdfAtual = documento;

        const nomeArquivo =
            buscar("#pdf-nome").textContent;

        const rotacoesAtuais = new Map(rotacoes);

        if (!iniciarOperacao()) return;

        status.textContent = "Gerando novo PDF...";

        try {

            //---------------------------------
            // EXCLUIR: altera o workspace
            //---------------------------------

            if (operacao === "excluir") {
                const identidadeAtual = selecao ? workspace.obterPaginas()[paginaAtual - 1]?.id : null;
                workspace.excluirPaginas(paginas);
                if (!selecao) miniaturas.selecao.limpar();
                if (selecao) {
                    const novaPosicao = workspace.obterPaginas().findIndex(p => p.id === identidadeAtual);
                    paginaAtual = novaPosicao < 0 ? Math.min(paginaAtual, workspace.getTotalPaginas()) : novaPosicao + 1;
                }

                status.textContent =
                    "Atualizando visualização...";

                await atualizarVisualizacaoWorkspace({ preservarPaginas: true });

                buscar("#pdf-selecao-paginas").value = "";

                status.textContent =
                    `${paginas.length} página(s) excluída(s). ` +
                    "Alterações ainda não exportadas.";

            }

            //---------------------------------
            // EXTRAIR: mantém comportamento atual
            //---------------------------------

            else {

                await salvarPaginasSelecionadas(
                    pdfAtual,
                    paginas,
                    operacao,
                    rotacoesAtuais,
                    nomeArquivo
                );

                if (pdfAtual === documento) {

                    status.textContent =
                        `${paginas.length} página(s) extraída(s) com sucesso.`;

                }

            }

        } catch (erro) {

            console.error(erro);

            if (operacao === "excluir" && documento instanceof PDFDocumentoVisual) {
                // A exclusão já foi aplicada. Se o novo PDF.js falhar, manter os
                // índices sincronizados e ocultar o canvas antigo, sem desfazer bytes.
                documento = documento.atualizar(workspace);
                paginaAtual = Math.min(paginaAtual, documento.numPages);
                await previa.atualizarDocumento(documento);
                previa.limparCanvas();
                await miniaturas.atualizarDocumento(documento, paginaAtual);
                resumo.atualizar(documento.obterDimensoes(), workspace.nomeArquivo, workspace.obterResumoAssinaturas());
                buscar("#pdf-pagina-atual").textContent = `${paginaAtual} / ${documento.numPages}`;
            }
            status.textContent =
                "Erro ao gerar PDF: " + erro.message;

        } finally {

            definirOcupado(false);
            if (selecao && ferramentaAtiva()) miniaturas.focarSelecionada();

        }

    }

    // O gerenciador de abas aguarda este encerramento apenas ao fechar o módulo.
    let encerramento;
    function destruir() {
        if (!encerramento) {
            encerrada = true;
            removerImportacaoDesktop();
            informarDocumento({ pending: false, busy: false });
            controlesVisualizacao.destruir();
            acoesPaginas.destruir();
            paineis.destruir();
            interfaceResize.destruir();
            interfaceImpressao.destruir();
            protecao.destruir();
            janelaExportacao.fecharDono(raiz);
            if (!exportacao) janelaExportacao.destruir();
            uniao.destruir();
            previa.invalidar();
            encerramento = (async () => {
                await sessaoResize.cancelar();
                await sessaoImpressao.cancelar();
                await modalPrevia.destruir();
                await miniaturas.destruir();
                await previa.destruir();
                documento = null;
            })();
        }
        return encerramento;
    }

    destruir.confirmarSaida = async () => {
        // Uma união/importação em andamento pode criar pendências ao terminar.
        while (ocupado) await fimOperacao;
        if (!ferramentaAtiva()) return false;
        return protecao.confirmarDescarte();
    };
    destruir.podeAlternar = () => ![...raiz.querySelectorAll("dialog")].some(dialogo => dialogo.open) && !(janelaExportacao.aberta && janelaExportacao.dono === raiz);
    destruir.menus = () => {
        const comando = (id, titulo, alvo, organizar = false) => {
            const elemento = () => buscar("#" + alvo) || app.querySelector("#" + alvo);
            const exigeDocumento = !["pdf-arquivo", "pdf-uniao-abrir", "pdf-painel-alternar"].includes(alvo);
            return { id, titulo,
                podeExecutar: () => ferramentaAtiva() && (!exigeDocumento || !!documento) && !!elemento() && !elemento().disabled && (!organizar || !buscar("#pdf-paginas-acoes").disabled),
                executar: () => elemento().click() };
        };
        return [
            { id:"arquivo", titulo:"Arquivo", itens:[
                comando("pdf-importar","Abrir / Importar PDF","pdf-arquivo"),
                comando("pdf-adicionar","Adicionar PDFs ao documento","pdf-uniao-abrir"),
                comando("pdf-exportar","Exportar / Salvar cópia","pdf-salvar-copia"),
                comando("pdf-informacoes","Informações do documento","pdf-informacoes")
            ] },
            { id:"editar", titulo:"Editar", itens:[
                comando("pdf-organizar","Organizar páginas","pdf-ferramenta-organizar"),
                comando("pdf-duplicar","Duplicar páginas selecionadas","pdf-duplicar",true),
                comando("pdf-excluir","Excluir páginas selecionadas","pdf-excluir-selecao",true),
                comando("pdf-girar-esquerda","Girar esquerda","pdf-girar-esquerda"),
                comando("pdf-girar-direita","Girar direita","pdf-girar-direita"),
                comando("pdf-mover-cima","Mover seleção para cima","pdf-mover-cima"),
                comando("pdf-mover-baixo","Mover seleção para baixo","pdf-mover-baixo")
            ] },
            { id:"exibir", titulo:"Exibir", itens:[
                comando("pdf-ajustar-pagina","Ajustar página","pdf-zoom-pagina"),
                comando("pdf-ajustar-largura","Ajustar largura","pdf-zoom-largura"),
                comando("pdf-zoom-real","Zoom 100%","pdf-zoom-real"),
                comando("pdf-painel",() => buscar("#pdf-painel-ferramentas").hidden ? "Mostrar ferramentas" : "Ocultar ferramentas","pdf-painel-alternar")
            ] }
        ];
    };
    let posicaoAba;
    destruir.suspender = async () => {
        while (ocupado) await fimOperacao;
        posicaoAba = { x: visualizacao.scrollLeft, y: visualizacao.scrollTop, miniaturas: miniaturas.container.scrollTop };
        suspensa = true; acoesPaginas.fechar(false); controlesVisualizacao.pausar(); paineis.suspensa = true;
        await previa.cancelarRender(); await miniaturas.pausar();
    };
    destruir.retomar = async () => {
        if (encerrada) return;
        suspensa = false; paineis.suspensa = false; paineis.aplicar();
        if (documento) {
            if (sessaoResize.ativa) await sessaoResize.reexibir();
            else await mostrarPagina(paginaDesejada, { propagarErro: true });
        }
        if (posicaoAba) {
            visualizacao.scrollLeft = posicaoAba.x; visualizacao.scrollTop = posicaoAba.y;
            miniaturas.container.scrollTop = posicaoAba.miniaturas;
        }
        miniaturas.retomar();
    };
    return destruir;

}
