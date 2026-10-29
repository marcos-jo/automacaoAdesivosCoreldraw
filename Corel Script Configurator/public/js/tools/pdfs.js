import { renderPDFInterface } from "./pdf-interface.js";
import { PDFPreview } from "./pdf-preview.js";
import { PDFMiniaturas } from "./pdf-thumbnails.js";
import { PDFProtecao } from "./pdf-protection.js";
import { PDFDocumentoVisual } from "./pdf-visual-document.js";
import { PDFInformacoesDocumento } from "./pdf-document-info.js";
import { PDFRedimensionarUI } from "./pdf-resize-ui.js";
import { PDFRedimensionarSessao } from "./pdf-resize-session.js";

import { salvarPaginasSelecionadas } from "./pdf-editor.js";

import { renderUnirPDFs } from "./pdf-merge.js";

import { PDFWorkspace } from "./pdf-workspace.js";

export function renderPDFs() {

    const app = document.getElementById("app");

    const raiz = renderPDFInterface(app, () => {
        if (!ocupado && documento && ferramentaAtiva()) {
            if (sessaoResize?.ativa) sessaoResize.reexibir().catch(erro => { status.textContent = erro.message; });
            else mostrarPagina(paginaAtual);
        }
    });

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
    let rotacoes = new Map();
    let encerrada = false;

    //---------------------------------
    // Documento de trabalho
    //---------------------------------

    let protecao;
    const workspace = new PDFWorkspace(() => protecao?.atualizar());
    protecao = new PDFProtecao(raiz, workspace);
    let ocupado = false;
    let fimOperacao = Promise.resolve();
    let concluirOperacao;
    let sessaoResize, interfaceResize;
    const previa = new PDFPreview(visualizacao);
    const miniaturas = new PDFMiniaturas(buscar("#pdf-miniaturas"), numero => {
        if (!ocupado && ferramentaAtiva()) mostrarPagina(numero);
    }, (origem, destino) => moverPagina(origem, destino));

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
        contexto: () => ({ total: documento?.numPages || 0, atual: paginaAtual, dimensoes: documento?.obterDimensoes?.() || [] }),
        interpretar: interpretarPaginas,
        prever: config => sessaoResize.prever(config),
        cancelar: () => sessaoResize.cancelar(), aplicar: () => sessaoResize.aplicar()
    });

    function ferramentaAtiva() {
        return !encerrada && app.querySelector("#pdf-ferramenta") === raiz;
    }

    function definirOcupado(valor) {
        ocupado = valor;
        const bloqueado = valor || !!sessaoResize?.ativa;
        arquivoInput.disabled = bloqueado;
        informacoes.disabled = bloqueado || !documento;

        for (const id of ["pdf-girar-esquerda", "pdf-girar-direita", "pdf-extrair", "pdf-excluir"]) {
            buscar("#" + id).disabled = bloqueado || !documento;
        }

        buscar("#pdf-anterior").disabled = bloqueado || !documento || paginaAtual <= 1;
        buscar("#pdf-proxima").disabled = bloqueado || !documento || paginaAtual >= documento.numPages;
        uniao.definirOcupado(bloqueado);
        miniaturas.definirOcupado(bloqueado);
        interfaceResize?.atualizarEstado({ ocupado: valor, ativa: !!sessaoResize?.ativa, aplicando: !!sessaoResize?.aplicando });
        atualizarControlesOrdem();
        atualizarBotaoSalvar();
        if (!valor) {
            const concluir = concluirOperacao;
            concluirOperacao = null;
            concluir?.();
        }
    }

    function iniciarOperacao(permitirPrevia = false) {
        if (ocupado || !permitirPrevia && sessaoResize?.ativa || protecao.confirmacao || !ferramentaAtiva()) return false;
        fimOperacao = new Promise(resolve => { concluirOperacao = resolve; });
        // Uma navegação pendente não deve trocar a página durante a edição.
        previa.invalidar();
        definirOcupado(true);
        return true;
    }

    function atualizarControlesOrdem() {
        const bloqueado = ocupado || !!sessaoResize?.ativa;
        buscar("#pdf-mover-cima").disabled = bloqueado || !documento || paginaAtual <= 1;
        buscar("#pdf-mover-baixo").disabled = bloqueado || !documento || paginaAtual >= documento.numPages;
        buscar("#pdf-mover-destino").disabled = bloqueado || !documento || documento.numPages < 2;
        buscar("#pdf-mover-aplicar").disabled = bloqueado || !documento || documento.numPages < 2;
        buscar("#pdf-mover-destino").max = documento?.numPages || 1;
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

    async function moverPagina(origem, destino, controle = null) {
        if (!documento || !iniciarOperacao()) return;
        const anterior = workspace.capturarEstado();
        const paginaAnterior = paginaAtual;
        try {
            const resultado = await workspace.moverPagina(origem, destino);
            if (!resultado.alterada) return;
            // A permutação conserva a identidade da página exibida, inclusive se
            // outra página tiver sido arrastada para antes/depois dela.
            paginaAtual = resultado.ordem.indexOf(paginaAnterior) + 1;
            if (await atualizarVisualizacaoWorkspace({ reutilizar: true })) {
                buscar("#pdf-selecao-paginas").value = "";
                status.textContent = `Página ${origem} movida para ${destino}. Alterações ainda não exportadas.`;
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
        moverPagina(paginaAtual, paginaAtual - 1, buscar("#pdf-mover-cima"));
    });
    buscar("#pdf-mover-baixo").addEventListener("click", () => {
        moverPagina(paginaAtual, paginaAtual + 1, buscar("#pdf-mover-baixo"));
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

            if (await atualizarVisualizacaoWorkspace()) {
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
        const pdf = documento;
        const versaoRolagem = miniaturas.versaoRolagem;
        try {
            const exibida = await previa.mostrar(numero, rotacoes.get(numero) || 0);
            if (!exibida || !ferramentaAtiva() || pdf !== documento) return;
            paginaAtual = numero;
            atualizarControlesOrdem();
            buscar("#pdf-pagina-atual").textContent = numero + " / " + pdf.numPages;
            buscar("#pdf-anterior").disabled = ocupado || numero === 1;
            buscar("#pdf-proxima").disabled = ocupado || numero === pdf.numPages;
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

    async function atualizarVisualizacaoWorkspace({ reutilizar = false, preservarPaginas = reutilizar, candidatoPreparado = null, invalidarIds = new Set() } = {}) {

        const anterior = documento;
        if (reutilizar && anterior instanceof PDFDocumentoVisual) {
            try {
                const novo = anterior.atualizar(workspace);
                if (!await previa.atualizarDocumento(novo) || !ferramentaAtiva()) return false;
                documento = novo;
                rotacoes.clear();
                await mostrarPagina(paginaAtual, { propagarErro: true });
                await miniaturas.atualizarDocumento(novo, paginaAtual);
                resumo.atualizar(novo.obterDimensoes(), workspace.nomeArquivo);
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

            // As rotações já estão no workspace; não aplicá-las novamente.
            rotacoes.clear();

            resumo.atualizar(novoDocumento.obterDimensoes(), workspace.nomeArquivo);

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
            mostrarPagina(paginaAtual - 1);
        }
    });

    buscar("#pdf-proxima").addEventListener("click", () => {
        if (!ocupado && documento && paginaAtual < documento.numPages) {
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
            ocupado || !!sessaoResize?.ativa || !workspace.pdf || !workspace.alterado;

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

    buscar("#pdf-salvar-copia")
    .addEventListener("click", async () => {

        if (!workspace.pdf || !workspace.alterado || !iniciarOperacao()) {
            return;
        }

        const pdfAtual = documento;
        const versaoExportada = workspace.capturarVersao();

        const nomeArquivo =
            buscar("#pdf-nome").textContent;

        status.textContent =
            "Gerando PDF modificado...";

        try {

            //---------------------------------
            // Obtém o PDF atualizado
            //---------------------------------

            const bytes = await workspace.obterBytes();

            if (!ferramentaAtiva()) return;

            //---------------------------------
            // Prepara o download
            //---------------------------------

            const blob = new Blob(
                [bytes],
                { type: "application/pdf" }
            );

            const url = URL.createObjectURL(blob);

            const link = document.createElement("a");

            link.href = url;

            link.download =
                nomeArquivo.replace(/\.pdf$/i, "") +
                "_editado.pdf";

            document.body.appendChild(link);

            link.click();
            workspace.marcarExportado(versaoExportada);

            link.remove();

            setTimeout(() => {
                URL.revokeObjectURL(url);
            }, 60000);

            if (pdfAtual === documento) {

                status.textContent =
                    "PDF modificado gerado com sucesso.";

            }

        } catch (erro) {

            console.error(erro);

            if (pdfAtual === documento) {

                status.textContent =
                    "Erro ao gerar PDF. Consulte o console.";

            }

        } finally {

            definirOcupado(false);

        }

    });

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

    async function processarPaginas(operacao) {

        if (!documento || ocupado || !ferramentaAtiva()) {
            return;
        }

        let paginas;

        try {

            paginas = interpretarPaginas(
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

                workspace.excluirPaginas(paginas);

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
                resumo.atualizar(documento.obterDimensoes(), workspace.nomeArquivo);
                buscar("#pdf-pagina-atual").textContent = `${paginaAtual} / ${documento.numPages}`;
            }
            status.textContent =
                "Erro ao gerar PDF: " + erro.message;

        } finally {

            definirOcupado(false);

        }

    }

    // main.js aguarda este encerramento antes de montar a próxima ferramenta.
    let encerramento;
    function destruir() {
        if (!encerramento) {
            encerrada = true;
            interfaceResize.destruir();
            protecao.destruir();
            uniao.destruir();
            previa.invalidar();
            encerramento = (async () => {
                await sessaoResize.cancelar();
                await miniaturas.destruir();
                await previa.destruir();
                documento = null;
            })();
        }
        return encerramento;
    }

    destruir.confirmarSaida = async () => {
        if (sessaoResize.preparando) await sessaoResize.cancelar();
        // Uma união/importação em andamento pode criar pendências ao terminar.
        while (ocupado) await fimOperacao;
        if (!ferramentaAtiva()) return false;
        return protecao.confirmarDescarte();
    };
    return destruir;

}
