import { renderPDFInterface } from "./pdf-interface.js";
import { PDFPreview } from "./pdf-preview.js";
import { PDFMiniaturas } from "./pdf-thumbnails.js";
import { PDFProtecao } from "./pdf-protection.js";

import { salvarPaginasSelecionadas } from "./pdf-editor.js";

import { renderUnirPDFs } from "./pdf-merge.js";

import { PDFWorkspace } from "./pdf-workspace.js";

export function renderPDFs() {

    const app = document.getElementById("app");

    const raiz = renderPDFInterface(app, () => {
        if (!ocupado && documento && ferramentaAtiva()) mostrarPagina(paginaAtual);
    });

    // Cada instância guarda seus elementos, inclusive durante tarefas assíncronas.
    const elementos = new Map(
        [...app.querySelectorAll("[id]")].map(elemento => ["#" + elemento.id, elemento])
    );
    const buscar = seletor => elementos.get(seletor);

    const arquivoInput = buscar("#pdf-arquivo");
    const status = buscar("#pdf-status");
    const conteudo = buscar("#pdf-conteudo");
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
    const previa = new PDFPreview(visualizacao);
    const miniaturas = new PDFMiniaturas(buscar("#pdf-miniaturas"), numero => {
        if (!ocupado && ferramentaAtiva()) mostrarPagina(numero);
    });

    const uniao = renderUnirPDFs(app, adicionarAoDocumento);

    function ferramentaAtiva() {
        return !encerrada && app.querySelector("#pdf-ferramenta") === raiz;
    }

    function definirOcupado(valor) {
        ocupado = valor;
        arquivoInput.disabled = valor;
        informacoes.disabled = valor || !documento;

        for (const id of ["pdf-girar-esquerda", "pdf-girar-direita", "pdf-extrair", "pdf-excluir"]) {
            buscar("#" + id).disabled = valor || !documento;
        }

        buscar("#pdf-anterior").disabled = valor || !documento || paginaAtual <= 1;
        buscar("#pdf-proxima").disabled = valor || !documento || paginaAtual >= documento.numPages;
        uniao.definirOcupado(valor);
        miniaturas.definirOcupado(valor);
        atualizarBotaoSalvar();
        if (!valor) {
            const concluir = concluirOperacao;
            concluirOperacao = null;
            concluir?.();
        }
    }

    function iniciarOperacao() {
        if (ocupado || protecao.confirmacao || !ferramentaAtiva()) return false;
        fimOperacao = new Promise(resolve => { concluirOperacao = resolve; });
        // Uma navegação pendente não deve trocar a página durante a edição.
        previa.invalidar();
        definirOcupado(true);
        return true;
    }

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

    const formatar = numero =>
        numero.toLocaleString("pt-BR", {
            maximumFractionDigits: 2
        });

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

    async function mostrarPagina(numero) {
        if (!documento || !ferramentaAtiva()) return;
        const pdf = documento;
        try {
            const exibida = await previa.mostrar(numero, rotacoes.get(numero) || 0);
            if (!exibida || !ferramentaAtiva() || pdf !== documento) return;
            paginaAtual = numero;
            buscar("#pdf-pagina-atual").textContent = numero + " / " + pdf.numPages;
            buscar("#pdf-anterior").disabled = ocupado || numero === 1;
            buscar("#pdf-proxima").disabled = ocupado || numero === pdf.numPages;
            miniaturas.selecionar(numero);
        } catch (erro) {
            if (!ferramentaAtiva()) return;
            console.error(erro);
            status.textContent = "Erro ao visualizar a página.";
        }
    }

    //---------------------------------
    // Atualizar PDF após alterações
    //---------------------------------

    async function atualizarVisualizacaoWorkspace() {

        const bytes = await workspace.obterBytes();

        if (!ferramentaAtiva()) return false;
        const candidato = await previa.preparar(bytes);
        if (!candidato) return false;
        let adotada = false;
        try {
            const novoDocumento = candidato.documento;

            //---------------------------------
            // Recalcula as dimensões
            //---------------------------------

            const paginas = [];

            for (
                let numero = 1;
                numero <= novoDocumento.numPages;
                numero++
            ) {

                const pagina = await novoDocumento.getPage(numero);
                if (!ferramentaAtiva()) return false;

                const viewport = pagina.getViewport({
                    scale: 1
                });

                paginas.push({
                    numero,
                    largura: viewport.width * 25.4 / 72,
                    altura: viewport.height * 25.4 / 72
                });

            }

            const primeira = paginas[0];

            const tamanhosDiferentes = paginas.some(pagina =>
                Math.abs(pagina.largura - primeira.largura) > 0.1 ||
                Math.abs(pagina.altura - primeira.altura) > 0.1
            );

            if (!ferramentaAtiva()) return false;

            await miniaturas.limpar();
            if (!ferramentaAtiva() || !await previa.adotar(candidato)) return false;
            documento = novoDocumento;
            adotada = true;

            // As rotações já estão no workspace; não aplicá-las novamente.
            rotacoes.clear();

            //---------------------------------
            // Atualiza o resumo
            //---------------------------------

            buscar("#pdf-nome").textContent = workspace.nomeArquivo;
            conteudo.hidden = false;

            buscar("#pdf-total").textContent =
                documento.numPages;

            buscar("#pdf-tamanho").textContent =
                tamanhosDiferentes
                    ? "Variados"
                    : `${formatar(primeira.largura)} × ${formatar(primeira.altura)} mm`;

            informacoes.classList.toggle(
                "pdf-alerta",
                tamanhosDiferentes
            );

            informacoes.textContent =
                tamanhosDiferentes
                    ? "⚠ Informações · 1 alerta"
                    : "Informações";

            buscar("#pdf-alertas").textContent =
                tamanhosDiferentes
                    ? "Atenção: páginas com dimensões diferentes."
                    : "Nenhuma diferença de tamanho encontrada.";

            //---------------------------------
            // Atualiza a lista de páginas
            //---------------------------------

            const lista = buscar("#pdf-lista-paginas");

            lista.replaceChildren();

            for (const pagina of paginas) {

                const item = document.createElement("li");

                item.textContent =
                    `Página ${pagina.numero}: ` +
                    `${formatar(pagina.largura)} × ` +
                    `${formatar(pagina.altura)} mm`;

                lista.appendChild(item);

            }

            //---------------------------------
            // Atualiza a página selecionada
            //---------------------------------

            paginaAtual = Math.min(
                paginaAtual,
                documento.numPages
            );

            raiz.classList.toggle("pdf-com-documento", true);
            buscar("#pdf-painel-miniaturas").hidden = false;
            miniaturas.definirDocumento(documento, paginaAtual);
            await mostrarPagina(paginaAtual);

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
            ocupado || !workspace.pdf || !workspace.alterado;

    }


    //---------------------------------
    // Girar página no workspace
    //---------------------------------

    async function girarPagina(graus) {

        if (!documento || !iniciarOperacao()) return;

        try {

            //---------------------------------
            // Aplica no documento de trabalho
            //---------------------------------

            workspace.girarPagina(
                paginaAtual,
                graus
            );

            status.textContent =
                "Atualizando rotação...";

            //---------------------------------
            // Atualiza a pré-visualização
            //---------------------------------

            await atualizarVisualizacaoWorkspace();

            status.textContent =
                "Página girada. Alterações ainda não exportadas.";

        } catch (erro) {

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

                await atualizarVisualizacaoWorkspace();

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
            protecao.destruir();
            uniao.destruir();
            previa.invalidar();
            encerramento = (async () => {
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
    return destruir;

}
