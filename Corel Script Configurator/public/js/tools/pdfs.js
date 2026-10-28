import { salvarPaginasSelecionadas } from "./pdf-editor.js";

import { renderUnirPDFs } from "./pdf-merge.js";

import { PDFWorkspace } from "./pdf-workspace.js";

export function renderPDFs() {

    const app = document.getElementById("app");

    app.innerHTML = `
        <div class="pdf-topo">
            <h2>Gerenciador de PDFs</h2>

            <label class="pdf-importar">
                Importar PDF
                <input id="pdf-arquivo"
                    type="file"
                    accept=".pdf,application/pdf"
                    hidden>
            </label>
        </div>

        <p id="pdf-status">
            Selecione um arquivo PDF para começar.
        </p>

        <div id="pdf-conteudo" hidden>

            <h3 id="pdf-nome"></h3>

            <div class="pdf-resumo">
                <div>
                    <small>Páginas</small>
                    <strong id="pdf-total">-</strong>
                </div>

                <div>
                    <small>Dimensões</small>
                    <strong id="pdf-tamanho">-</strong>
                </div>

                <div>
                    <small>Cores</small>
                    <strong>Não verificado</strong>
                </div>
            </div>

            <div class="pdf-barra">
                <div class="pdf-navegacao">
                    <button id="pdf-anterior">Anterior</button>

                    <span id="pdf-pagina-atual"></span>

                    <button id="pdf-proxima">Próxima</button>
                </div>

                <div class="pdf-acoes">

                    <button id="pdf-girar-esquerda">
                        ↶ Girar esquerda
                    </button>

                    <button id="pdf-girar-direita">
                        ↷ Girar direita
                    </button>

                    <button id="pdf-salvar-copia" disabled>
                        Salvar cópia
                    </button>

                </div>

                    <button id="pdf-informacoes">
                        Informações
                    </button>
            </div>

            <div class="pdf-selecao">

                <label for="pdf-selecao-paginas">
                    Selecionar páginas
                </label>

                <div class="pdf-selecao-controles">

                    <input
                        id="pdf-selecao-paginas"
                        type="text"
                        placeholder="Ex.: 3, 7, 12 ou 4-8"
                    >

                    <button id="pdf-extrair">
                        Extrair páginas
                    </button>

                    <button id="pdf-excluir">
                        Excluir páginas
                    </button>

                </div>

                <small>
                    Informe páginas separadas por vírgula
                    ou intervalos com hífen.
                    O arquivo original será preservado.
                </small>

            </div>

            <div id="pdf-detalhes"
                class="pdf-detalhes"
                hidden>

                <h3>Informações do arquivo</h3>

                <p id="pdf-alertas"></p>

                <p>
                    <strong>Cores:</strong>
                    Análise RGB/CMYK ainda não implementada.
                </p>

                <p>
                    <strong>Sangria, fontes e resolução:</strong>
                    Ainda não verificadas.
                </p>

                <h4>Dimensões por página</h4>

                <ul id="pdf-lista-paginas"></ul>
            </div>

            <div id="pdf-visualizacao"
                class="pdf-visualizacao">
            </div>

        </div>
    `;

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

    // O conteúdo é recriado ao abrir a ferramenta; #app permanece na página.
    conteudo.addEventListener("click", (evento) => {

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
    let tarefaPDF = null;
    let paginaAtual = 1;
    let renderAtual = 0;
    let rotacoes = new Map();

    //---------------------------------
    // Documento de trabalho
    //---------------------------------

    const workspace = new PDFWorkspace();
    let ocupado = false;

    const uniao = renderUnirPDFs(app, adicionarAoDocumento);

    function ferramentaAtiva() {
        return app.querySelector("#pdf-conteudo") === conteudo;
    }

    function definirOcupado(valor) {
        ocupado = valor;
        arquivoInput.disabled = valor;

        for (const id of ["pdf-girar-esquerda", "pdf-girar-direita", "pdf-extrair", "pdf-excluir"]) {
            buscar("#" + id).disabled = valor || !documento;
        }

        buscar("#pdf-anterior").disabled = valor || !documento || paginaAtual <= 1;
        buscar("#pdf-proxima").disabled = valor || !documento || paginaAtual >= documento.numPages;
        uniao.definirOcupado(valor);
        atualizarBotaoSalvar();
    }

    function iniciarOperacao() {
        if (ocupado || !ferramentaAtiva()) return false;
        // Uma navegação pendente não deve trocar a página durante a edição.
        ++renderAtual;
        definirOcupado(true);
        return true;
    }

    async function adicionarAoDocumento(arquivos, nomeArquivo) {
        if (!iniciarOperacao()) {
            throw new Error("Aguarde a operação atual terminar.");
        }

        const anterior = {
            pdf: workspace.pdf,
            nomeArquivo: workspace.nomeArquivo,
            alterado: workspace.alterado
        };

        status.textContent = "Adicionando PDFs ao documento...";

        try {
            const adicionadas = await workspace.adicionarArquivos(arquivos, nomeArquivo);
            if (await atualizarVisualizacaoWorkspace()) {
                status.textContent = `${adicionadas} página(s) adicionada(s). Alterações ainda não exportadas.`;
            }
            return adicionadas;
        } catch (erro) {
            Object.assign(workspace, anterior);
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

        const anterior = {
            pdf: workspace.pdf,
            nomeArquivo: workspace.nomeArquivo,
            alterado: workspace.alterado
        };
        const paginaAnterior = paginaAtual;

        status.textContent = "Analisando PDF...";

        try {
            await workspace.abrir(arquivo);
            paginaAtual = 1;

            if (await atualizarVisualizacaoWorkspace()) {
                detalhes.hidden = true;
                status.textContent = "PDF carregado.";
            }
        } catch (erro) {
            Object.assign(workspace, anterior);
            paginaAtual = paginaAnterior;
            console.error(erro);
            status.textContent = "Não foi possível abrir o PDF. Verifique o arquivo ou o console.";
        } finally {
            definirOcupado(false);
        }
    });

    async function mostrarPagina(numero) {

        if (!documento || !ferramentaAtiva()) return;

        const render = ++renderAtual;
        const pdf = documento;

        try {

            const pagina = await pdf.getPage(numero);

            if (render !== renderAtual || pdf !== documento || !ferramentaAtiva()) {
                return;
            }

            const giro = rotacoes.get(numero) || 0;

            const angulo =
                ((pagina.rotate + giro) % 360 + 360) % 360;

            const base = pagina.getViewport({
                scale: 1,
                rotation: angulo
            });

            const escala = Math.min(
                1.8,
                850 / base.width,
                1100 / base.height
            );

            const viewport = pagina.getViewport({
                scale: escala,
                rotation: angulo
            });

            const canvas = document.createElement("canvas");

            canvas.width = Math.ceil(viewport.width);
            canvas.height = Math.ceil(viewport.height);

            const contexto = canvas.getContext("2d");

            await pagina.render({
                canvasContext: contexto,
                viewport,
                canvas
            }).promise;

            if (render !== renderAtual || pdf !== documento || !ferramentaAtiva()) {
                return;
            }

            visualizacao.replaceChildren(canvas);

            paginaAtual = numero;

            buscar("#pdf-pagina-atual").textContent =
                `${numero} / ${pdf.numPages}`;

            buscar("#pdf-anterior").disabled = ocupado || numero === 1;

            buscar("#pdf-proxima").disabled =
                ocupado || numero === pdf.numPages;

        } catch (erro) {

            if (render !== renderAtual || !ferramentaAtiva()) return;

            console.error(erro);
            status.textContent = "Erro ao visualizar a página.";
        }
    }

    //---------------------------------
    // Atualizar PDF após alterações
    //---------------------------------

    async function atualizarVisualizacaoWorkspace() {

        const bytes = await workspace.obterBytes();

        const pdfjs = await import(
            "/vendor/pdfjs/pdf.js"
        );

        pdfjs.GlobalWorkerOptions.workerSrc =
            "/vendor/pdfjs/pdf.worker.js";

        // Prepara a nova visualização antes de encerrar a anterior.

        const novaTarefa = pdfjs.getDocument({
            data: new Uint8Array(bytes)
        });

        let adotada = false;

        try {

            const novoDocumento = await novaTarefa.promise;

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

            ++renderAtual;
            const tarefaAnterior = tarefaPDF;
            tarefaPDF = null;
            documento = null;

            if (tarefaAnterior) {
                await tarefaAnterior.destroy();
            }

            tarefaPDF = novaTarefa;
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

            await mostrarPagina(paginaAtual);

            atualizarBotaoSalvar();

            return true;

        } finally {
            if (!adotada) await novaTarefa.destroy();
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

}
