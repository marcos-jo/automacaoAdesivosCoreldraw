import {
    salvarPdfGirado,
    salvarPaginasSelecionadas
} from "./pdf-editor.js";

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

    const buscar = seletor => app.querySelector(seletor);

    const arquivoInput = buscar("#pdf-arquivo");
    const status = buscar("#pdf-status");
    const conteudo = buscar("#pdf-conteudo");
    const visualizacao = buscar("#pdf-visualizacao");
    const informacoes = buscar("#pdf-informacoes");
    const detalhes = buscar("#pdf-detalhes");

    //---------------------------------
    // Controle dos botões de páginas
    //---------------------------------

    app.addEventListener("click", (evento) => {

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
    let carregamentoAtual = 0;
    let renderAtual = 0;
    let rotacoes = new Map();

    const formatar = numero =>
        numero.toLocaleString("pt-BR", {
            maximumFractionDigits: 2
        });

    arquivoInput.addEventListener("change", async () => {

        const arquivo = arquivoInput.files?.[0];

        if (!arquivo) return;

        arquivoInput.value = "";

        const carregamento = ++carregamentoAtual;
        ++renderAtual;

        rotacoes.clear();

        buscar("#pdf-salvar-copia").disabled = true;

        status.textContent = "Analisando PDF...";
        conteudo.hidden = true;
        detalhes.hidden = true;

        try {

            const pdfjs = await import(
                "/vendor/pdfjs/pdf.js"
            );

            pdfjs.GlobalWorkerOptions.workerSrc =
                "/vendor/pdfjs/pdf.worker.js";


            //---------------------------------
            // Fecha o carregamento anterior
            //---------------------------------

            const tarefaAnterior = tarefaPDF;

            tarefaPDF = null;
            documento = null;

            visualizacao.replaceChildren();

            if (tarefaAnterior) {

                await tarefaAnterior.destroy();

            }

            if (carregamento !== carregamentoAtual) {
                return;
            }

            //---------------------------------
            // Abre o novo PDF
            //---------------------------------

            const dados = new Uint8Array(
                await arquivo.arrayBuffer()
            );

            //---------------------------------
            // Carrega e registra o novo PDF
            //---------------------------------

            const tarefa = pdfjs.getDocument({
                data: dados
            });

            tarefaPDF = tarefa;

            const novoDocumento = await tarefa.promise;

            if (carregamento !== carregamentoAtual) {
                await tarefa.destroy();
                return;
            }

            documento = novoDocumento;

            const paginas = [];

            for (let numero = 1;
                 numero <= documento.numPages;
                 numero++) {

                status.textContent =
                    `Analisando página ${numero} de ${documento.numPages}...`;

                const pagina = await documento.getPage(numero);

                const viewport = pagina.getViewport({
                    scale: 1
                });

                const largura = viewport.width * 25.4 / 72;
                const altura = viewport.height * 25.4 / 72;

                paginas.push({
                    numero,
                    largura,
                    altura
                });

                if (carregamento !== carregamentoAtual) {
                    return;
                }
            }

            const primeira = paginas[0];

            const tamanhosDiferentes = paginas.some(pagina =>
                Math.abs(pagina.largura - primeira.largura) > 0.1 ||
                Math.abs(pagina.altura - primeira.altura) > 0.1
            );

            buscar("#pdf-nome").textContent = arquivo.name;

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
                    ? "Atenção: este PDF possui páginas com dimensões diferentes."
                    : "Nenhuma diferença de tamanho encontrada.";

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

            conteudo.hidden = false;
            paginaAtual = 1;

            await mostrarPagina(1);

            if (carregamento === carregamentoAtual) {
                status.textContent = "PDF carregado.";
            }

        } catch (erro) {

            if (carregamento !== carregamentoAtual) return;

            console.error(erro);

            conteudo.hidden = true;

            status.textContent =
                "Não foi possível abrir o PDF. Verifique o arquivo ou o console.";
        }
    });

    async function mostrarPagina(numero) {

        if (!documento) return;

        const render = ++renderAtual;
        const pdf = documento;

        try {

            const pagina = await pdf.getPage(numero);

            if (render !== renderAtual || pdf !== documento) {
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

            if (render !== renderAtual || pdf !== documento) {
                return;
            }

            visualizacao.replaceChildren(canvas);

            paginaAtual = numero;

            buscar("#pdf-pagina-atual").textContent =
                `${numero} / ${pdf.numPages}`;

            buscar("#pdf-anterior").disabled = numero === 1;

            buscar("#pdf-proxima").disabled =
                numero === pdf.numPages;

        } catch (erro) {

            if (render !== renderAtual) return;

            console.error(erro);
            status.textContent = "Erro ao visualizar a página.";
        }
    }

    buscar("#pdf-anterior").addEventListener("click", () => {
        if (documento && paginaAtual > 1) {
            mostrarPagina(paginaAtual - 1);
        }
    });

    buscar("#pdf-proxima").addEventListener("click", () => {
        if (documento && paginaAtual < documento.numPages) {
            mostrarPagina(paginaAtual + 1);
        }
    });

    informacoes.addEventListener("click", () => {
        detalhes.hidden = !detalhes.hidden;
    });

    //---------------------------------
    // Atualiza botão Salvar
    //---------------------------------

    function atualizarBotaoSalvar() {

        const alterado = [...rotacoes.values()]
            .some(giro => giro !== 0);

        buscar("#pdf-salvar-copia").disabled =
            !documento || !alterado;

    }


    //---------------------------------
    // Girar página selecionada
    //---------------------------------

    function girarPagina(graus) {

        if (!documento) return;

        const atual =
            rotacoes.get(paginaAtual) || 0;

        const novo =
            ((atual + graus) % 360 + 360) % 360;

        if (novo === 0) {

            rotacoes.delete(paginaAtual);

        } else {

            rotacoes.set(paginaAtual, novo);

        }

        atualizarBotaoSalvar();

        mostrarPagina(paginaAtual);

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

        if (!documento || rotacoes.size === 0) {
            return;
        }

        const pdfAtual = documento;

        const rotacoesAtuais = new Map(rotacoes);

        const nomeArquivo =
            buscar("#pdf-nome").textContent;

        const botao =
            buscar("#pdf-salvar-copia");

        botao.disabled = true;

        status.textContent =
            "Gerando PDF modificado...";

        try {

            await salvarPdfGirado(
                pdfAtual,
                rotacoesAtuais,
                nomeArquivo
            );

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

            atualizarBotaoSalvar();

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

    let processandoPaginas = false;

    async function processarPaginas(operacao) {

        if (!documento || processandoPaginas) {
            return;
        }

        const botaoExtrair = buscar("#pdf-extrair");
        const botaoExcluir = buscar("#pdf-excluir");

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

        processandoPaginas = true;

        botaoExtrair.disabled = true;
        botaoExcluir.disabled = true;

        arquivoInput.disabled = true;

        status.textContent = "Gerando novo PDF...";

        try {

            await salvarPaginasSelecionadas(
                pdfAtual,
                paginas,
                operacao,
                rotacoesAtuais,
                nomeArquivo
            );

            if (pdfAtual === documento) {

                status.textContent =
                    operacao === "extrair"
                        ? `${paginas.length} página(s) extraída(s) com sucesso.`
                        : `${paginas.length} página(s) excluída(s) da nova cópia.`;

            }

        } catch (erro) {

            console.error(erro);

            status.textContent =
                "Erro ao gerar PDF: " + erro.message;

        } finally {

            processandoPaginas = false;

            botaoExtrair.disabled = false;
            botaoExcluir.disabled = false;

            arquivoInput.disabled = false;

        }

    }

}