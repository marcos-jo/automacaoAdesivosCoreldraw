export function renderUnirPDFs(app) {

    //---------------------------------
    // Interface
    //---------------------------------

    const container = document.createElement("div");

    container.className = "pdf-uniao-container";

    container.innerHTML = `

        <button
            id="pdf-uniao-abrir"
            class="pdf-uniao-abrir"
            type="button">
            Unir PDFs
        </button>

        <dialog
            id="pdf-uniao-janela"
            class="pdf-uniao"
            aria-labelledby="pdf-uniao-titulo">

            <div class="pdf-uniao-cabecalho">

                <h3 id="pdf-uniao-titulo">
                    Unir arquivos PDF
                </h3>

                <button
                    id="pdf-uniao-fechar"
                    type="button"
                    aria-label="Fechar janela">
                    ×
                </button>

            </div>

            <div class="pdf-uniao-conteudo">

            <p>
                Selecione dois ou mais PDFs e organize
                a sequência antes de gerar o arquivo.
            </p>

            <label class="pdf-uniao-adicionar"
                   for="pdf-uniao-arquivos">
                + Adicionar PDFs
            </label>

            <input
                id="pdf-uniao-arquivos"
                type="file"
                accept=".pdf,application/pdf"
                multiple
                hidden
            >

            <ol id="pdf-uniao-lista"></ol>

            <div class="pdf-uniao-rodape">

                <label for="pdf-uniao-nome">
                    Nome do arquivo final
                </label>

                <input
                    id="pdf-uniao-nome"
                    type="text"
                    value="arquivos_unidos"
                >

                <div class="pdf-uniao-botoes">

                    <button id="pdf-uniao-limpar">
                        Limpar lista
                    </button>

                    <button id="pdf-uniao-gerar">
                        Unir e baixar
                    </button>

                </div>

            </div>

                <p id="pdf-uniao-status" aria-live="polite"></p>

            </div>

        </dialog>
    `;

    const painel = container.querySelector(
        "#pdf-uniao-janela"
    );

    app.querySelector("#pdf-status").before(container);

    //---------------------------------
    // Elementos
    //---------------------------------

    const buscar = seletor =>
        painel.querySelector(seletor);

    //---------------------------------
    // Abrir e fechar janela
    //---------------------------------

    container.querySelector("#pdf-uniao-abrir")
        .addEventListener("click", () => {

            painel.showModal();

        });

    buscar("#pdf-uniao-fechar")
        .addEventListener("click", () => {

            painel.close();

        });

    const entrada = buscar("#pdf-uniao-arquivos");
    const lista = buscar("#pdf-uniao-lista");
    const status = buscar("#pdf-uniao-status");

    const botaoGerar = buscar("#pdf-uniao-gerar");
    const botaoLimpar = buscar("#pdf-uniao-limpar");

    let arquivos = [];
    let processando = false;

    //---------------------------------
    // Atualizar lista
    //---------------------------------

    function atualizarLista() {

        lista.replaceChildren();

        arquivos.forEach((arquivo, indice) => {

            const item = document.createElement("li");

            item.className = "pdf-uniao-item";

            const nome = document.createElement("span");

            nome.className = "pdf-uniao-nome";

            nome.textContent =
                `${indice + 1}. ${arquivo.name}`;

            const acoes = document.createElement("div");

            acoes.className = "pdf-uniao-acoes";

            const criarBotao = (texto, acao, desabilitado) => {

                const botao = document.createElement("button");

                botao.textContent = texto;
                botao.dataset.acao = acao;
                botao.dataset.indice = indice;
                botao.disabled = desabilitado || processando;

                return botao;

            };

            acoes.appendChild(
                criarBotao("↑", "subir", indice === 0)
            );

            acoes.appendChild(
                criarBotao(
                    "↓",
                    "descer",
                    indice === arquivos.length - 1
                )
            );

            acoes.appendChild(
                criarBotao("×", "remover", false)
            );

            item.append(nome, acoes);

            lista.appendChild(item);

        });

        botaoGerar.disabled =
            processando || arquivos.length < 2;

        botaoLimpar.disabled =
            processando || arquivos.length === 0;

        entrada.disabled = processando;

    }

    //---------------------------------
    // Adicionar arquivos
    //---------------------------------

    entrada.addEventListener("change", () => {

        if (processando) return;

        const selecionados = Array.from(entrada.files);

        const validos = selecionados.filter(arquivo =>
            arquivo.name.toLowerCase().endsWith(".pdf")
        );

        arquivos.push(...validos);

        entrada.value = "";

        status.textContent =
            validos.length !== selecionados.length
                ? "Alguns arquivos foram ignorados porque não são PDFs."
                : `${arquivos.length} arquivo(s) na lista.`;

        atualizarLista();

    });

    //---------------------------------
    // Reorganizar ou remover
    //---------------------------------

    lista.addEventListener("click", evento => {

        if (processando) return;

        const botao = evento.target.closest(
            "button[data-acao]"
        );

        if (!botao) return;

        const indice = Number(botao.dataset.indice);
        const acao = botao.dataset.acao;

        if (
            !Number.isInteger(indice) ||
            indice < 0 ||
            indice >= arquivos.length
        ) {
            return;
        }

        if (acao === "remover") {

            arquivos.splice(indice, 1);

        } else {

            const destino = acao === "subir"
                ? indice - 1
                : indice + 1;

            if (destino < 0 || destino >= arquivos.length) {
                return;
            }

            [
                arquivos[indice],
                arquivos[destino]
            ] = [
                arquivos[destino],
                arquivos[indice]
            ];

        }

        atualizarLista();

    });

    //---------------------------------
    // Limpar lista
    //---------------------------------

    botaoLimpar.addEventListener("click", () => {

        if (processando) return;

        arquivos = [];

        status.textContent = "Lista limpa.";

        atualizarLista();

    });

    //---------------------------------
    // Unir arquivos
    //---------------------------------

    botaoGerar.addEventListener("click", async () => {

        if (processando || arquivos.length < 2) return;

        processando = true;

        const arquivosParaUnir = [...arquivos];

        atualizarLista();

        try {

            const { PDFDocument } = await import(
                "/vendor/pdf-lib/pdf-lib.js"
            );

            const resultado = await PDFDocument.create();

            let totalPaginas = 0;

            for (let i = 0; i < arquivosParaUnir.length; i++) {

                const arquivo = arquivosParaUnir[i];

                status.textContent =
                    `Processando ${i + 1} de ${arquivosParaUnir.length}: ` +
                    arquivo.name;

                try {

                    const dados = await arquivo.arrayBuffer();

                    const origem = await PDFDocument.load(dados);

                    const paginas = await resultado.copyPages(
                        origem,
                        origem.getPageIndices()
                    );

                    for (const pagina of paginas) {

                        resultado.addPage(pagina);

                    }

                    totalPaginas += paginas.length;

                } catch (erro) {

                    throw new Error(
                        `Falha ao processar "${arquivo.name}": ` +
                        erro.message
                    );

                }

            }

            //---------------------------------
            // Exportar resultado
            //---------------------------------

            status.textContent = "Gerando PDF final...";

            const bytes = await resultado.save();

            const blob = new Blob(
                [bytes],
                { type: "application/pdf" }
            );

            const url = URL.createObjectURL(blob);

            let nome = buscar("#pdf-uniao-nome")
                .value.trim()
                .replace(/[\\/:*?"<>|]/g, "_")
                .replace(/\.pdf$/i, "");

            if (!nome) {
                nome = "arquivos_unidos";
            }

            const link = document.createElement("a");

            link.href = url;
            link.download = nome + ".pdf";

            document.body.appendChild(link);

            link.click();
            link.remove();

            setTimeout(() => {
                URL.revokeObjectURL(url);
            }, 60000);

            status.textContent =
                `PDF gerado: ${arquivosParaUnir.length} arquivos, ` +
                `${totalPaginas} páginas.`;

        } catch (erro) {

            console.error(erro);

            status.textContent =
                "Erro ao unir PDFs: " + erro.message;

        } finally {

            processando = false;

            atualizarLista();

        }

    });

    atualizarLista();

}