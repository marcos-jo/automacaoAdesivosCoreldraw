export async function salvarPdfGirado( documento, rotacoes, nomeArquivo) {

    const { PDFDocument, degrees } = await import(
        "/vendor/pdf-lib/pdf-lib.js"
    );

    //---------------------------------
    // Lê o PDF original
    //---------------------------------

    const dados = await documento.getData();

    const pdf = await PDFDocument.load(dados);

    //---------------------------------
    // Aplica as rotações
    //---------------------------------

    for (const [numero, giro] of rotacoes) {

        if (giro === 0) continue;

        const pagina = pdf.getPage(numero - 1);

        const rotacaoOriginal =
            pagina.getRotation().angle;

        const novaRotacao =
            ((rotacaoOriginal + giro) % 360 + 360) % 360;

        pagina.setRotation(
            degrees(novaRotacao)
        );

    }

    //---------------------------------
    // Gera o novo PDF
    //---------------------------------

    const resultado = await pdf.save();

    const blob = new Blob(
        [resultado],
        { type: "application/pdf" }
    );

    const url = URL.createObjectURL(blob);

    const link = document.createElement("a");

    link.href = url;

    link.download =
        nomeArquivo.replace(/\.pdf$/i, "") +
        "_girado.pdf";

    document.body.appendChild(link);

    link.click();

    link.remove();

    setTimeout(() => {
        URL.revokeObjectURL(url);
    }, 60000);

}

//---------------------------------
// Extrair ou excluir páginas
//---------------------------------

export async function salvarPaginasSelecionadas(
    documento,
    numeros,
    operacao,
    rotacoes,
    nomeArquivo
) {

    const { PDFDocument, degrees } = await import(
        "/vendor/pdf-lib/pdf-lib.js"
    );

    if (!["extrair", "excluir"].includes(operacao)) {
        throw new Error("Operação inválida.");
    }

    //---------------------------------
    // Carrega o PDF original
    //---------------------------------

    const dados = await documento.getData();

    const original = await PDFDocument.load(dados);

    const total = original.getPageCount();

    const selecionadas = new Set(numeros);

    if (
        selecionadas.size === 0 ||
        [...selecionadas].some(numero =>
            !Number.isInteger(numero) ||
            numero < 1 ||
            numero > total
        )
    ) {
        throw new Error("Seleção de páginas inválida.");
    }

    if (
        operacao === "excluir" &&
        selecionadas.size === total
    ) {
        throw new Error(
            "Não é possível excluir todas as páginas."
        );
    }

    const normalizarAngulo = angulo =>
        ((angulo % 360) + 360) % 360;

    let resultado;

    //---------------------------------
    // EXTRAIR PÁGINAS
    //---------------------------------

    if (operacao === "extrair") {

        resultado = await PDFDocument.create();

        const indices = [...selecionadas]
            .sort((a, b) => a - b)
            .map(numero => numero - 1);

        const paginas = await resultado.copyPages(
            original,
            indices
        );

        for (let i = 0; i < paginas.length; i++) {

            const pagina = paginas[i];

            const numeroOriginal = indices[i] + 1;

            const giro = rotacoes.get(numeroOriginal) || 0;

            if (giro !== 0) {

                const angulo = normalizarAngulo(
                    pagina.getRotation().angle + giro
                );

                pagina.setRotation(degrees(angulo));

            }

            resultado.addPage(pagina);

        }

    }

    //---------------------------------
    // EXCLUIR PÁGINAS
    //---------------------------------

    else {

        resultado = original;

        // Aplica rotações nas páginas mantidas.

        for (let i = 0; i < total; i++) {

            const numero = i + 1;

            if (selecionadas.has(numero)) {
                continue;
            }

            const giro = rotacoes.get(numero) || 0;

            if (giro !== 0) {

                const pagina = resultado.getPage(i);

                const angulo = normalizarAngulo(
                    pagina.getRotation().angle + giro
                );

                pagina.setRotation(degrees(angulo));

            }

        }

        // Exclui de trás para frente,
        // preservando os índices.

        for (let i = total - 1; i >= 0; i--) {

            if (selecionadas.has(i + 1)) {

                resultado.removePage(i);

            }

        }

    }

    //---------------------------------
    // Gera o arquivo PDF
    //---------------------------------

    const bytes = await resultado.save();

    const blob = new Blob(
        [bytes],
        { type: "application/pdf" }
    );

    const url = URL.createObjectURL(blob);

    const sufixo = operacao === "extrair"
        ? "_extraidas.pdf"
        : "_paginas_excluidas.pdf";

    const link = document.createElement("a");

    link.href = url;

    link.download =
        nomeArquivo.replace(/\.pdf$/i, "") + sufixo;

    document.body.appendChild(link);

    link.click();

    link.remove();

    setTimeout(() => {

        URL.revokeObjectURL(url);

    }, 60000);

}