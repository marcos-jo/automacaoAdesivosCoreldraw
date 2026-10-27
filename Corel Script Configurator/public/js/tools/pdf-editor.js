export async function salvarPdfGirado(
    documento,
    rotacoes,
    nomeArquivo
) {

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