import {
    PDFDocument,
    degrees
} from "/vendor/pdf-lib/pdf-lib.js";


export class PDFWorkspace {

    constructor(aoAlterar = () => {}) {

        this.pdf = null;

        this.nomeArquivo = "";

        this.revisao = 0;
        this.revisaoExportada = 0;
        this.aoAlterar = aoAlterar;

    }

    get alterado() {
        return !!this.pdf && this.revisao !== this.revisaoExportada;
    }

    registrarAlteracao() {
        ++this.revisao;
        this.aoAlterar();
    }

    capturarVersao() {
        return { pdf: this.pdf, revisao: this.revisao };
    }

    marcarExportado(versao) {
        // Uma exportação antiga não pode quitar outra edição ou outro documento.
        if (!this.pdf || versao.pdf !== this.pdf || versao.revisao !== this.revisao) return false;
        this.revisaoExportada = versao.revisao;
        this.aoAlterar();
        return true;
    }

    capturarEstado() {
        return {
            pdf: this.pdf, nomeArquivo: this.nomeArquivo,
            revisao: this.revisao, revisaoExportada: this.revisaoExportada
        };
    }

    restaurarEstado(estado) {
        Object.assign(this, estado);
        this.aoAlterar();
    }

    //---------------------------------
    // Abrir documento
    //---------------------------------

    async abrir(arquivo) {

        const dados = await arquivo.arrayBuffer();

        const pdf = await PDFDocument.load(dados);

        this.pdf = pdf;

        this.nomeArquivo = arquivo.name;

        ++this.revisao;
        this.revisaoExportada = this.revisao;
        this.aoAlterar();

    }


    //---------------------------------
    // Acrescentar PDFs sem perder as edições
    //---------------------------------

    async adicionarArquivos(arquivos, nomeArquivo = "arquivos_unidos.pdf") {

        if (!arquivos.length) {
            throw new Error("Selecione ao menos um PDF.");
        }

        // Só substitui o documento se todos os arquivos forem válidos.
        const resultado = this.pdf
            ? await PDFDocument.load(await this.obterBytes())
            : await PDFDocument.create();

        let adicionadas = 0;

        for (const arquivo of arquivos) {

            try {

                const origem = await PDFDocument.load(await arquivo.arrayBuffer());
                const paginas = await resultado.copyPages(origem, origem.getPageIndices());

                if (!paginas.length) {
                    throw new Error("O arquivo não possui páginas.");
                }

                for (const pagina of paginas) {
                    resultado.addPage(pagina);
                }

                adicionadas += paginas.length;

            } catch (erro) {

                throw new Error(`Falha ao processar "${arquivo.name}": ${erro.message}`);

            }

        }

        this.pdf = resultado;
        this.nomeArquivo ||= nomeArquivo;
        this.registrarAlteracao();

        return adicionadas;

    }


    //---------------------------------
    // Quantidade de páginas
    //---------------------------------

    getTotalPaginas() {

        if (!this.pdf) return 0;

        return this.pdf.getPageCount();

    }


    //---------------------------------
    // Girar página
    //---------------------------------

    girarPagina(numero, graus) {

        if (!this.pdf) {
            throw new Error("Nenhum PDF aberto.");
        }

        if (
            !Number.isInteger(numero) ||
            numero < 1 ||
            numero > this.getTotalPaginas()
        ) {
            throw new Error("Página inválida.");
        }

        if (!Number.isInteger(graus) || graus % 90 !== 0) {
            throw new Error("Rotação inválida.");
        }

        const pagina = this.pdf.getPage(numero - 1);

        const atual = pagina.getRotation().angle;

        const novo =
            ((atual + graus) % 360 + 360) % 360;

        pagina.setRotation(degrees(novo));

        this.registrarAlteracao();

    }


    //---------------------------------
    // Excluir páginas
    //---------------------------------

    excluirPaginas(numeros) {

        if (!this.pdf) {
            throw new Error("Nenhum PDF aberto.");
        }

        const total = this.getTotalPaginas();

        const paginas = [...new Set(numeros)];

        if (
            paginas.length === 0 ||
            paginas.some(numero =>
                !Number.isInteger(numero) ||
                numero < 1 ||
                numero > total
            )
        ) {
            throw new Error("Seleção de páginas inválida.");
        }

        if (paginas.length === total) {
            throw new Error(
                "Não é possível excluir todas as páginas."
            );
        }

        // Excluir da última para a primeira.

        paginas.sort((a, b) => b - a);

        for (const numero of paginas) {

            this.pdf.removePage(numero - 1);

        }

        this.registrarAlteracao();

    }


    //---------------------------------
    // Gerar versão atualizada
    //---------------------------------

    async obterBytes() {

        if (!this.pdf) {
            throw new Error("Nenhum PDF aberto.");
        }

        return await this.pdf.save();

    }

}
