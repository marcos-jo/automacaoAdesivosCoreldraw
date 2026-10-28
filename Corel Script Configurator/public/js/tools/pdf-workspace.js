import {
    PDFDocument,
    PDFName,
    PDFPageLeaf,
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

    // Identidade e rotação vêm sempre do documento editável, nunca da posição.
    obterPaginas() {
        const paginas = [];
        // pdf-lib 1.17.1 não invalida getPages() em removePage(). Consultar a
        // árvore evita reutilizar posições excluídas sem tocar no cache privado.
        this.pdf?.catalog.Pages().traverse((node, ref) => {
            if (!(node instanceof PDFPageLeaf)) return;
            const angulo = node.getInheritableAttribute(PDFName.of("Rotate"))?.asNumber() || 0;
            paginas.push({ id: ref.toString(), rotacao: ((angulo % 360) + 360) % 360 });
        });
        return paginas;
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

        const identidade = this.obterPaginas()[numero - 1].id;
        const pagina = this.pdf.getPages().find(p => p.ref.toString() === identidade);

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
    // Reorganizar páginas, mantendo objetos e referências do PDF
    //---------------------------------

    async reordenarPaginas(ordem) {
        if (!this.pdf) throw new Error("Nenhum PDF aberto.");
        const total = this.getTotalPaginas();
        const sequencia = Array.isArray(ordem) ? Array.from(ordem) : [];
        if (sequencia.length !== total || new Set(sequencia).size !== total ||
            sequencia.some(numero => !Number.isInteger(numero) || numero < 1 || numero > total)) {
            throw new Error("Ordem de páginas inválida.");
        }
        if (sequencia.every((numero, indice) => numero === indice + 1)) return false;

        const anterior = this.capturarVersao();
        // Clona o documento completo, em vez de copiar páginas para um PDF vazio.
        const resultado = await PDFDocument.load(await this.obterBytes(), { updateMetadata: false });
        const paginas = resultado.getPages();
        for (const pagina of paginas) {
            // pdf-lib 1.17.1: mover folhas pode mudar seu Parent. Materializar a
            // herança antes de removê-las preserva recursos, caixas e rotação.
            for (const chave of ["Resources", "MediaBox", "CropBox", "Rotate"]) {
                const nome = PDFName.of(chave);
                const valor = pagina.node.getInheritableAttribute(nome);
                if (valor) pagina.node.set(nome, valor);
            }
        }
        for (let indice = total - 1; indice >= 0; indice--) resultado.removePage(indice);
        for (const numero of sequencia) resultado.addPage(paginas[numero - 1]);

        if (this.pdf !== anterior.pdf || this.revisao !== anterior.revisao) {
            throw new Error("O documento mudou durante a reorganização. Tente novamente.");
        }
        this.pdf = resultado;
        this.registrarAlteracao();
        return true;
    }

    async moverPagina(origem, destino) {
        const total = this.getTotalPaginas();
        if (!this.pdf || !Number.isInteger(origem) || !Number.isInteger(destino) ||
            origem < 1 || origem > total || destino < 1 || destino > total) {
            throw new Error("Movimento de página inválido.");
        }
        const ordem = Array.from({ length: total }, (_, indice) => indice + 1);
        ordem.splice(destino - 1, 0, ...ordem.splice(origem - 1, 1));
        return { alterada: await this.reordenarPaginas(ordem), ordem };
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
