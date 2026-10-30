import {
    PDFDocument,
    PDFName,
    PDFDict,
    PDFArray,
    PDFPageLeaf,
    degrees
} from "/vendor/pdf-lib/pdf-lib.js";
import { redimensionarDocumento, validarRedimensionamento, verificarCancelamento } from "./pdf-resize-engine.js";
import { diagnosticarAssinaturas } from "./pdf-signatures.js";
import { reconstruirCopiaImpressao, cancelarImpressao } from "./pdf-print-copy.js";

function origemAssinaturas(pdf, tamanho, arquivo, paginasDestino) {
    const diagnostico = diagnosticarAssinaturas(pdf, { tamanhoArquivo: tamanho });
    return { arquivo, diagnostico, resolvida: false,
        paginasVisuaisIds: [...new Set(diagnostico.assinaturas.flatMap(s => s.widgets.filter(w => w.visivel).map(w => paginasDestino[w.pagina - 1])))] };
}


export class PDFWorkspace {

    constructor(aoAlterar = () => {}) {

        this.pdf = null;

        this.nomeArquivo = "";

        this.revisao = 0;
        this.revisaoExportada = 0;
        this.aoAlterar = aoAlterar;
        this.origensAssinaturas = [];
        this.preparoImpressao = null;

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
            revisao: this.revisao, revisaoExportada: this.revisaoExportada,
            origensAssinaturas: this.origensAssinaturas, preparoImpressao: this.preparoImpressao
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

        const pdf = await PDFDocument.load(dados, { updateMetadata: false });
        // Diagnosticar os bytes de origem ANTES de qualquer save(). Não valida certificados.
        const origem = origemAssinaturas(pdf, dados.byteLength, arquivo.name, pdf.getPages().map(p => p.ref.toString()));

        this.pdf = pdf;

        this.nomeArquivo = arquivo.name;
        this.origensAssinaturas = [origem];
        this.preparoImpressao = null;

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
        const origens = [...this.origensAssinaturas];

        for (const arquivo of arquivos) {

            try {

                const dados = await arquivo.arrayBuffer();
                const origem = await PDFDocument.load(dados, { updateMetadata: false });
                const paginas = await resultado.copyPages(origem, origem.getPageIndices());

                if (!paginas.length) {
                    throw new Error("O arquivo não possui páginas.");
                }

                for (const pagina of paginas) {
                    resultado.addPage(pagina);
                }
                origens.push(origemAssinaturas(origem, dados.byteLength, arquivo.name, paginas.map(p => p.ref.toString())));

                adicionadas += paginas.length;

            } catch (erro) {

                throw new Error(`Falha ao processar "${arquivo.name}": ${erro.message}`);

            }

        }

        this.pdf = resultado;
        this.nomeArquivo ||= nomeArquivo;
        this.origensAssinaturas = origens;
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

    async duplicarPaginas(numeros) {
        const total = this.getTotalPaginas(), selecao = [...new Set(numeros)].sort((a, b) => a - b);
        if (!this.pdf || !selecao.length || selecao.some(n => !Number.isInteger(n) || n < 1 || n > total)) throw new Error("Seleção de duplicação inválida.");
        // copyPages não remapeia com segurança árvores de formulário, destinos e
        // semântica compartilhada para as novas identidades. Bloqueio conservador.
        const proibidas = new Set(["Annots", "AcroForm", "StructTreeRoot", "StructParents", "StructParent", "Outlines", "Dests", "Names", "OpenAction", "AA", "OCProperties", "Threads", "Perms", "PageLabels", "PrintPageRange"]);
        const atributosPagina = new Set(["Type", "Parent", "Resources", "MediaBox", "CropBox", "BleedBox", "TrimBox", "ArtBox", "Rotate", "Contents", "Group", "UserUnit", "Metadata", "LastModified", "Annots"]);
        this.pdf.catalog.Pages().traverse(pagina => {
            if (!(pagina instanceof PDFPageLeaf)) return;
            for (const chave of pagina.keys()) if (!atributosPagina.has(chave.decodeText())) throw new Error(`Duplicação bloqueada: atributo de página ${chave.decodeText()} ainda não validado. O documento foi preservado.`);
        });
        const visitados = new Set();
        const verificar = objeto => {
            if (!objeto || visitados.has(objeto)) return; visitados.add(objeto);
            if (objeto instanceof PDFDict) for (const [chave, valor] of objeto.entries()) {
                if (chave.decodeText() === "Annots") {
                    const anotacoes = this.pdf.context.lookup(valor);
                    if (anotacoes instanceof PDFArray && anotacoes.size() === 0) continue;
                }
                if (proibidas.has(chave.decodeText())) throw new Error(`Duplicação bloqueada: ${chave.decodeText()}. Este PDF possui recursos internos cuja cópia fiel ainda não foi validada. O documento foi preservado.`);
                verificar(valor);
            }
            else if (objeto instanceof PDFArray) objeto.asArray().forEach(verificar);
            else if (objeto.dict instanceof PDFDict) verificar(objeto.dict);
        };
        if (diagnosticarAssinaturas(this.pdf).estado !== "nao_identificada") throw new Error("Duplicação bloqueada: há campos ou estruturas de assinatura digital. Prepare uma cópia para impressão e confira antes de duplicar.");
        for (const [, objeto] of this.pdf.context.enumerateIndirectObjects()) verificar(objeto);
        verificar(this.pdf.catalog);
        const anterior = this.capturarVersao();
        const copia = await PDFDocument.load(await this.obterBytes(), { updateMetadata: false });
        const novas = await copia.copyPages(copia, selecao.map(n => n - 1));
        const apos = selecao.at(-1);
        novas.forEach((pagina, i) => copia.insertPage(apos + i, pagina));
        if (this.pdf !== anterior.pdf || this.revisao !== anterior.revisao) throw new Error("O documento mudou durante a duplicação. Tente novamente.");
        this.pdf = copia; this.registrarAlteracao();
        return novas.map((_, i) => apos + i + 1);
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

    async prepararRedimensionamento(config, { signal, validarImpressao } = {}) {
        if (!this.pdf) throw new Error("Nenhum PDF aberto.");
        validarRedimensionamento(config, this.getTotalPaginas());
        const versao = this.capturarVersao();
        verificarCancelamento(signal);
        const copia = await PDFDocument.load(await this.obterBytes(), { updateMetadata: false });
        verificarCancelamento(signal);
        const resultado = await redimensionarDocumento(copia, config, { signal, validarImpressao });
        if (this.pdf !== versao.pdf || this.revisao !== versao.revisao) throw new Error("O documento mudou durante a preparação.");
        return { ...resultado, pdf: copia, versao, origem: this, signal };
    }

    aplicarRedimensionamento(preparacao) {
        if (preparacao?.origem !== this || preparacao.signal?.aborted || preparacao.versao.pdf !== this.pdf || preparacao.versao.revisao !== this.revisao) {
            throw new Error("Prévia vencida ou cancelada. Gere uma nova prévia.");
        }
        if (!preparacao.alteradas.length && !preparacao.preparadoParaImpressao) return false;
        this.pdf = preparacao.pdf;
        this.registrarAlteracao();
        return true;
    }

    obterResumoAssinaturas() {
        const atuais = new Map(this.obterPaginas().map((p, i) => [p.id, i + 1]));
        const fontes = this.origensAssinaturas.filter(o => o.diagnostico.estado !== "nao_identificada");
        const pendentes = fontes.filter(o => !o.resolvida);
        const soma = chave => fontes.reduce((n, o) => n + o.diagnostico[chave], 0);
        const somaPendente = chave => pendentes.reduce((n, o) => n + o.diagnostico[chave], 0);
        return { fontes, pendentes, presentes: soma("presentes"), vazios: soma("vazios"), inconclusivas: soma("inconclusivas"),
            presentesPendentes: somaPendente("presentes"), vaziosPendentes: somaPendente("vazios"), inconclusivasPendentes: somaPendente("inconclusivas"),
            invisiveis: fontes.reduce((n, o) => n + o.diagnostico.assinaturas.filter(s => s.invisivel).length, 0),
            invisiveisPendentes: pendentes.reduce((n, o) => n + o.diagnostico.assinaturas.filter(s => s.invisivel).length, 0),
            paginasVisuais: [...new Set(pendentes.flatMap(o => o.paginasVisuaisIds.map(id => atuais.get(id)).filter(Boolean)))].sort((a, b) => a - b),
            preparada: !!this.preparoImpressao && pendentes.length === 0,
            validade: "Diagnóstico dos arquivos de origem, antes de reserializar. Validade criptográfica não verificada; edições/reescrita podem invalidar assinaturas." };
    }

    async prepararCopiaImpressao(config, opcoes) {
        if (!this.pdf) throw new Error("Abra um PDF antes de preparar a cópia.");
        const versao = this.capturarVersao(); cancelarImpressao(opcoes?.signal);
        const bytes = await this.obterBytes();
        await opcoes?.carregarFonte?.(bytes);
        cancelarImpressao(opcoes?.signal);
        const copia = await PDFDocument.load(bytes, { updateMetadata: false });
        const resultado = await reconstruirCopiaImpressao(copia, config, opcoes);
        if (this.pdf !== versao.pdf || this.revisao !== versao.revisao) throw new Error("O documento mudou durante o preparo.");
        return { ...resultado, versao, origem: this, signal: opcoes?.signal };
    }

    aplicarCopiaImpressao(preparacao) {
        if (preparacao?.origem !== this || preparacao.signal?.aborted || preparacao.versao.pdf !== this.pdf || preparacao.versao.revisao !== this.revisao || preparacao.verificacao?.compativel !== true) throw new Error("Cópia sem conferência válida, vencida ou cancelada. Prepare novamente.");
        this.pdf = preparacao.pdf;
        let base = this.nomeArquivo.replace(/\.pdf$/i, "");
        if (this.preparoImpressao) base = base.replace(/_IMPRESSAO$/i, "");
        else if (/_IMPRESSAO$/i.test(base)) base += "_COPIA";
        this.nomeArquivo = base + "_IMPRESSAO.pdf";
        this.origensAssinaturas = this.origensAssinaturas.map(o => ({ ...o, resolvida: true }));
        this.preparoImpressao = { dpi: preparacao.plano.dpi, rasterizadas: [...preparacao.plano.rasterizadas], verificacao: preparacao.verificacao };
        this.registrarAlteracao();
    }

    obterNomeExportacao() {
        return this.obterResumoAssinaturas().preparada ? this.nomeArquivo : this.nomeArquivo.replace(/\.pdf$/i, "") + "_editado.pdf";
    }

}
