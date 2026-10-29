// Índices visuais sobre uma fonte PDF.js imutável. O workspace continua sendo
// dono da ordem, das rotações e dos bytes usados na extração/exportação.
export class PDFDocumentoVisual {
    constructor(fonte, paginas, workspace) {
        this.fonte = fonte;
        this.paginas = paginas;
        this.workspace = workspace;
        this.numPages = paginas.length;
    }

    static async criar(fonte, workspace, anterior = null, invalidarIds = new Set()) {
        const metadados = workspace.obterPaginas();
        if (fonte.numPages !== metadados.length) throw new Error("Prévia e workspace possuem quantidades diferentes.");
        const conhecidas = new Map(anterior?.paginas.map(p => [p.id, p]) || []);
        const paginas = [];
        for (let indice = 0; indice < metadados.length; indice++) {
            const dados = metadados[indice];
            const conhecida = !invalidarIds.has(dados.id) && conhecidas.get(dados.id);
            let largura = conhecida?.larguraBase, altura = conhecida?.alturaBase;
            if (!conhecida) {
                const pagina = await fonte.getPage(indice + 1);
                const base = pagina.getViewport({ scale: 1, rotation: 0 });
                largura = base.width;
                altura = base.height;
            }
            paginas.push({ ...dados, numeroFonte: indice + 1, larguraBase: largura, alturaBase: altura });
        }
        return new PDFDocumentoVisual(fonte, paginas, workspace);
    }

    atualizar(workspace) {
        const conhecidas = new Map(this.paginas.map(p => [p.id, p]));
        const paginas = workspace.obterPaginas().map(dados => {
            const anterior = conhecidas.get(dados.id);
            if (!anterior) throw new Error("A edição exige recarregar a fonte PDF.");
            return { ...anterior, ...dados };
        });
        if (!paginas.length || new Set(paginas.map(p => p.id)).size !== paginas.length) {
            throw new Error("Mapa visual de páginas inválido.");
        }
        return new PDFDocumentoVisual(this.fonte, paginas, workspace);
    }

    obterChaveMiniatura(numero) {
        const pagina = this.paginas[numero - 1];
        return pagina && `${pagina.id}:${pagina.rotacao}`;
    }

    obterDimensoes() {
        return this.paginas.map((p, indice) => {
            const girada = p.rotacao % 180 !== 0;
            return { numero: indice + 1,
                largura: (girada ? p.alturaBase : p.larguraBase) * 25.4 / 72,
                altura: (girada ? p.larguraBase : p.alturaBase) * 25.4 / 72 };
        });
    }

    async getPage(numero) {
        const dados = this.paginas[numero - 1];
        if (!dados) throw new Error("Página visual inválida.");
        const pagina = await this.fonte.getPage(dados.numeroFonte);
        return {
            pageNumber: numero,
            rotate: dados.rotacao,
            getViewport: (opcoes = {}) => pagina.getViewport({ ...opcoes, rotation: opcoes.rotation ?? dados.rotacao }),
            render: opcoes => pagina.render(opcoes)
        };
    }

    // Nunca extrair da fonte PDF.js antiga depois de girar/reordenar.
    getData() { return this.workspace.obterBytes(); }
}
