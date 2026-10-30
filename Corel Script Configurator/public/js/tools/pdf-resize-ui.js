import { PDFFormatos, medidaMm } from "./pdf-formats.js";
import { PDFProporcaoMedidas } from "./pdf-dimension-ratio.js";

export class PDFRedimensionarUI {
    constructor(buscar, acoes, formatos = new PDFFormatos(), confirmar = mensagem => globalThis.confirm(mensagem)) {
        Object.assign(this, { buscar, acoes, formatos, confirmar });
        this.listeners = [];
        this.geracao = 0;
        this.evento = (id, tipo, handler) => {
            const elemento = buscar("#" + id); elemento.addEventListener(tipo, handler);
            this.listeners.push(() => elemento.removeEventListener(tipo, handler));
        };
        this.campo = id => buscar("#pdf-redim-" + id);
        this.mensagem = texto => { buscar("#pdf-redim-resultado").textContent = texto; };
        this.campo("largura").value = "210"; this.campo("altura").value = "297";
        this.campo("modo").value = "ajustar"; this.campo("alvo").value = "atual";
        this.campo("orientacao").value = "retrato";
        this.proporcao = new PDFProporcaoMedidas(buscar);
        this.atualizarLista("A4");
        formatos.pronta?.then(() => { if (!this.encerrada) this.atualizarLista(this.campo("formato").value); });
        this.evento("pdf-redim-formato", "change", () => {
            const escolhido = formatos.listar().find(f => f.id === this.campo("formato").value);
            if (escolhido) {
                this.campo("largura").value = String(escolhido.largura).replace(".", ",");
                this.campo("altura").value = String(escolhido.altura).replace(".", ",");
                this.campo("orientacao").value = escolhido.largura > escolhido.altura ? "paisagem" : "retrato";
                buscar("#pdf-formato-nome").value = escolhido.padrao ? "" : escolhido.nome;
            }
            this.proporcao.sincronizar(); this.atualizarBiblioteca(); this.invalidar();
        });
        for (const id of ["largura", "altura", "intervalo"]) this.evento("pdf-redim-" + id, "input", () => {
            if (id !== "intervalo") {
                this.proporcao.alterar(id);
                const selecionado = formatos.listar().find(f => f.id === this.campo("formato").value);
                if (selecionado?.padrao) this.campo("formato").value = "personalizado";
                this.atualizarBiblioteca();
            }
            this.invalidar();
        });
        for (const id of ["alvo", "modo"]) this.evento("pdf-redim-" + id, "change", () => this.invalidar());
        this.evento("pdf-redim-orientacao", "change", () => {
            this.tentar(() => {
                const a = medidaMm(this.campo("largura").value), b = medidaMm(this.campo("altura").value);
                const paisagem = this.campo("orientacao").value === "paisagem";
                this.campo("largura").value = String(paisagem ? Math.max(a, b) : Math.min(a, b)).replace(".", ",");
                this.campo("altura").value = String(paisagem ? Math.min(a, b) : Math.max(a, b)).replace(".", ",");
                this.proporcao.sincronizar();
                this.invalidar();
            });
        });
        this.evento("pdf-redim-prever", "click", () => this.tentar(() => this.prever()));
        this.evento("pdf-redim-impressao", "click", () => this.tentar(async () => {
            if (!this.diagnostico?.podePreparar || this.ocupado) return;
            const recursos = this.diagnostico.removiveis.map(r => "• " + r.descricao).join("\n");
            if (!this.confirmar(`Preparar cópia para impressão?\n\nSerão removidos do documento INTEIRO, mesmo ao redimensionar apenas uma página:\n${recursos}\n\nLinks internos, acessibilidade e semântica estruturada poderão ser removidos. Textos, fontes, imagens e vetores serão mantidos; não haverá rasterização. O arquivo original não será modificado.\n\nA próxima etapa é uma prévia temporária: Cancelar preserva o workspace; Aplicar confirma a preparação e o redimensionamento. Exportação somente em Salvar cópia.\n\nContinuar?`)) return;
            await this.prever(true);
        }));
        this.evento("pdf-redim-cancelar", "click", () => this.tentar(async () => {
            if (await this.invalidar()) this.mensagem("Prévia descartada. Documento de trabalho preservado.");
        }));
        this.evento("pdf-redim-aplicar", "click", () => this.tentar(async () => {
            try {
                await acoes.aplicar();
                this.mensagem("Redimensionamento aplicado ao documento. Exporte em Salvar cópia quando desejar.");
            } finally { this.pronta = false; this.atualizarEstado(); }
        }));
        for (const [id, editar] of [["pdf-formato-salvar", false], ["pdf-formato-atualizar", true]]) this.evento(id, "click", () => this.tentar(async () => {
            const f = await formatos.salvar({ nome: buscar("#pdf-formato-nome").value, largura: this.campo("largura").value, altura: this.campo("altura").value }, editar ? this.campo("formato").value : null);
            this.atualizarLista(f.id); buscar("#pdf-formatos-status").textContent = formatos.obterAviso() || "Formato salvo neste navegador.";
        }));
        this.evento("pdf-formato-excluir", "click", () => this.tentar(async () => {
            const f = formatos.listar().find(f => f.id === this.campo("formato").value);
            if (!f || f.padrao || !confirmar(`Excluir o formato personalizado “${f.nome}”?`)) return;
            await formatos.excluir(f.id); this.atualizarLista("personalizado");
            buscar("#pdf-formatos-status").textContent = formatos.obterAviso() || "Formato excluído. PDF preservado.";
        }));
        this.atualizarEstado();
    }
    async tentar(acao) {
        try { await acao(); }
        catch (erro) {
            if (erro.name !== "AbortError" && !this.encerrada) {
                this.diagnostico = erro.diagnostico;
                this.revisaoDiagnostico = this.acoes.contexto().revisao;
                this.pronta = false; this.mensagem(erro.message); this.atualizarEstado();
            }
        }
    }
    async prever(modoImpressao = false) {
        const geracao = ++this.geracao;
        const config = { ...this.configuracao(), modoImpressao };
        this.diagnostico = null; this.pronta = false; this.atualizarEstado();
        this.mensagem("Preparando prévia temporária...");
        let resultado;
        try { resultado = await this.acoes.prever(config); }
        catch (erro) { if (geracao !== this.geracao || this.encerrada) return; throw erro; }
        if (!resultado || geracao !== this.geracao) return;
        const detalhe = resultado.detalhes.find(d => d.numero === config.paginaPrevia);
        const fmt = n => n.toLocaleString("pt-BR", { maximumFractionDigits: 2 });
        this.pronta = resultado.alteradas.length > 0 || !!resultado.preparadoParaImpressao;
        this.mensagem(`Prévia temporária — página ${config.paginaPrevia}: ${fmt(detalhe.original.largura)} × ${fmt(detalhe.original.altura)} mm → ${fmt(config.largura)} × ${fmt(config.altura)} mm. ${resultado.avisos.join(" ")}${this.pronta ? "" : " As páginas já têm essas dimensões; nada a aplicar."}`);
        this.atualizarEstado();
    }
    atualizarLista(id) {
        const lista = this.campo("formato"); lista.replaceChildren();
        for (const f of [...this.formatos.listar(), { id: "personalizado", nome: "Personalizado" }]) {
            const opcao = document.createElement("option"); opcao.value = f.id;
            opcao.textContent = f.largura ? `${f.nome} — ${f.largura} × ${f.altura} mm` : f.nome;
            lista.appendChild(opcao);
        }
        lista.value = id; this.atualizarBiblioteca();
        this.buscar("#pdf-formatos-status").textContent = this.formatos.obterAviso();
    }
    atualizarBiblioteca() {
        const f = this.formatos.listar().find(f => f.id === this.campo("formato").value);
        for (const id of ["pdf-formato-atualizar", "pdf-formato-excluir"]) this.buscar("#" + id).disabled = !f || !!f.padrao;
    }
    configuracao() {
        this.proporcao.validar();
        const { total, atual } = this.acoes.contexto();
        if (!total) throw new Error("Abra um PDF antes de pré-visualizar.");
        const alvo = this.campo("alvo").value;
        const paginas = alvo === "todas" ? Array.from({ length: total }, (_, i) => i + 1)
            : alvo === "intervalo" ? this.acoes.interpretar(this.campo("intervalo").value, total) : [atual];
        return { largura: medidaMm(this.campo("largura").value), altura: medidaMm(this.campo("altura").value),
            modo: this.campo("modo").value, paginas, paginaPrevia: paginas.includes(atual) ? atual : paginas[0] };
    }
    invalidar() {
        ++this.geracao;
        this.diagnostico = null;
        this.pronta = false; this.mensagem(""); this.atualizarEstado();
        return this.acoes.cancelar().then(() => true, erro => { this.mensagem(erro.message); return false; });
    }
    atualizarEstado({ ocupado = this.ocupado || false, ativa = this.ativa || false, aplicando = this.aplicando || false } = {}) {
        const descricoes = {
            ajustar: "Cabe inteiro; pode sobrar margem.",
            centralizar: "Mantém o tamanho original do conteúdo.",
            preencher: "Preenche toda a folha; pode cortar bordas."
        };
        this.campo("modo-descricao").textContent = descricoes[this.campo("modo").value] || "";
        this.ocupado = ocupado;
        this.ativa = ativa; this.aplicando = aplicando;
        if (this.diagnostico && this.revisaoDiagnostico !== this.acoes.contexto().revisao) this.diagnostico = null;
        this.campo("prever").disabled = ocupado || !this.acoes.contexto().total;
        this.campo("impressao").hidden = !this.diagnostico?.podePreparar;
        this.campo("impressao").disabled = ocupado || aplicando || !this.acoes.contexto().total;
        this.campo("aplicar").disabled = ocupado || !this.pronta;
        this.campo("cancelar").disabled = aplicando || !ativa && !this.pronta;
        for (const id of ["formato", "largura", "altura", "orientacao", "alvo", "intervalo", "modo"]) this.campo(id).disabled = aplicando;
        this.campo("vincular").disabled = aplicando;
        const intervalo = this.campo("alvo").value === "intervalo";
        this.campo("intervalo").hidden = !intervalo; this.buscar("#pdf-redim-intervalo-label").hidden = !intervalo;
        let texto = "";
        try {
            const config = this.configuracao(), contexto = this.acoes.contexto();
            const dimensoes = contexto.dimensoes || [];
            if (config.modo !== "ajustar" && config.paginas.some(n => {
                const d = dimensoes[n - 1]; if (!d) return false;
                return config.modo === "preencher" ? Math.abs(d.largura / d.altura - config.largura / config.altura) > .001 : d.largura > config.largura || d.altura > config.altura;
            })) texto = "Atenção: este modo pode cortar parte do conteúdo. Confira a prévia.";
        } catch { /* Validação completa ao pré-visualizar. */ }
        this.buscar("#pdf-redim-aviso").textContent = texto;
    }
    destruir() { this.encerrada = true; this.proporcao.destruir(); for (const remover of this.listeners) remover(); }
}
