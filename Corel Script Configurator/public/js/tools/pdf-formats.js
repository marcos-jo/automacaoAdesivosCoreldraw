import { PDFArmazenamentoFormatos } from "./pdf-format-storage.js";

export const FORMATOS_PADRAO = Object.freeze([
    { id: "A3", nome: "A3", largura: 297, altura: 420 },
    { id: "A4", nome: "A4", largura: 210, altura: 297 },
    { id: "A5", nome: "A5", largura: 148, altura: 210 },
    { id: "A6", nome: "A6", largura: 105, altura: 148 }
].map(f => Object.freeze({ ...f, padrao: true })));

export function medidaMm(valor) {
    const texto = String(valor ?? "").trim();
    if (!/^\d+(?:[.,]\d+)?$/.test(texto)) throw new Error("Informe medidas decimais, sem separador de milhar.");
    const numero = Number(texto.replace(",", "."));
    if (!Number.isFinite(numero) || numero < .1 || numero > 5000) throw new Error("As medidas devem ficar entre 0,1 e 5000 mm.");
    return numero;
}
const nomeNormalizado = nome => String(nome ?? "").trim().replace(/\s+/g, " ").normalize("NFC");
const comparar = nome => nomeNormalizado(nome).normalize("NFKC").toLocaleLowerCase("pt-BR");

export class PDFFormatos {
    constructor(armazenamento = new PDFArmazenamentoFormatos()) {
        this.armazenamento = armazenamento;
        this.personalizados = [];
        this.aviso = "";
        this.listeners = new Set();
        this.fila = Promise.resolve();
        this.documento = { versao: 1, formatos: [] };
        this.pronta = this.carregar();
    }
    async carregar() {
        let texto;
        try { texto = this.armazenamento.ler(); if (texto?.then) texto = await texto; }
        catch { this.bloqueado = true; this.aviso = "Não foi possível carregar a biblioteca de formatos. Edição bloqueada para preservar os dados."; return; }
        this.textoLido = texto;
        if (texto !== null) {
            try {
                const dados = JSON.parse(texto);
                if (dados.versao !== 1 || !Array.isArray(dados.formatos) || dados.formatos.length > 200) throw new Error();
                for (const item of dados.formatos) {
                    if (typeof item.id !== "string" || !/^custom-[\w-]+$/.test(item.id) || this.personalizados.some(f => f.id === item.id)) throw new Error();
                    this.validar(item);
                    if (typeof item.nome !== 'string') throw new Error();
                    this.personalizados.push({ ...item });
                }
                this.documento = { ...dados };
            } catch {
                this.personalizados = [];
                this.bloqueado = true;
                this.aviso = "Biblioteca armazenada inválida. Edição bloqueada; os dados originais foram preservados.";
            }
        }
    }
    listar() { return [...FORMATOS_PADRAO, ...this.personalizados.map(f => ({ ...f, largura: medidaMm(f.largura), altura: medidaMm(f.altura), padrao: false }))].map(f => ({ ...f })); }
    assinar(listener) { this.listeners.add(listener); return () => this.listeners.delete(listener); }
    notificar() { for (const listener of this.listeners) listener(); }
    enfileirar(acao) {
        const executar = async () => {
            await this.pronta;
            if (this.bloqueado) throw new Error(this.aviso);
            return this.armazenamento.executarExclusivo ? this.armazenamento.executarExclusivo(acao) : acao();
        };
        const resultado = this.fila.then(executar);
        this.fila = resultado.catch(() => {});
        return resultado;
    }
    validar(dados, id = null) {
        const nome = nomeNormalizado(dados.nome);
        if (!nome || nome.length > 80) throw new Error("Informe um nome de 1 a 80 caracteres.");
        if (this.listar().some(f => f.id !== id && comparar(f.nome) === comparar(nome))) throw new Error("Já existe um formato com esse nome.");
        return { nome, largura: medidaMm(dados.largura), altura: medidaMm(dados.altura) };
    }
    async salvar(dados, id = null) {
        const entrada = { ...dados };
        return this.enfileirar(async () => {
            const anterior = this.personalizados.find(f => f.id === id);
            if (id && !anterior) throw new Error("Somente formatos personalizados podem ser editados.");
            if (!id && this.personalizados.length >= 200) throw new Error("Limite de 200 formatos personalizados.");
            const formato = { ...anterior, ...this.validar(entrada, id), id: id || `custom-${globalThis.crypto?.randomUUID?.() || Date.now() + "-" + Math.random().toString(36).slice(2)}` };
            const proxima = id ? this.personalizados.map(f => f.id === id ? formato : f) : [...this.personalizados, formato];
            await this.persistir(proxima);
            return { ...formato };
        });
    }
    async excluir(id) {
        return this.enfileirar(async () => {
            if (!this.personalizados.some(f => f.id === id)) throw new Error("Somente formatos personalizados podem ser excluídos.");
            await this.persistir(this.personalizados.filter(f => f.id !== id));
        });
    }
    async persistir(proxima) {
        // Conferência dentro do lock e sem await entre leitura/escrita síncronas.
        let atual = this.armazenamento.ler();
        if (atual?.then) atual = await atual;
        if (atual !== this.textoLido) {
            this.aviso = "A biblioteca mudou em outra janela. Reabra o aplicativo ou recarregue esta janela antes de salvar; nenhum dado foi sobrescrito.";
            throw new Error(this.aviso);
        }
        const documento = { ...this.documento, versao: 1, formatos: proxima };
        const texto = JSON.stringify(documento);
        const gravacao = this.armazenamento.gravar(texto);
        if (gravacao?.then) await gravacao;
        this.personalizados = proxima;
        this.documento = documento;
        this.textoLido = texto;
        this.aviso = "";
        this.notificar();
    }
    obterAviso() { return this.aviso || this.armazenamento.aviso || ""; }
}
