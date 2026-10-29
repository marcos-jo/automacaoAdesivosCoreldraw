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
const comparar = nome => nome.normalize("NFKC").toLocaleLowerCase("pt-BR");

export class PDFFormatos {
    constructor(armazenamento = new PDFArmazenamentoFormatos()) {
        this.armazenamento = armazenamento;
        this.personalizados = [];
        this.aviso = "";
        this.pronta = this.carregar();
    }
    async carregar() {
        let texto;
        try { texto = this.armazenamento.ler(); if (texto?.then) texto = await texto; }
        catch { this.aviso = "Não foi possível carregar a biblioteca de formatos."; return; }
        if (texto !== null) {
            try {
                const dados = JSON.parse(texto);
                if (dados.versao !== 1 || !Array.isArray(dados.formatos) || dados.formatos.length > 200) throw new Error();
                for (const item of dados.formatos) {
                    if (typeof item.id !== "string" || !/^custom-[\w-]+$/.test(item.id) || this.personalizados.some(f => f.id === item.id)) throw new Error();
                    this.personalizados.push({ id: item.id, ...this.validar(item) });
                }
            } catch {
                this.personalizados = [];
                this.aviso = "Biblioteca armazenada inválida; foi ignorada. Salvar um formato substituirá esses dados.";
            }
        }
    }
    listar() { return [...FORMATOS_PADRAO, ...this.personalizados].map(f => ({ ...f })); }
    validar(dados, id = null) {
        const nome = nomeNormalizado(dados.nome);
        if (!nome || nome.length > 80) throw new Error("Informe um nome de 1 a 80 caracteres.");
        if (this.listar().some(f => f.id !== id && comparar(f.nome) === comparar(nome))) throw new Error("Já existe um formato com esse nome.");
        return { nome, largura: medidaMm(dados.largura), altura: medidaMm(dados.altura) };
    }
    async salvar(dados, id = null) {
        await this.pronta;
        if (id && !this.personalizados.some(f => f.id === id)) throw new Error("Somente formatos personalizados podem ser editados.");
        if (!id && this.personalizados.length >= 200) throw new Error("Limite de 200 formatos personalizados.");
        const formato = { ...this.validar(dados, id), id: id || `custom-${globalThis.crypto?.randomUUID?.() || Date.now() + "-" + Math.random().toString(36).slice(2)}` };
        if (id) this.personalizados = this.personalizados.map(f => f.id === id ? formato : f);
        else this.personalizados.push(formato);
        await this.persistir();
        return { ...formato };
    }
    async excluir(id) {
        await this.pronta;
        if (!this.personalizados.some(f => f.id === id)) throw new Error("Somente formatos personalizados podem ser excluídos.");
        this.personalizados = this.personalizados.filter(f => f.id !== id);
        await this.persistir();
    }
    async persistir() {
        await this.armazenamento.gravar(JSON.stringify({ versao: 1, formatos: this.personalizados }));
        this.aviso = "";
    }
    obterAviso() { return this.aviso || this.armazenamento.aviso || ""; }
}
