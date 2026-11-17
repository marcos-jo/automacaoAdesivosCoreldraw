import { PDFArmazenamentoFormatos } from "./pdf-format-storage.js";
import { criarBackupFormatos, lerBackupFormatos } from "./pdf-format-backup.js";

// Referências ISO para reconhecer medidas de PDFs, não opções protegidas de UI.
export const FORMATOS_PADRAO = Object.freeze([
    { id: "A3", nome: "A3", largura: 297, altura: 420 },
    { id: "A4", nome: "A4", largura: 210, altura: 297 },
    { id: "A5", nome: "A5", largura: 148, altura: 210 },
    { id: "A6", nome: "A6", largura: 105, altura: 148 },
    { id: "A7", nome: "A7", largura: 74, altura: 105 }
].map(f => Object.freeze({ ...f })));
export const FORMATOS_INICIAIS = Object.freeze([
    ...FORMATOS_PADRAO,
    { id: 'F4', nome: 'F4', largura: 328, altura: 478 },
    { id: 'F8', nome: 'F8', largura: 238, altura: 328 },
    { id: 'F6', nome: 'F6', largura: 318, altura: 318 },
    { id: 'Faixa', nome: 'Faixa', largura: 320, altura: 700 }
].map(f => Object.freeze({ ...f })));

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
                if (dados.versao !== 1 || !Array.isArray(dados.formatos) || dados.formatos.length > 205) throw new Error();
                if (dados.catalogoCompleto !== undefined && dados.catalogoCompleto !== true) throw new Error();
                for (const item of dados.formatos) {
                    const inicial = FORMATOS_INICIAIS.some(f => f.id === item.id) || item.id === 'SRA3';
                    if (typeof item.id !== "string" || (!inicial && !/^custom-[\w-]+$/.test(item.id)) || this.personalizados.some(f => f.id === item.id || comparar(f.nome) === comparar(item.nome))) throw new Error();
                    if (typeof item.nome !== 'string' || !nomeNormalizado(item.nome) || nomeNormalizado(item.nome).length > 80) throw new Error();
                    medidaMm(item.largura); medidaMm(item.altura);
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
    registros() {
        // Somente ausência da chave significa primeiro uso. Documento existente,
        // inclusive legado ou vazio, é a lista integral e nunca recebe sugestões.
        return this.textoLido === null ? FORMATOS_INICIAIS : this.personalizados;
    }
    listar() { return this.registros().map(f => ({ ...f, largura: medidaMm(f.largura), altura: medidaMm(f.altura), padrao: false })); }
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
    conferirRegistro(id, esperado) {
        if (!esperado) return;
        const atual = this.registros().find(f => f.id === id);
        if (!atual || ['nome', 'largura', 'altura'].some(k => (k === 'nome' ? atual[k] : medidaMm(atual[k])) !== esperado[k])) throw new Error('Este tamanho mudou em outra ferramenta. Selecione novamente.');
    }
    async salvar(dados, id = null, esperado = null) {
        const entrada = { ...dados };
        return this.enfileirar(async () => {
            this.conferirRegistro(id, esperado);
            const registros = this.registros(), anterior = registros.find(f => f.id === id);
            if (id && !anterior) throw new Error("Formato não encontrado. Selecione novamente.");
            if (!id && registros.length >= 205) throw new Error("Limite de 205 formatos.");
            const formato = { ...anterior, ...this.validar(entrada, id), id: id || `custom-${globalThis.crypto?.randomUUID?.() || Date.now() + "-" + Math.random().toString(36).slice(2)}` };
            const proxima = id ? registros.map(f => f.id === id ? formato : f) : [...registros, formato];
            await this.persistir(proxima);
            return { ...formato };
        });
    }
    async excluir(id, esperado = null) {
        return this.enfileirar(async () => {
            this.conferirRegistro(id, esperado);
            const registros = this.registros();
            if (!registros.some(f => f.id === id)) throw new Error("Formato não encontrado. Selecione novamente.");
            await this.persistir(registros.filter(f => f.id !== id));
        });
    }
    documentoAtual() { return { ...this.documento, versao: 1, catalogoCompleto: true, formatos: this.registros() }; }
    async exportarBackup() {
        return this.enfileirar(async () => {
            await this.conferirArmazenamento();
            if (this.armazenamento.aviso) throw new Error('Armazenamento indisponível. Não redefina o catálogo antes de verificar seus dados persistidos.');
            return criarBackupFormatos(this.textoLido, this.documentoAtual());
        });
    }
    async redefinir(backup) {
        return this.enfileirar(async () => {
            const dados = await lerBackupFormatos(backup);
            if (dados.original !== this.textoLido || JSON.stringify(dados.documento) !== JSON.stringify(this.documentoAtual())) throw new Error('O backup não corresponde ao catálogo atual. Exporte e confira um novo backup.');
            if (this.armazenamento.aviso) throw new Error('Redefinição bloqueada: armazenamento indisponível.');
            await this.persistir(FORMATOS_INICIAIS.map(f => ({ ...f })), null, true);
        });
    }
    async restaurarBackup(backup) {
        return this.enfileirar(async () => {
            const dados = await lerBackupFormatos(backup);
            const texto = dados.original ?? JSON.stringify(dados.documento);
            // Validar a lista integral antes de substituir somente esta chave.
            const validacao = new PDFFormatos({ ler: () => texto });
            await validacao.pronta;
            if (validacao.bloqueado) throw new Error('O catálogo do backup é inválido. Nenhum dado foi alterado.');
            if (this.armazenamento.aviso) throw new Error('Restauração bloqueada: armazenamento indisponível.');
            await this.persistir(validacao.personalizados, texto, true);
        });
    }
    conferirArmazenamento() {
        // Conferência dentro do lock e sem await entre leitura/escrita síncronas.
        const conferir = atual => {
            if (atual !== this.textoLido) {
                this.aviso = "A biblioteca mudou em outra janela. Reabra o aplicativo ou recarregue esta janela antes de salvar; nenhum dado foi sobrescrito.";
                throw new Error(this.aviso);
            }
        };
        const atual = this.armazenamento.ler();
        return atual?.then ? atual.then(conferir) : conferir(atual);
    }
    async persistir(proxima, original = null, exigirPersistencia = false) {
        const conferencia = this.conferirArmazenamento();
        if (conferencia?.then) await conferencia;
        if (exigirPersistencia && this.armazenamento.aviso) throw new Error('Operação bloqueada: armazenamento indisponível. Backup preservado.');
        const documento = original === null ? { ...this.documento, versao: 1, catalogoCompleto: true, formatos: proxima } : JSON.parse(original);
        const texto = original ?? JSON.stringify(documento);
        const gravacao = this.armazenamento.gravar(texto);
        if (gravacao?.then) await gravacao;
        if (exigirPersistencia && this.armazenamento.aviso) throw new Error('Gravação persistente não confirmada. Backup preservado; interrompa e confira o armazenamento.');
        this.personalizados = proxima;
        this.documento = documento;
        this.textoLido = texto;
        this.aviso = "";
        this.notificar();
    }
    obterAviso() { return this.aviso || this.armazenamento.aviso || ""; }
}
