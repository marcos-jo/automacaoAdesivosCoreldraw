// Contrato mínimo ler/gravar; pode ser substituído por um adaptador de API.
const memoria = new Map();
const somenteMemoria = new Set();
export class PDFArmazenamentoFormatos {
    constructor(obterStorage = () => globalThis.localStorage, chave = "4printshops.pdf.formatos.v1") {
        this.obterStorage = obterStorage;
        this.chave = chave;
        this.aviso = "";
    }
    ler() {
        if (somenteMemoria.has(this.chave)) {
            this.aviso = "As últimas alterações de formatos estão disponíveis somente nesta sessão.";
            return memoria.get(this.chave) ?? null;
        }
        try {
            const valor = this.obterStorage().getItem(this.chave);
            if (valor !== null) memoria.set(this.chave, valor);
            else memoria.delete(this.chave);
            return valor;
        } catch {
            this.aviso = "Armazenamento indisponível. Formatos disponíveis somente nesta sessão.";
            return memoria.get(this.chave) ?? null;
        }
    }
    gravar(valor) {
        memoria.set(this.chave, valor);
        try { this.obterStorage().setItem(this.chave, valor); somenteMemoria.delete(this.chave); this.aviso = ""; }
        catch { somenteMemoria.add(this.chave); this.aviso = "Não foi possível persistir. O formato ficará somente nesta sessão."; }
    }
}
