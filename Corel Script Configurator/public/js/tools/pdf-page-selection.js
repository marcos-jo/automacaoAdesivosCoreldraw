// Identidades independentes da janela virtual e da posição atual de cada página.
export class PDFSelecaoPaginas {
    constructor() { this.ids = new Set(); this.lista = []; }
    atualizar(lista, atual = 1) {
        this.lista = lista;
        this.ids = new Set([...this.ids].filter(id => lista.includes(id)));
        if (!this.ids.size && lista.length) this.ids.add(lista[Math.min(atual - 1, lista.length - 1)]);
        if (!lista.includes(this.ancora)) this.ancora = [...this.ids][0];
    }
    clicar(numero, { shiftKey = false, ctrlKey = false, metaKey = false } = {}) {
        const id = this.lista[numero - 1]; if (!id) return;
        if (shiftKey && this.ancora) {
            const a = this.lista.indexOf(this.ancora), b = numero - 1;
            const intervalo = this.lista.slice(Math.min(a, b), Math.max(a, b) + 1);
            this.ids = new Set(ctrlKey || metaKey ? [...this.ids, ...intervalo] : intervalo);
        } else if (ctrlKey || metaKey) {
            if (this.ids.has(id)) this.ids.delete(id); else this.ids.add(id);
            this.ancora = id;
        } else { this.ids = new Set([id]); this.ancora = id; }
    }
    numeros() { return this.lista.flatMap((id, i) => this.ids.has(id) ? [i + 1] : []); }
    limpar() { this.ids.clear(); this.ancora = null; }
}

export function ordemMoverPaginas(total, numeros, destino, divisoria) {
    const mover = [...new Set(numeros)].sort((a, b) => a - b);
    if (!mover.length || mover.some(n => !Number.isInteger(n) || n < 1 || n > total)) throw new Error("Seleção de movimento inválida.");
    const conjunto = new Set(mover), restantes = Array.from({ length: total }, (_, i) => i + 1).filter(n => !conjunto.has(n));
    const posicao = divisoria === undefined ? destino - 1 : divisoria - mover.filter(n => n <= divisoria).length;
    if (!Number.isInteger(posicao) || posicao < 0 || posicao > restantes.length) throw new Error(`Informe uma posição entre 1 e ${restantes.length + 1}.`);
    restantes.splice(posicao, 0, ...mover); return restantes;
}
