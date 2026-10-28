export class Elemento {
    constructor(tag) {
        this.tag = tag;
        this.children = [];
        this.listeners = new Map();
        this.dataset = {};
        this.style = {};
        this.value = "";
        this.textContent = "";
        this.classList = { toggle: (classe, ativo) => { this[classe] = ativo; } };
    }
    append(...filhos) { for (const filho of filhos) { filho.parent = this; this.children.push(filho); } }
    appendChild(filho) { this.append(filho); }
    replaceChildren(...filhos) { this.children = []; this.append(...filhos); }
    before(filho) { filho.parent = this.parent; this.parent.children.splice(this.parent.children.indexOf(this), 0, filho); }
    remove() { this.parent.children = this.parent.children.filter(filho => filho !== this); }
    set innerHTML(html) {
        this.replaceChildren();
        const pilha = [this];
        for (const [, fecha, tag, atributos] of html.matchAll(/<(\/)?([\w-]+)([^>]*)>/g)) {
            if (fecha) { pilha.pop(); continue; }
            const filho = new Elemento(tag);
            for (const [, nome, valor] of atributos.matchAll(/([\w-]+)="([^"]*)"/g)) filho[nome] = valor;
            filho.disabled = /\bdisabled\b/.test(atributos);
            filho.hidden = /\bhidden\b/.test(atributos);
            pilha.at(-1).append(filho);
            if (!["input", "br", "hr"].includes(tag)) pilha.push(filho);
        }
    }
    matches(seletor) {
        if (seletor === "[id]") return !!this.id;
        if (seletor.startsWith("#")) return this.id === seletor.slice(1);
        const dado = seletor.match(/^button\[data-(\w+)\]$/);
        if (dado) return this.tag === "button" && this.dataset[dado[1]] !== undefined;
        return this.tag === seletor;
    }
    querySelectorAll(seletor) {
        return this.children.flatMap(filho => [
            ...(filho.matches(seletor) ? [filho] : []), ...filho.querySelectorAll(seletor)
        ]);
    }
    querySelector(seletor) { return this.querySelectorAll(seletor)[0]; }
    closest(seletor) { return this.matches(seletor) ? this : this.parent?.closest(seletor); }
    setAttribute(nome, valor) { this[nome] = valor; }
    getAttribute(nome) { return this[nome]; }
    addEventListener(tipo, handler) {
        if (!this.listeners.has(tipo)) this.listeners.set(tipo, []);
        this.listeners.get(tipo).push(handler);
    }
    removeEventListener(tipo, handler) {
        this.listeners.set(tipo, (this.listeners.get(tipo) || []).filter(atual => atual !== handler));
    }
    emitir(tipo, alvo = this) {
        for (let atual = this; atual; atual = atual.parent) {
            for (const handler of atual.listeners.get(tipo) || []) handler({ target: alvo, preventDefault() {} });
        }
    }
    click() {
        if (this.disabled) return;
        if (this.tag === "a") {
            globalThis.__pdfIntegration.downloads.push({ nome: this.download, blob: this.hrefBlob });
        } else this.emitir("click");
    }
    set href(valor) { this.hrefBlob = globalThis.__pdfIntegration.blobs.get(valor); }
    showModal() { this.open = true; }
    close() { this.open = false; }
    focus() { this.focado = true; }
    getContext() { return {}; }
}

export class JanelaTeste extends EventTarget {
    constructor() { super(); this.listeners = new Map(); }
    addEventListener(tipo, handler, opcoes) {
        if (!this.listeners.has(tipo)) this.listeners.set(tipo, new Set());
        this.listeners.get(tipo).add(handler);
        super.addEventListener(tipo, handler, opcoes);
    }
    removeEventListener(tipo, handler, opcoes) {
        this.listeners.get(tipo)?.delete(handler);
        super.removeEventListener(tipo, handler, opcoes);
    }
}
