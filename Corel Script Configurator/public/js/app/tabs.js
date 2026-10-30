// Uma instância por ferramenta. Alternar não desmonta nem confirma descarte.
export class AbasAplicacao {
    constructor(modulos, interfacePrincipal) {
        this.modulos = modulos; this.interface = interfacePrincipal;
        this.abas = new Map([["inicio", { id: "inicio", controle: {}, painel: interfacePrincipal.inicio }]]);
        this.ativa = "inicio"; this.fila = Promise.resolve();
        this.interface.exibir(this.abas.get("inicio"), [...this.abas.values()]);
    }
    executar(acao) {
        const tarefa = this.fila.then(acao);
        this.fila = tarefa.catch(erro => this.interface.informar(erro.message));
        return tarefa;
    }
    ativar(id) { return this.executar(() => this.abrir(id)); }
    async abrir(id) {
        if (this.ativa === id) return true;
        const anterior = this.abas.get(this.ativa);
        if (anterior?.controle.podeAlternar?.() === false) return false;
        let aba = this.abas.get(id);
        if (!aba) {
            const modulo = this.modulos.find(m => m.id === id);
            if (!modulo) throw new Error("Ferramenta não disponível.");
            const painel = this.interface.criarPainel(id);
            try { aba = { id, titulo: modulo.titulo, painel, controle: await modulo.montar(painel) || {} }; }
            catch (erro) { this.interface.removerPainel(painel); throw erro; }
            this.abas.set(id, aba);
        }
        await anterior?.controle.suspender?.();
        this.ativa = id; this.interface.exibir(aba, [...this.abas.values()]);
        await aba.controle.retomar?.(); return true;
    }
    fechar(id) {
        return this.executar(async () => {
            if (id === "inicio" || !this.abas.has(id)) return false;
            if (this.abas.get(this.ativa)?.controle.podeAlternar?.() === false) return false;
            if (!await this.abrir(id)) return false;
            const aba = this.abas.get(id);
            if (aba.controle.confirmarFechamento && !await aba.controle.confirmarFechamento()) return false;
            const ids = [...this.abas.keys()], proxima = ids[Math.max(0, ids.indexOf(id) - 1)];
            await aba.controle.destruir?.();
            this.interface.removerPainel(aba.painel); this.abas.delete(id); this.ativa = null;
            await this.abrir(proxima); return true;
        });
    }
}
