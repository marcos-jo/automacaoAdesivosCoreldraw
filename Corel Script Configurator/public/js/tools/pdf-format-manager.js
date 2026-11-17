import { obterCatalogoFormatos } from './pdf-format-catalog.js';

// Um único serviço para o cadastro geral e os cadastros rápidos das ferramentas.
export class GerenciadorTamanhos {
    constructor(container, { formatos = obterCatalogoFormatos(), confirmar = mensagem => globalThis.confirm(mensagem) } = {}) {
        this.formatos = formatos; this.confirmar = confirmar;
        this.dialog = document.createElement('dialog');
        this.dialog.id = 'pdf-tamanhos'; this.dialog.className = 'app-exportacao-dialogo';
        this.dialog.setAttribute('aria-labelledby', 'pdf-tamanhos-titulo');
        this.dialog.innerHTML = `<h2 id="pdf-tamanhos-titulo">Gerenciar tamanhos</h2>
            <p>Os tamanhos são compartilhados com Redimensionar, Step &amp; Repeat e N-up. Alterar o cadastro não modifica o PDF.</p>
            <label>Tamanhos cadastrados<select data-campo="lista"><option value="">Novo tamanho</option></select></label>
            <form><label>Nome<input data-campo="nome" maxlength="80" required autocomplete="off"></label>
            <label>Largura (mm)<input data-campo="largura" inputmode="decimal" required></label>
            <label>Altura (mm)<input data-campo="altura" inputmode="decimal" required></label>
            <p data-campo="status" role="status"></p>
            <div class="app-exportacao-acoes"><button type="button" data-campo="novo">Novo</button><button type="button" data-campo="excluir">Excluir</button><button type="submit" data-campo="salvar">Adicionar</button></div></form>
            <div class="app-exportacao-acoes"><button type="button" data-campo="fechar">Fechar</button></div>`;
        container.append(this.dialog);
        this.get = campo => this.dialog.querySelector(`[data-campo="${campo}"]`);
        this.dialog.querySelector('form').addEventListener('submit', e => { e.preventDefault(); this.executar(() => this.salvar()); });
        this.get('lista').addEventListener('change', () => this.selecionar(this.get('lista').value));
        this.get('novo').addEventListener('click', () => this.selecionar(''));
        this.get('excluir').addEventListener('click', () => this.executar(() => this.excluir()));
        this.get('fechar').addEventListener('click', () => { if (!this.ocupado) this.dialog.close(); });
        this.dialog.addEventListener('cancel', e => { if (this.ocupado) e.preventDefault(); });
        this.removerCatalogo = formatos.assinar(() => this.atualizar());
        this.atualizar();
    }
    async open() {
        await this.formatos.pronta;
        this.atualizar(); this.selecionar('');
        if (!this.dialog.open) this.dialog.showModal();
    }
    atualizar() {
        const lista = this.get('lista'), id = this.selecionado?.id || '';
        lista.replaceChildren(new Option('Novo tamanho', ''));
        for (const f of this.formatos.listar()) lista.add(new Option(`${f.nome} — ${f.largura} × ${f.altura} mm`, f.id));
        lista.value = this.formatos.listar().some(f => f.id === id) ? id : '';
        // Não sobrescrever um formulário aberto com dados recebidos de outra ferramenta.
        this.get('status').textContent = this.formatos.obterAviso();
        this.botoes();
    }
    selecionar(id) {
        const f = this.formatos.listar().find(f => f.id === id);
        this.selecionado = f ? { ...f } : null; this.get('lista').value = f?.id || '';
        for (const campo of ['nome', 'largura', 'altura']) this.get(campo).value = f?.[campo] ?? '';
        this.get('status').textContent = this.formatos.obterAviso(); this.botoes();
    }
    botoes() {
        for (const campo of ['lista', 'nome', 'largura', 'altura', 'novo', 'salvar', 'fechar']) this.get(campo).disabled = !!this.ocupado;
        this.get('excluir').disabled = !!this.ocupado || !this.selecionado;
        this.get('salvar').textContent = this.selecionado ? 'Salvar alterações' : 'Adicionar';
    }
    conferirSelecao() {
        if (!this.selecionado) return;
        const atual = this.formatos.listar().find(f => f.id === this.selecionado.id);
        if (!atual || ['nome', 'largura', 'altura'].some(k => atual[k] !== this.selecionado[k])) throw Error('Este tamanho mudou em outra ferramenta. Selecione novamente antes de salvar ou excluir.');
    }
    async salvar() {
        this.conferirSelecao();
        const dados = Object.fromEntries(['nome', 'largura', 'altura'].map(k => [k, this.get(k).value]));
        const f = await this.formatos.salvar(dados, this.selecionado?.id, this.selecionado);
        this.selecionar(f.id); this.get('status').textContent = this.formatos.obterAviso() || 'Tamanho salvo nas três ferramentas.';
    }
    async excluir() {
        this.conferirSelecao();
        const f = this.selecionado;
        if (!f || !this.confirmar(`Excluir o tamanho “${f.nome}”? Os documentos abertos serão preservados.`)) return;
        await this.formatos.excluir(f.id, f); this.selecionar('');
        this.get('status').textContent = this.formatos.obterAviso() || 'Tamanho excluído. Documentos preservados.';
    }
    async executar(acao) {
        if (this.ocupado) return;
        this.ocupado = true; this.botoes();
        try { await acao(); }
        catch (e) { this.get('status').textContent = e.message; }
        finally { this.ocupado = false; this.botoes(); }
    }
    destroy() { this.removerCatalogo(); this.dialog.remove(); }
}
