const mensagens = {
    PDF_SENHA: 'Senha de abertura necessária.',
    PDF_RESTRITO: 'PDF com restrições de operações.',
    PDF_CRIPTOGRAFADO: 'PDF criptografado. A combinação não está disponível para este arquivo.',
    PDF_INVALIDO: 'PDF inválido ou incompatível com a leitura.'
};

async function aguardar(promessa, signal) {
    let cancelar;
    const interrupcao = new Promise((_, reject) => {
        cancelar = () => reject(signal.reason);
        if (signal.aborted) cancelar();
        else signal.addEventListener('abort', cancelar, { once: true });
    });
    try {
        return await Promise.race([promessa, interrupcao]);
    } finally { signal.removeEventListener('abort', cancelar); }
}

// A conferência mantém uma cópia da seleção até que o operador decida combinar.
export class PreparacaoConferencia {
    constructor({ list, count, PDFCombinacao, PDFPreview, wordReports, changed = () => {}, removido = () => {}, prazo = 60_000 }) {
        Object.assign(this, { list, count, PDFCombinacao, PDFPreview, wordReports, changed, removido, prazo });
        this.batch = null;
    }
    copiar(draft) {
        return Object.assign(new this.PDFCombinacao(draft.executar), {
            sources: [...draft.sources], pages: [...draft.pages], serial: draft.serial, size: draft.size
        });
    }
    begin(files, previous, cancel) {
        if (this.batch) throw Error('Conclua ou cancele a preparação anterior.');
        const batch = { draft: this.copiar(previous), cancel, items: [], abort: new AbortController() };
        this.batch = batch;
        this.list.hidden = false;
        this.list.replaceChildren();
        this.count.hidden = false;
        for (const file of files) {
            const row = document.createElement('li'); row.className = 'pdf-import-item';
            const thumb = document.createElement('div'); thumb.className = 'pdf-combinar-miniatura pdf-import-loading'; thumb.setAttribute('aria-label', 'Aguardando preparação');
            const label = document.createElement('span'); label.textContent = `${file.name} · ${file.type}`;
            const details = document.createElement('div');
            const state = document.createElement('span'); state.className = 'pdf-import-state'; state.textContent = 'Aguardando';
            const remove = document.createElement('button'); remove.type = 'button'; remove.hidden = true; remove.textContent = 'Remover da seleção';
            remove.setAttribute('aria-label', `Remover ${file.name} da seleção`);
            const item = { row, thumb, state, remove, ready: false };
            remove.addEventListener('click', () => {
                if (!this.current(batch) || !batch.finished || !item.error || item.excluded) return;
                item.excluded = true;
                row.hidden = true;
                this.update();
                this.removido();
            });
            details.append(state, remove); row.append(thumb, label, details);
            this.list.append(row); batch.items.push(item);
        }
        this.update();
        return batch;
    }
    current(batch) { return !!batch && this.batch === batch && !batch.abort.signal.aborted; }
    update() {
        const b = this.batch;
        if (!b) return;
        const ativos = b.items.filter(i => !i.excluded);
        const falhas = ativos.filter(i => i.error).length;
        const resumo = ativos.length ? `${ativos.filter(i => i.ready).length} de ${ativos.length} arquivo(s) pronto(s).` : 'Nenhum arquivo na seleção.';
        const aviso = b.failed ? ' Falha na seleção; nenhum arquivo foi importado.' : falhas ? ` Falha em ${falhas} arquivo(s).` : '';
        this.count.textContent = resumo + aviso;
        this.changed();
    }
    stage(index, state) {
        const b = this.batch, item = b?.items[index];
        if (!item || !this.current(b) || item.ready || item.error || item.processing) return;
        item.state.textContent = state === 'word' ? 'Convertendo Word' : state === 'failed' ? 'Falha' : state === 'pdf' ? 'Lendo PDF' : 'Lendo imagem';
        this.update();
    }
    async accept(index, file, batch = this.batch) {
        if (!this.current(batch)) return false;
        const item = batch.items[index];
        if (!item || item.ready || item.error || item.processing) throw Error('Sequência de preparação inválida.');
        item.processing = true;
        const draft = this.copiar(batch.draft);
        const host = document.createElement('div');
        const preview = new this.PDFPreview(host);
        preview.calcularEscala = base => Math.min(100 / base.width, 115 / base.height);
        const timeout = new AbortController();
        const signal = AbortSignal.any([batch.abort.signal, timeout.signal]);
        const timer = setTimeout(() => timeout.abort(new DOMException('Prazo de preparação excedido.', 'TimeoutError')), this.prazo);
        try {
            item.state.textContent = /\.pdf$/i.test(file.name) ? 'Lendo PDF' : 'Preparando imagem';
            this.update();
            try {
                await aguardar(draft.add([file], { signal }), signal);
            } catch (error) {
                if (error.codigo !== 'PDF_CRIPTOGRAFADO') throw error;
                const bytes = await aguardar(file.arrayBuffer(), signal);
                const codigo = await aguardar(preview.verificarProtecao(bytes, signal), signal);
                throw Object.assign(new Error(mensagens[codigo]), { codigo });
            }
            const source = draft.sources.at(-1);
            item.state.textContent = 'Gerando miniaturas';
            this.update();
            const candidate = await aguardar(preview.preparar(source.previewBytes, { signal }), signal);
            if (!candidate || !await aguardar(preview.adotar(candidate), signal) || !await aguardar(preview.mostrar(1), signal)) {
                throw Error('Não foi possível validar a miniatura.');
            }
            signal.throwIfAborted();
            const canvas = host.querySelector('canvas');
            if (!canvas?.width) throw Error('Miniatura indisponível.');
            const img = document.createElement('img'); img.src = canvas.toDataURL('image/png'); img.alt = 'Primeira página validada';
            item.thumb.replaceChildren(img);
            item.thumb.setAttribute('aria-label', 'Miniatura pronta');
            if (file.wordReport) this.wordReports.set(source, file.wordReport);
            batch.draft = draft;
            item.ready = true;
            item.state.textContent = 'Pronto';
            return true;
        } catch (error) {
            if (this.current(batch)) {
                error = timeout.signal.aborted ? timeout.signal.reason : error;
                item.error = error;
                item.row.dataset.falha = 'true';
                item.thumb.setAttribute('aria-label', 'Miniatura indisponível');
                const codigo = error.name === 'InvalidPDFException' ? 'PDF_INVALIDO' : error.codigo;
                item.state.textContent = mensagens[codigo] || (error.name === 'TimeoutError'
                    ? 'A preparação excedeu o prazo. Remova este arquivo para continuar.' : 'Falha na preparação. ' + error.message);
            }
            return false;
        } finally {
            try { await aguardar(preview.destruir(), signal); } catch {}
            clearTimeout(timer);
            item.processing = false;
            item.thumb.classList.remove('pdf-import-loading');
            if (this.current(batch)) this.update();
        }
    }
    async complete() {
        const b = this.batch;
        if (!this.current(b) || !b.items.length || b.items.some(i => !i.ready && !i.error)) throw Error('A seleção ainda não está pronta.');
        b.finished = true;
        const falhas = b.items.filter(i => i.error && !i.excluded);
        for (const item of falhas) item.remove.hidden = false;
        if (falhas.length || !b.items.some(i => i.ready)) { this.update(); return null; }
        this.reset();
        return b.draft;
    }
    fail(index) {
        const b = this.batch;
        if (!this.current(b)) return;
        b.failed = true; b.draft = null; b.abort.abort();
        for (const [i, item] of b.items.entries()) {
            item.remove.hidden = true;
            item.thumb.classList.remove('pdf-import-loading');
            if (i === index || !item.ready) item.state.textContent = i === index ? 'Falha' : 'Não importado';
        }
        this.update();
    }
    reset() {
        this.batch = null;
        this.list.hidden = true;
        this.list.replaceChildren();
        this.count.hidden = true;
        this.changed();
    }
    cancel() {
        const b = this.batch;
        if (!b) return;
        b.abort.abort(); b.draft = null;
        this.reset();
        b.cancel?.();
    }
}
