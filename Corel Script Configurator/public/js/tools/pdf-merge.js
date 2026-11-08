import { PDFCombinacao } from "./pdf-combine-session.js";
import { PDFPreview } from "./pdf-preview.js";

// Janela única de combinação. O rascunho não altera o documento em edição.
export function renderUnirPDFs(app, adicionarAoDocumento, temDocumento = () => false) {
    const container = document.createElement('div');
    container.className = 'pdf-uniao-container';
    container.innerHTML = `
        <button id="pdf-uniao-abrir" class="pdf-uniao-abrir" type="button">Combinar arquivos</button>
        <dialog id="pdf-uniao-janela" class="pdf-uniao" aria-labelledby="pdf-uniao-titulo">
            <header class="pdf-uniao-cabecalho"><h3 id="pdf-uniao-titulo">Conferir e combinar arquivos</h3><button id="pdf-uniao-fechar" type="button" aria-label="Fechar janela">×</button></header>
            <div class="pdf-uniao-conteudo">
                <label class="pdf-uniao-adicionar" for="pdf-uniao-arquivos">+ Adicionar mais arquivos</label>
                <input id="pdf-uniao-arquivos" type="file" accept=".pdf,.jpg,.jpeg,.png,.docx,.doc,application/pdf,image/jpeg,image/png,application/msword,application/vnd.openxmlformats-officedocument.wordprocessingml.document" multiple hidden>
                <details id="pdf-uniao-word" hidden><summary>Documentos Word convertidos — confira fontes e paginação</summary><p id="pdf-uniao-word-info"></p></details>
                <p class="pdf-combinar-ajuda">Arraste páginas ou use ↑/↓. Remova uma página ou todo o arquivo antes de combinar.</p>
                <ol id="pdf-uniao-lista" aria-label="Páginas na ordem de combinação"></ol>
                <div class="pdf-combinar-config">
                    <label>Destino<select id="pdf-uniao-destino"><option value="adicionar">Acrescentar ao PDF aberto</option><option value="novo">Criar novo PDF combinado</option></select></label>
                    <label>Páginas finais<select id="pdf-uniao-ajuste"><option value="original">Manter tamanho original</option><option value="ajustar">Caber em A4 proporcionalmente</option><option value="preencher">Preencher A4 com possível corte</option></select></label>
                    <label>Nome do novo documento<input id="pdf-uniao-nome" value="arquivos_unidos"></label>
                </div>
                <label id="pdf-uniao-corte-label" hidden><input id="pdf-uniao-corte" type="checkbox"> Aceito o possível corte das bordas ao preencher A4.</label>
                <details><summary>Dimensões, cores e recursos dos arquivos</summary><p>Imagens recebem dimensão inicial a 150 ppi, sem confiar no DPI declarado. EXIF é respeitado. A4 acompanha a orientação de cada página. Ao acrescentar, a opção A4 inclui as páginas já abertas.</p><p>JPEG RGB/cinza e PNG com transparência são suportados. CMYK, perfis ICC, intenções de saída PDF, camadas opcionais e codificações ainda não validadas são bloqueados com explicação. Cor de impressão não verificada. PDFs preservam conteúdo vetorial; marcadores e estrutura de navegação podem não acompanhar a combinação. Formulários exigem cópia para impressão conferida. Assinaturas conservam o diagnóstico da origem; a reescrita pode invalidar sua validade criptográfica.</p></details>
                <p id="pdf-uniao-status" role="status" aria-live="polite"></p>
            </div>
            <footer class="pdf-combinar-footer"><span id="pdf-uniao-total"></span><button id="pdf-uniao-limpar" type="button">Limpar lista</button><button id="pdf-uniao-gerar" type="button">Combinar</button></footer>
        </dialog>`;
    app.querySelector('#pdf-uniao-opcoes').append(container);
    const dialog = container.querySelector('#pdf-uniao-janela');
    const get = id => container.querySelector('#pdf-uniao-' + id) || dialog.querySelector('#pdf-uniao-' + id);
    // O diálogo também abre pela importação superior, com o painel direito recolhido.
    app.querySelector('#pdf-ferramenta').append(dialog);
    const list = get('lista'), status = get('status');
    get('ajuste').value = 'original';
    let draft = new PDFCombinacao(), busy = false, occupied = false, destroyed = false;
    let generation = 0, dragId = null, observer, preview, source, draining = false;
    const queue = [], cache = new Map(), wordReports = new WeakMap();
    const host = document.createElement('div');
    host.hidden = true; container.append(host);

    function controls() {
        for (const id of ['arquivos','nome','destino','ajuste','corte','limpar','fechar']) get(id).disabled = busy || occupied;
        get('gerar').disabled = busy || occupied || !draft.pages.length || (get('ajuste').value === 'preencher' && !get('corte').checked);
        for (const button of list.querySelectorAll('button')) button.disabled = busy || occupied;
        get('corte-label').hidden = get('ajuste').value !== 'preencher';
        get('total').textContent = `${draft.sources.length} arquivo(s) · ${draft.pages.length} página(s)`;
    }
    async function release() {
        generation++; observer?.disconnect(); queue.length = 0; cache.clear();
        const old = preview; preview = null; source = null;
        await old?.destruir();
    }
    async function drain() {
        if (draining) return;
        draining = true;
        try {
            while (queue.length && dialog.open && !destroyed) {
                const { page, target, version } = queue.shift();
                if (version !== generation || !target.isConnected || !target.dataset.visible) continue;
                try {
                    let image = cache.get(page.id);
                    if (!image) {
                        preview ||= new PDFPreview(host);
                        preview.calcularEscala = base => Math.min(100 / base.width, 115 / base.height);
                        const current = preview;
                        if (source !== page.source) {
                            const candidate = await current.preparar(page.source.previewBytes);
                            if (!candidate || version !== generation) continue;
                            if (!await current.adotar(candidate)) continue;
                            source = page.source;
                        }
                        if (!await current.mostrar(page.index + 1) || version !== generation) continue;
                        image = host.querySelector('canvas').toDataURL('image/png');
                        cache.set(page.id, image);
                        while (cache.size > 48) cache.delete(cache.keys().next().value);
                    }
                    if (version !== generation || !target.dataset.visible) continue;
                    const img = document.createElement('img'); img.src = image; img.alt = 'Miniatura da página ' + (page.index + 1);
                    target.replaceChildren(img);
                } catch (error) {
                    if (version !== generation || !target.isConnected) continue;
                    const retry = document.createElement('button'); retry.type = 'button'; retry.textContent = 'Falha · tentar novamente';
                    retry.title = error.message;
                    retry.addEventListener('click', () => { target.textContent = 'Carregando…'; queue.push({ page, target, version }); drain(); });
                    target.replaceChildren(retry);
                }
            }
        } finally { draining = false; }
    }
    function render() {
        const reports=draft.sources.map(source=>wordReports.get(source)).filter(Boolean);
        get('word').hidden=!reports.length;
        get('word-info').textContent=reports.map(report=>`${report.originalName}: PDF com ${report.pages} página(s); origem ${report.sourcePages ?? 'não verificada'} (metadados podem estar desatualizados). Fontes declaradas: ${report.requestedFonts.join(', ')||'não verificadas'}. Fontes no PDF: ${report.fonts.map(f=>f.name+(f.embedded?' (incorporada)':' (incorporação não confirmada)')).join(', ')||'não identificadas'}.`).join(' ')+(reports.length?' A conversão local pode alterar fontes, quebras, tabelas e paginação. Confira as miniaturas e compare com o original antes de combinar. Substituição de fontes e fidelidade ao Word não verificadas automaticamente. Originais preservados.':'');
        observer?.disconnect(); queue.length = 0; list.replaceChildren();
        const version = generation;
        observer = new IntersectionObserver(entries => {
            for (const entry of entries) {
                const target = entry.target, page = draft.pages.find(p => p.id === Number(target.dataset.id));
                if (!page) continue;
                if (entry.isIntersecting) {
                    target.dataset.visible = 'true';
                    queue.push({ page, target, version });
                } else { delete target.dataset.visible; target.textContent = 'Miniatura'; }
            }
            drain();
        }, { root: list, rootMargin: '120px' });
        draft.pages.forEach((page, index) => {
            const row = document.createElement('li'); row.className = 'pdf-uniao-item'; row.draggable = true; row.dataset.id = page.id;
            const thumb = document.createElement('div'); thumb.className = 'pdf-combinar-miniatura'; thumb.dataset.id = page.id; thumb.textContent = 'Miniatura';
            const label = document.createElement('span'); label.className = 'pdf-uniao-nome';
            const mm = value => value.toLocaleString('pt-BR', { maximumFractionDigits: 2 });
            label.textContent = `${index + 1}. ${page.source.name} · ${page.source.type} · página ${page.index + 1} de ${page.source.total} · ${mm(page.width)} × ${mm(page.height)} mm`;
            const actions = document.createElement('div'); actions.className = 'pdf-uniao-acoes';
            for (const [action, text, title] of [['subir','↑','Mover página para cima'],['descer','↓','Mover página para baixo'],['remover','×','Remover página'],['arquivo','Remover arquivo','Remover todas as páginas deste arquivo']]) {
                const button = document.createElement('button'); button.type = 'button'; button.dataset.acao = action; button.dataset.indice = index; button.textContent = text; button.title = title; button.setAttribute('aria-label', title); actions.append(button);
            }
            row.append(thumb, label, actions); list.append(row); observer.observe(thumb);
        });
        controls();
    }
    async function add(files) {
        if (busy || occupied || destroyed || !files.length) return;
        busy = true; controls(); status.textContent = 'Analisando arquivos…';
        try {
            if(files.some(file=>/\.docx?$/i.test(file.name)))throw Error('Para converter Word localmente, use o aplicativo Desktop.');
            const previous=draft.sources.length;
            await draft.add(files);
            draft.sources.slice(previous).forEach((source,index)=>{if(files[index].wordReport)wordReports.set(source,files[index].wordReport);});
            if (!destroyed) { render(); status.textContent = 'Arquivos preparados. Confira a sequência, as páginas e as dimensões antes de combinar.'; }
        }
        catch (error) { if (!destroyed) status.textContent = error.message; }
        finally { busy = false; if (!destroyed) controls(); }
    }
    async function open(files = []) {
        if (destroyed || occupied) return;
        if (!dialog.open) dialog.showModal();
        get('destino').querySelector('option').disabled = !temDocumento();
        get('destino').value = temDocumento() ? 'adicionar' : 'novo';
        render(); await add(files);
    }
    get('abrir').addEventListener('click', () => open());
    get('fechar').addEventListener('click', () => dialog.close());
    dialog.addEventListener('cancel', event => { if (busy || occupied) event.preventDefault(); });
    dialog.addEventListener('close', () => { if (!dialog.open) release(); });
    get('arquivos').addEventListener('change', event => { const files = event.detail?.files || Array.from(get('arquivos').files); get('arquivos').value = ''; add(files); });
    get('ajuste').addEventListener('change', controls); get('corte').addEventListener('change', controls);
    get('limpar').addEventListener('click', async () => { if (busy || occupied) return; await release(); draft = new PDFCombinacao(); render(); status.textContent = 'Lista limpa.'; });
    list.addEventListener('click', event => {
        if (busy || occupied) return;
        const button = event.target.closest('button[data-acao]'); if (!button) return;
        const index = Number(button.dataset.indice), page = draft.pages[index]; if (!page) return;
        if (button.dataset.acao === 'remover' || button.dataset.acao === 'arquivo') draft.remove(page.id, button.dataset.acao === 'arquivo');
        else { const to = index + (button.dataset.acao === 'subir' ? -1 : 1); if (to < 0 || to >= draft.pages.length) return; draft.move(index, to); }
        render();
    });
    list.addEventListener('dragstart', event => {
        const row = event.target.closest('li[data-id]');
        if (busy || occupied || !row) { event.preventDefault(); return; }
        dragId = Number(row.dataset.id); event.dataTransfer.setData('text/plain', String(dragId)); event.dataTransfer.effectAllowed = 'move';
    });
    list.addEventListener('dragover', event => {
        if (!dragId || busy || occupied) return;
        event.preventDefault(); for (const row of list.querySelectorAll('.pdf-combinar-insercao')) row.classList.remove('pdf-combinar-insercao');
        event.target.closest('li')?.classList.add('pdf-combinar-insercao');
    });
    const endDrag = () => { dragId = null; for (const row of list.querySelectorAll('.pdf-combinar-insercao')) row.classList.remove('pdf-combinar-insercao'); };
    list.addEventListener('dragend', endDrag);
    list.addEventListener('drop', event => {
        event.preventDefault(); const target = event.target.closest('li[data-id]');
        const from = draft.pages.findIndex(p => p.id === dragId), to = draft.pages.findIndex(p => p.id === Number(target?.dataset.id));
        endDrag(); if (busy || occupied || from < 0 || to < 0) return; draft.move(from, to); render();
    });
    get('gerar').addEventListener('click', async () => {
        if (busy || occupied || !draft.pages.length) return;
        busy = true; controls(); status.textContent = 'Combinando páginas…';
        try {
            if (get('ajuste').value === 'preencher' && !get('corte').checked) throw new Error('Confirme o possível corte antes de combinar.');
            const name = (get('nome').value.trim().replace(/[\\/:*?"<>|]/g, '_').replace(/\.pdf$/i, '') || 'arquivos_unidos') + '.pdf';
            const bytes = await draft.build();
            if (destroyed) return;
            const result = await adicionarAoDocumento([new File([bytes], name, { type: 'application/pdf' })], name, { novo: get('destino').value === 'novo', ajuste: get('ajuste').value, fontes: draft.origins() });
            if (result === false || destroyed) { status.textContent = 'Combinação cancelada. Documento anterior preservado.'; return; }
            draft = new PDFCombinacao(); dialog.close(); status.textContent = 'Documento combinado. Use Salvar cópia para exportar.';
        } catch (error) { if (!destroyed) { console.error(error); status.textContent = error.message; } }
        finally { busy = false; if (!destroyed) render(); }
    });
    controls();
    return {
        abrir: open,
        definirOcupado(value) { occupied = value; if (!destroyed) controls(); },
        destruir() { destroyed = true; dialog.close(); draft = new PDFCombinacao(); release(); }
    };
}
