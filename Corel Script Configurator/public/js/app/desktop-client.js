// Único contato opcional do frontend com a ponte restrita do Desktop.
export function informarDocumento({ pending, busy }) {
    globalThis.window?.desktop4PrintShops?.estadoDocumento({ pending: !!pending, busy: !!busy });
}

export function instalarImportacaoDesktop(raiz) {
    const desktop = globalThis.window?.desktop4PrintShops;
    if (!desktop?.abrirPDF) return () => {};
    let encerrada = false, escolhendo = false;
    const selecionar = async evento => {
        const input = evento.target;
        if (input.tagName !== 'INPUT' || input.type !== 'file' || !input.accept?.includes('pdf')) return;
        evento.preventDefault();
        if (input.disabled || escolhendo) return;
        escolhendo = true;
        try {
            const arquivos = await desktop.abrirPDF(input.multiple, input.accept.includes('image/'));
            if (encerrada || !arquivos.length) return;
            const lista = new DataTransfer();
            for (const arquivo of arquivos) {
                const type = /\.png$/i.test(arquivo.name) ? 'image/png' : /\.jpe?g$/i.test(arquivo.name) ? 'image/jpeg' : 'application/pdf';
                lista.items.add(new File([arquivo.bytes], arquivo.name, { type }));
            }
            input.files = lista.files; input.dispatchEvent(new Event('change', { bubbles:true }));
        } catch (error) {
            if (!encerrada) raiz.querySelector('#pdf-status').textContent = 'Não foi possível importar: ' + error.message;
        } finally { escolhendo = false; }
    };
    raiz.addEventListener('click', selecionar);
    return () => { encerrada = true; raiz.removeEventListener('click', selecionar); };
}
