export async function salvarArquivoPDF(bytes, nome, { signal } = {}) {
    const interromper = () => { if (signal?.aborted) throw new DOMException('Exportação cancelada', 'AbortError'); };
    interromper();
    const desktop = globalThis.window?.desktop4PrintShops;
    if (desktop) {
        const resultado = await desktop.salvarPDF(bytes, nome);
        interromper();
        if (!resultado?.saved) throw new DOMException('Exportação cancelada', 'AbortError');
        return;
    }
    const url = URL.createObjectURL(new Blob([bytes], { type: 'application/pdf' }));
    const link = document.createElement('a'); link.href = url; link.download = nome;
    document.body.appendChild(link);
    try { link.click(); } finally { link.remove(); setTimeout(() => URL.revokeObjectURL(url), 60000); }
}
