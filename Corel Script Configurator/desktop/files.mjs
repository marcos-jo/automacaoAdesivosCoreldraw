import { open, rename, unlink, readFile, stat } from 'node:fs/promises';
import path from 'node:path';
import { randomUUID } from 'node:crypto';
import { pdfName, pdfData } from './security.mjs';

export async function openPDFs(window, dialog, multiple = false) {
    if (typeof multiple !== 'boolean') throw new Error('Seleção de arquivos inválida.');
    const choice = await dialog.showOpenDialog(window, { title: 'Abrir PDFs',
        filters: [{ name: 'Documentos PDF', extensions: ['pdf'] }], properties: multiple ? ['openFile','multiSelections'] : ['openFile'] });
    if (choice.canceled) return [];
    const paths = choice.filePaths || [];
    if (!multiple && paths.length > 1) throw new Error('Selecione apenas um PDF.');
    const entries = []; let total = 0;
    for (const file of paths) {
        const info = await stat(file);
        if (!info.isFile() || total + info.size > 512 * 1024 * 1024) throw new Error('A seleção ultrapassa o limite de 512 MiB.');
        const bytes = await readFile(file); total += bytes.length;
        if (total > 512 * 1024 * 1024) throw new Error('A seleção ultrapassa o limite de 512 MiB.');
        // Validação de conteúdo continua no parser existente; não rejeitar PDFs
        // que ele já aceita por terem um prefixo antes do cabeçalho %PDF.
        entries.push({ name: path.basename(file), bytes });
    }
    return entries;
}

// Só recebe caminho do diálogo do processo principal, nunca do renderer.
export async function savePDF(window, dialog, bytes, name, directory) {
    pdfName(name); pdfData(bytes);
    const choice = await dialog.showSaveDialog(window, { title: 'Salvar cópia PDF', defaultPath: directory ? path.join(directory, name) : name,
        filters: [{ name: 'Documento PDF', extensions: ['pdf'] }], properties: ['showOverwriteConfirmation', 'createDirectory'] });
    if (choice.canceled || !choice.filePath) return { saved: false };
    const temp = path.join(path.dirname(choice.filePath), '.4printshops-' + randomUUID() + '.tmp');
    let handle;
    try {
        handle = await open(temp, 'wx'); await handle.writeFile(bytes); await handle.sync(); await handle.close(); handle = null;
        await rename(temp, choice.filePath);
        return { saved: true };
    } finally { await handle?.close().catch(() => {}); await unlink(temp).catch(() => {}); }
}
