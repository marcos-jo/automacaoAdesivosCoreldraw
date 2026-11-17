export const CHAVE_FORMATOS = '4printshops.pdf.formatos.v1';

async function sha256(texto) {
    const hash = await globalThis.crypto.subtle.digest('SHA-256', new TextEncoder().encode(texto));
    return Array.from(new Uint8Array(hash), b => b.toString(16).padStart(2, '0')).join('');
}
// Envelope verificável: texto original exato, lista efetiva e metadados completos.
export async function criarBackupFormatos(original, documento) {
    const conteudo = JSON.stringify({ chave: CHAVE_FORMATOS, original, documento });
    return { tipo: '4printshops.formatos.backup', versao: 1, conteudo, sha256: await sha256(conteudo) };
}
export async function lerBackupFormatos(backup) {
    const envelope = JSON.parse(typeof backup === 'string' ? backup : JSON.stringify(backup));
    if (envelope?.tipo !== '4printshops.formatos.backup' || envelope.versao !== 1 || typeof envelope.conteudo !== 'string' || envelope.sha256 !== await sha256(envelope.conteudo)) throw new Error('Backup inválido ou SHA-256 divergente. Nenhum dado foi alterado.');
    const dados = JSON.parse(envelope.conteudo);
    if (dados.chave !== CHAVE_FORMATOS || (dados.original !== null && typeof dados.original !== 'string') || dados.documento?.versao !== 1 || !Array.isArray(dados.documento.formatos)) throw new Error('Backup de catálogo inválido.');
    return dados;
}
