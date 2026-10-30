import path from 'node:path';

export const ORIGIN = 'app://4printshops';
export const START_URL = ORIGIN + '/index.html';
export const CSP = "default-src 'self'; script-src 'self' 'wasm-unsafe-eval'; style-src 'self' 'unsafe-inline'; img-src 'self' blob: data:; font-src 'self' data:; connect-src 'self' blob: data:; worker-src 'self' blob:; object-src 'none'; frame-src 'none'; base-uri 'none'; form-action 'none'";
export function isAppURL(value) {
    try { const u = new URL(value); return u.protocol === 'app:' && u.hostname === '4printshops' && !u.port && !u.username && !u.password; }
    catch { return false; }
}
export function assetPath(root, value) {
    if (!isAppURL(value)) return null;
    let pathname;
    try { pathname = decodeURIComponent(new URL(value).pathname); } catch { return null; }
    if (pathname === '/') pathname = '/index.html';
    if (pathname.includes('\\') || pathname.includes('\0') || pathname.split('/').some(p => p === '..' || p.startsWith('.'))) return null;
    const allowed = pathname === '/index.html' || /^\/css\/[a-z0-9-]+\.css$/.test(pathname) || pathname === '/img/logo.png'
        || /^\/js\/main\.js$/.test(pathname) || /^\/js\/app\/(shell|tabs|menus|export-dialog|save-file|desktop-client)\.js$/.test(pathname)
        || /^\/js\/tools\/[a-z0-9-]+\.js$/.test(pathname) || /^\/vendor\/(pdfjs|pdf-lib)\/[a-zA-Z0-9._/-]+$/.test(pathname);
    if (!allowed) return null;
    const target = path.resolve(root, '.' + pathname), relative = path.relative(root, target);
    return relative && !relative.startsWith('..') && !path.isAbsolute(relative) ? target : null;
}
export function validSender(event, window) {
    return !!window && !window.isDestroyed() && event.sender === window.webContents
        && event.senderFrame === window.webContents.mainFrame && isAppURL(event.senderFrame.url);
}
export function pdfName(value) {
    if (typeof value !== 'string' || !value.trim() || value.length > 180 || /[<>:"/\\|?*\x00-\x1f]/.test(value) || /[. ]$/.test(value)) throw new Error('Nome PDF inválido.');
    if (!/\.pdf$/i.test(value) || /^(con|prn|aux|nul|com[1-9]|lpt[1-9])(?:\.|$)/i.test(value)) throw new Error('Nome PDF inválido.');
    return value;
}
export function pdfData(value) {
    if (!(value instanceof Uint8Array) || value.length < 8 || value.length > 512 * 1024 * 1024 || Buffer.from(value.subarray(0, 5)).toString('ascii') !== '%PDF-') throw new Error('Dados PDF inválidos ou maiores que 512 MiB.');
    return value;
}
