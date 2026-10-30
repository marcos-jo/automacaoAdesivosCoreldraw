import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import { ORIGIN, START_URL, CSP, assetPath, isAppURL, validSender } from './security.mjs';
import { savePDF, openPDFs } from './files.mjs';
import { updateAction } from './update-ui.mjs';

// Dependências injetáveis permitem testar diálogos sem acrescentar backdoors ao executável.
export async function startDesktop({ app, BrowserWindow, Menu, session, protocol, ipcMain, dialog, nativeImage }, { profile } = {}) {
    app.setName('4PrintShops');
    app.setPath('userData', profile || path.join(app.getPath('appData'), '4PrintShops'));
    if (!app.requestSingleInstanceLock()) { app.quit(); return null; }
    protocol.registerSchemesAsPrivileged([{ scheme: 'app', privileges: { standard: true, secure: true, supportFetchAPI: true, corsEnabled: true, stream: true } }]);
    let window, saving = false, picking = false;
    const documentState = { pending: false, busy: false };
    app.on('second-instance', () => { if (window) { if (window.isMinimized()) window.restore(); window.show(); window.focus(); } });
    await app.whenReady();
    if (process.platform === 'win32') app.setAppUserModelId('br.com.4printshops.desktop');
    const root = fileURLToPath(new URL('../public/', import.meta.url));
    const mime = { '.html':'text/html; charset=utf-8', '.js':'text/javascript; charset=utf-8', '.mjs':'text/javascript; charset=utf-8', '.css':'text/css; charset=utf-8', '.png':'image/png', '.wasm':'application/wasm', '.bcmap':'application/octet-stream', '.pfb':'application/octet-stream', '.ttf':'font/ttf' };
    protocol.handle('app', async request => {
        const file = assetPath(root, request.url);
        if (!file || !['GET','HEAD'].includes(request.method)) return new Response('Não encontrado', { status: 404 });
        try {
            const data = await readFile(file);
            return new Response(request.method === 'HEAD' ? null : data, { headers: { 'Content-Type': mime[path.extname(file)] || 'application/octet-stream',
                'Content-Security-Policy': CSP, 'X-Content-Type-Options':'nosniff', 'Cache-Control':'no-store' } });
        } catch { return new Response('Não encontrado', { status: 404 }); }
    });
    const ses = session.defaultSession;
    ses.setPermissionRequestHandler((_contents, _permission, callback) => callback(false));
    ses.setPermissionCheckHandler(() => false);
    ses.webRequest.onBeforeRequest((details, callback) => callback({ cancel: !(isAppURL(details.url) || details.url.startsWith('blob:' + ORIGIN + '/') || details.url.startsWith('data:')) }));
    Menu.setApplicationMenu(null);
    window = new BrowserWindow({ title: '4PrintShops', width: 1440, height: 1000, minWidth: 820, minHeight: 580, show: false,
        icon: nativeImage.createFromPath(path.join(root, 'img/logo.png')),
        webPreferences: { preload: fileURLToPath(new URL('./preload.cjs', import.meta.url)), contextIsolation: true,
            nodeIntegration: false, sandbox: true, webSecurity: true, allowRunningInsecureContent: false, webviewTag: false, spellcheck: false } });
    window.removeMenu();
    window.webContents.setWindowOpenHandler(() => ({ action: 'deny' }));
    window.webContents.on('will-navigate', (event, url) => { if (url !== START_URL) event.preventDefault(); });
    window.webContents.on('will-redirect', event => event.preventDefault());
    window.webContents.on('will-attach-webview', event => event.preventDefault());
    window.on('close', event => {
        if (saving || picking || documentState.busy) { event.preventDefault(); dialog.showMessageBoxSync(window, { type:'info', title:'Operação em andamento', message:'Aguarde a conclusão da operação antes de fechar o 4PrintShops.', buttons:['Continuar'], defaultId:0 }); }
    });
    window.webContents.on('will-prevent-unload', event => {
        const answer = dialog.showMessageBoxSync(window, { type:'warning', title:'Alterações não exportadas',
            message:'Você possui alterações no documento atual que ainda não foram exportadas. Se sair agora, essas alterações serão perdidas.',
            buttons:['Continuar editando','Descartar e sair'], defaultId:0, cancelId:0, noLink:true });
        if (answer === 1) event.preventDefault(); // Electron: preventDefault aqui autoriza sair.
    });
    ipcMain.on('document:state', (event, state) => {
        if (!validSender(event, window) || !state || typeof state.pending !== 'boolean' || typeof state.busy !== 'boolean') return;
        documentState.pending = state.pending; documentState.busy = state.busy;
    });
    ipcMain.handle('pdf:save', async (event, bytes, name) => {
        if (!validSender(event, window) || saving) throw new Error('Solicitação de gravação recusada.');
        saving = true;
        try { return await savePDF(window, dialog, bytes, name, app.getPath('downloads')); }
        finally { saving = false; }
    });
    ipcMain.handle('pdf:open', async (event, multiple) => {
        if (!validSender(event, window) || saving || picking || documentState.busy) throw new Error('Solicitação de abertura recusada.');
        picking = true;
        try { return await openPDFs(window, dialog, multiple); }
        finally { picking = false; }
    });
    const checkUpdates=updateAction(app,{getWindow:()=>window,dialog,canInstall:()=>!!window&&!documentState.pending&&!documentState.busy&&!saving&&!picking});
    ipcMain.handle('updates:check',event=>{
        if(!validSender(event,window))throw new Error('Solicitação de atualização recusada.');
        return checkUpdates();
    });
    ses.on('will-download', (_event, item) => {
        if (!item.getFilename().toLowerCase().endsWith('.pdf')) { item.cancel(); return; }
        item.setSaveDialogOptions({ title:'Salvar páginas PDF', defaultPath:item.getFilename(), filters:[{name:'PDF',extensions:['pdf']}] });
    });
    window.once('ready-to-show', () => window.show());
    window.on('closed', () => { window = null; });
    app.on('window-all-closed', () => app.quit());
    app.on('will-quit', () => { ipcMain.removeHandler('pdf:save'); ipcMain.removeHandler('pdf:open'); ipcMain.removeHandler('updates:check'); protocol.unhandle('app'); });
    await window.loadURL(START_URL);
    return { get window() { return window; }, documentState, get busy() { return saving || documentState.busy; } };
}
