import { readFile } from 'node:fs/promises';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import { ORIGIN, START_URL, CSP, assetPath, isAppURL, validSender } from './security.mjs';
import { savePDF, openPDFs } from './files.mjs';
import { updateAction } from './update-ui.mjs';
import { workStateRequest } from './work-state.mjs';
import {ExternalPDFQueue,pdfArguments} from './external-pdfs.mjs';
import {assistedAction} from './assisted-ui.mjs';
import {PrintConfigClient,validatePrintRequest} from './print-config.mjs';
import {openExternalCopy} from './external-reader.mjs';
import {HoldTestCoordinator} from './hold-test.mjs';
import {createHash} from 'node:crypto';

// Dependências injetáveis permitem testar diálogos sem acrescentar backdoors ao executável.
export async function startDesktop({ app, BrowserWindow, Menu, session, protocol, ipcMain, dialog, nativeImage, shell }, { profile } = {}) {
    // Electron exige nome/perfil/protocolo antes do primeiro await (antes de ready).
    const pkg=JSON.parse(readFileSync(new URL('../package.json',import.meta.url),'utf8'));
    const lab=pkg.desktopUpdateLab===true&&pkg.name==='4printshops-update-lab';
    const openTest=pkg.desktopOpenTest===true&&pkg.name==='4printshops-open-test';
    const assistedLab=pkg.desktopAssistedLab===true&&pkg.name==='4printshops-assisted-lab';
    const iconTest=pkg.desktopIconTest===true&&pkg.name==='4printshops-icon-test';
    const printTest=pkg.desktopPrintConfigTest===true&&pkg.name==='4printshops-print-config-test';
    const holdTest=pkg.desktopHoldTest===true&&pkg.name==='4printshops-hold-test';
    const holdNative=pkg.desktopHoldNativeTest===true&&pkg.name==='4printshops-hold-test';
    const product=lab?'4PrintShops Update Lab':openTest?'4PrintShops PDF Test':assistedLab?'4PrintShops Assisted Lab':iconTest?'4PrintShops Icon Test':printTest?'4PrintShops Print Config Test':holdTest||holdNative?'4PrintShops Hold Test':'4PrintShops';
    app.setName(product);
    app.setPath('userData', profile || path.join(app.getPath('appData'), product));
    let window, saving = false, picking = false, updateLocked = false, externalReady=false;
    const external=new ExternalPDFQueue(()=>{if(externalReady&&window&&!window.isDestroyed())window.webContents.send('pdf:external-available');});
    const initialPaths=pdfArguments(process.argv,{packaged:app.isPackaged});
    if (!app.requestSingleInstanceLock({pdfPaths:initialPaths})) { app.quit(); return null; }
    external.add(initialPaths);
    protocol.registerSchemesAsPrivileged([{ scheme: 'app', privileges: { standard: true, secure: true, supportFetchAPI: true, corsEnabled: true, stream: true } }]);
    const documentState = { pending: false, busy: false };
    app.on('second-instance', (_event,_argv,_cwd,data) => {
        if (window) { if (window.isMinimized()) window.restore(); window.show(); window.focus(); }
        try{if(Array.isArray(data?.pdfPaths))external.add(data.pdfPaths);}catch(error){dialog.showErrorBox('Não foi possível receber PDFs',error.message);}
    });
    await app.whenReady();
    if (process.platform === 'win32') app.setAppUserModelId(lab?'br.com.4printshops.updatelab':openTest?'br.com.4printshops.opentest':assistedLab?'br.com.4printshops.assistedlab':iconTest?'br.com.4printshops.icontest':printTest?'br.com.4printshops.printconfigtest':holdTest?'br.com.4printshops.holdtest':'br.com.4printshops.desktop');
    if(holdNative)app.setAppUserModelId('br.com.4printshops.holdtest');
    const nativeFolder=holdNative?'hold-native':'native';
    const nativeRoot=app.isPackaged?path.join(process.resourcesPath,nativeFolder):fileURLToPath(new URL('../build/'+nativeFolder+'/',import.meta.url));
    const printConfig=new PrintConfigClient({helper:path.join(nativeRoot,'4PrintShops.PrintConfig.exe'),
        manifest:fileURLToPath(new URL('../build/'+nativeFolder+'/manifest.json',import.meta.url)),experimental:holdNative});
    const holdCoordinator=new HoldTestCoordinator(printConfig,dialog,()=>window);
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
    window = new BrowserWindow({ title: product, width: 1440, height: 1000, minWidth: 820, minHeight: 580, show: false,
        icon: nativeImage.createFromPath(fileURLToPath(new URL('../build/icons/app.png',import.meta.url))),
        webPreferences: { preload: fileURLToPath(new URL('./preload.cjs', import.meta.url)), contextIsolation: true,
            nodeIntegration: false, sandbox: true, webSecurity: true, allowRunningInsecureContent: false, webviewTag: false, spellcheck: false } });
    window.removeMenu();
    window.webContents.on('did-start-loading',()=>{externalReady=false;external.active=false;});
    window.webContents.setWindowOpenHandler(() => ({ action: 'deny' }));
    window.webContents.on('will-navigate', (event, url) => { if (url !== START_URL) event.preventDefault(); });
    window.webContents.on('will-redirect', event => event.preventDefault());
    window.webContents.on('will-attach-webview', event => event.preventDefault());
    window.on('close', event => {
        if (saving || picking || external.active || documentState.busy || printConfig.busy || updateLocked && checkUpdates.getState().phase!=='installing') { event.preventDefault(); dialog.showMessageBoxSync(window, { type:'info', title:'Operação em andamento', message:'Aguarde a conclusão da operação antes de fechar o 4PrintShops.', buttons:['Continuar'], defaultId:0 }); }
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
        if (!validSender(event, window) || saving || external.active || printConfig.busy || updateLocked) throw new Error('Solicitação de gravação recusada.');
        saving = true;
        try { return await savePDF(window, dialog, bytes, name, app.getPath('downloads')); }
        finally { saving = false; }
    });
    ipcMain.handle('pdf:open', async (event, multiple) => {
        if (!validSender(event, window) || saving || picking || external.active || documentState.busy || printConfig.busy || updateLocked) throw new Error('Solicitação de abertura recusada.');
        picking = true;
        try { return await openPDFs(window, dialog, multiple); }
        finally { picking = false; }
    });
    ipcMain.handle('pdf:external-copy',async(event,bytes,name)=>{
        if(!validSender(event,window)||saving||picking||external.active||printConfig.busy||updateLocked)throw new Error('Conclua a operação atual antes de abrir a cópia externa.');
        saving=true;
        try{return await openExternalCopy(window,dialog,bytes,name,app.getPath('temp'));}
        finally{saving=false;}
    });
    const work=workStateRequest(()=>window);
    ipcMain.on('updates:work-response',work.receive);
    const send=(channel,value)=>{if(window&&!window.isDestroyed())window.webContents.send(channel,value);};
    const locks=new Set();
    const setLock=(owner,value)=>{if(value)locks.add(owner);else locks.delete(owner);updateLocked=locks.size>0;send('updates:lock',updateLocked);};
    const safeWork=async()=>{const current=await work.request();return !!window&&!current.pending&&!current.busy&&!documentState.pending&&!documentState.busy&&!saving&&!picking&&!printConfig.busy&&!external.items.size&&!external.active;};
    const checkUpdates=updateAction(app,{getWindow:()=>window,dialog,
        publish:state=>send('updates:state',state),
        lock:value=>setLock('integrated',value),
        canInstall:async()=>!locks.has('assisted')&&await safeWork()});
    const assisted=assistedAction(app,pkg,{getWindow:()=>window,dialog,shell,
        publish:state=>send('assisted:state',state),lock:value=>setLock('assisted',value),
        canPrepare:async()=>!locks.has('integrated')&&checkUpdates.getState().phase!=='downloading'&&await safeWork()});
    ipcMain.handle('print:config',async(event,input)=>{
        if(!validSender(event,window)||saving||picking||external.active||documentState.busy||updateLocked||printConfig.busy)throw new Error('Aguarde a operação atual antes de configurar a impressão.');
        const request=validatePrintRequest(input),ownerWindow=window;
        const modal=request.action==='properties';
        const handle=modal?ownerWindow.getNativeWindowHandle():null;
        const owner=handle?(handle.length===8?handle.readBigUInt64LE():BigInt(handle.readUInt32LE())).toString():'0';
        if(modal)ownerWindow.setEnabled(false);
        try{const result=await printConfig.request(request,{owner});if(holdNative&&result?.session)result.experimentalHold=true;return result;}
        finally{if(modal&&!ownerWindow.isDestroyed()){ownerWindow.setEnabled(true);ownerWindow.focus();}}
    });
    ipcMain.handle('print:hold-test',async(event,input,bytes)=>{
        if(!holdNative||!validSender(event,window)||saving||picking||external.active||printConfig.busy||updateLocked)throw new Error('Ensaio retido indisponível.');
        if(input?.kind==='fixture'){
            const manifest=JSON.parse(await readFile(fileURLToPath(new URL('../build/hold-native/manifest.json',import.meta.url)),'utf8'));
            const fixture=await readFile(path.join(nativeRoot,'synthetic.pdf'));
            if(createHash('sha256').update(fixture).digest('hex')!==manifest.files['synthetic.pdf'])throw new Error('Sintético de laboratório incompatível.');
            return fixture;
        }
        if(input?.kind!=='prepare')throw new Error('Solicitação de ensaio inválida.');
        const state=validatePrintRequest({action:'summary',session:input.session,revision:input.revision});
        saving=true;
        try{return await holdCoordinator.run(state,bytes);}
        finally{saving=false;}
    });
    ipcMain.handle('pdf:external-list',event=>{
        if(!validSender(event,window))throw new Error('Consulta de arquivos recusada.');
        externalReady=true;return external.list();
    });
    ipcMain.handle('pdf:external-read',(event,ids)=>{
        if(!validSender(event,window)||!externalReady||saving||picking||documentState.busy||printConfig.busy||updateLocked)throw new Error('Aguarde a operação atual antes de abrir outro PDF.');
        return external.read(ids);
    });
    ipcMain.handle('pdf:external-finish',(event,ids)=>{
        if(!validSender(event,window))throw new Error('Conclusão de arquivos recusada.');
        external.finish(ids);
    });
    ipcMain.handle('pdf:external-release',(event,ids)=>{
        if(!validSender(event,window))throw new Error('Liberação de arquivos recusada.');
        external.release(ids);
    });
    ipcMain.handle('updates:check',event=>{
        if(!validSender(event,window)||updateLocked)throw new Error('Solicitação de atualização recusada.');
        return checkUpdates();
    });
    ipcMain.handle('updates:state',event=>{
        if(!validSender(event,window))throw new Error('Solicitação de atualização recusada.');
        return {...checkUpdates.getState(),currentVersion:app.getVersion(),laboratory:lab};
    });
    ipcMain.handle('assisted:action',(event,action)=>{
        if(!validSender(event,window)||updateLocked)throw new Error('Solicitação de atualização assistida recusada.');
        return assisted(action);
    });
    ipcMain.handle('assisted:state',event=>{
        if(!validSender(event,window))throw new Error('Consulta de atualização assistida recusada.');
        return {...assisted.getState(),enabled:!!pkg.assistedUpdates,currentVersion:app.getVersion()};
    });
    ses.on('will-download', (_event, item) => {
        if (!item.getFilename().toLowerCase().endsWith('.pdf')) { item.cancel(); return; }
        item.setSaveDialogOptions({ title:'Salvar páginas PDF', defaultPath:item.getFilename(), filters:[{name:'PDF',extensions:['pdf']}] });
    });
    window.once('ready-to-show', () => window.show());
    window.on('closed', () => { window = null; });
    app.on('window-all-closed', () => app.quit());
    app.on('will-quit', () => { printConfig.destroy();assisted.destroy();checkUpdates.destroy();work.destroy();ipcMain.removeListener('updates:work-response',work.receive);for(const name of ['pdf:save','pdf:open','pdf:external-copy','pdf:external-list','pdf:external-read','pdf:external-finish','pdf:external-release','updates:check','updates:state','assisted:action','assisted:state','print:config','print:hold-test'])ipcMain.removeHandler(name);protocol.unhandle('app'); });
    await window.loadURL(START_URL);
    // Não aguardar a rede para abrir a interface e não mostrar diálogos no startup.
    void checkUpdates.start();
    void assisted.start();
    return { get window() { return window; }, documentState, get busy() { return saving || documentState.busy; } };
}
