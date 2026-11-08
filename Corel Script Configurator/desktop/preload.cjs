const { contextBridge, ipcRenderer } = require('electron');
contextBridge.exposeInMainWorld('desktop4PrintShops', Object.freeze({
    exportarImagens: (request,bytes) => ipcRenderer.invoke('image:export',request,bytes),
    salvarPDF: (bytes, nome) => ipcRenderer.invoke('pdf:save', bytes, nome),
    importarWord: () => ipcRenderer.invoke('word:import'),
    importarArquivos: multiple => ipcRenderer.invoke('files:import',multiple === true),
    cancelarImportacao: () => ipcRenderer.invoke('files:cancel'),
    cancelarWord: () => ipcRenderer.invoke('word:cancel'),
    abrirPDF: (multiple, mixed = false) => ipcRenderer.invoke('pdf:open', multiple === true, mixed === true),
    configuracaoImpressao: request => ipcRenderer.invoke('print:config',request),
    imprimirDocumento: (request,bytes) => ipcRenderer.invoke('print:document',request,bytes),
    ensaioRetido: (request,bytes) => ipcRenderer.invoke('print:hold-test',request,bytes),
    abrirCopiaExterna: (bytes,nome) => ipcRenderer.invoke('pdf:external-copy',bytes,nome),
    arquivosWindows:()=>ipcRenderer.invoke('pdf:external-list'),
    lerArquivosWindows:ids=>ipcRenderer.invoke('pdf:external-read',ids),
    concluirArquivosWindows:ids=>ipcRenderer.invoke('pdf:external-finish',ids),
    liberarArquivosWindows:ids=>ipcRenderer.invoke('pdf:external-release',ids),
    observarArquivosWindows:callback=>{
        const handler=()=>callback();ipcRenderer.on('pdf:external-available',handler);
        return ()=>ipcRenderer.removeListener('pdf:external-available',handler);
    },
    verificarAtualizacoes: () => ipcRenderer.invoke('updates:check'),
    atualizacaoAssistida: action => ipcRenderer.invoke('assisted:action',action),
    estadoAssistida: () => ipcRenderer.invoke('assisted:state'),
    observarAssistida: callback => {
        const handler=(_event,state)=>callback(state);ipcRenderer.on('assisted:state',handler);
        return ()=>ipcRenderer.removeListener('assisted:state',handler);
    },
    estadoAtualizacao: () => ipcRenderer.invoke('updates:state'),
    observarAtualizacoes: callback => {
        const handler=(_event,state)=>callback(state);ipcRenderer.on('updates:state',handler);
        return ()=>ipcRenderer.removeListener('updates:state',handler);
    },
    consultarTrabalho: callback => {
        const handler=(_event,token)=>{let state;try{state=callback();}catch{state={pending:true,busy:true};}ipcRenderer.send('updates:work-response',token,state);};
        ipcRenderer.on('updates:work-request',handler);
        return ()=>ipcRenderer.removeListener('updates:work-request',handler);
    },
    observarBloqueio: callback => {
        const handler=(_event,locked)=>callback(locked===true);ipcRenderer.on('updates:lock',handler);
        return ()=>ipcRenderer.removeListener('updates:lock',handler);
    },
    estadoDocumento: estado => ipcRenderer.send('document:state', estado)
}));
