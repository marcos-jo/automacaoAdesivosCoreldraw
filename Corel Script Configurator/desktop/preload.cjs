const { contextBridge, ipcRenderer } = require('electron');
contextBridge.exposeInMainWorld('desktop4PrintShops', Object.freeze({
    salvarPDF: (bytes, nome) => ipcRenderer.invoke('pdf:save', bytes, nome),
    abrirPDF: multiple => ipcRenderer.invoke('pdf:open', multiple === true),
    verificarAtualizacoes: () => ipcRenderer.invoke('updates:check'),
    estadoDocumento: estado => ipcRenderer.send('document:state', estado)
}));
