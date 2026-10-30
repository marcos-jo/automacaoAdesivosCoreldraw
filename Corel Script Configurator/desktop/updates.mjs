import {EventEmitter} from 'node:events';
import {readFile} from 'node:fs/promises';
import path from 'node:path';

const notes=value=>(Array.isArray(value)?value.map(n=>n.note||'').join('\n'):String(value||'')).replace(/<[^>]*>/g,'').slice(0,4000);
export class UpdateController extends EventEmitter {
    constructor(updater,{canInstall,confirmInstall}) {
        super();this.updater=updater;this.canInstall=canInstall;this.confirmInstall=confirmInstall;this.state={phase:'idle'};
        updater.autoDownload=false;updater.autoInstallOnAppQuit=false;
        updater.on('checking-for-update',()=>this.set({phase:'checking'}));
        updater.on('update-not-available',()=>this.set({phase:'current'}));
        updater.on('update-available',info=>this.set({phase:'available',version:info.version,notes:notes(info.releaseNotes)}));
        updater.on('download-progress',progress=>this.set({...this.state,phase:'downloading',percent:Math.max(0,Math.min(100,Number(progress.percent)||0))}));
        updater.on('update-downloaded',info=>this.set({phase:'ready',version:info.version,notes:notes(info.releaseNotes)}));
        updater.on('error',error=>this.set({phase:'error',message:error.message}));
    }
    set(state){this.state=state;this.emit('state',state);}
    async check(){if(['checking','downloading','installing'].includes(this.state.phase))return this.state;try{await this.updater.checkForUpdates();}catch(error){this.set({phase:'error',message:error.message});}return this.state;}
    async download(){if(this.state.phase!=='available')throw new Error('Nenhuma atualização selecionada.');this.set({...this.state,phase:'downloading',percent:0});try{await this.updater.downloadUpdate();}catch(error){this.set({phase:'error',message:error.message});}return this.state;}
    async install(){
        if(this.state.phase!=='ready')throw new Error('Atualização ainda não está pronta.');
        if(!this.canInstall())throw new Error('Exporte ou descarte o documento e conclua a operação antes de instalar.');
        if(!await this.confirmInstall())return false;
        // A confirmação pode ser assíncrona. Nunca quitar/descartar trabalho aqui.
        if(!this.canInstall())throw new Error('O documento mudou durante a confirmação. Instalação cancelada.');
        this.set({...this.state,phase:'installing'});this.updater.quitAndInstall(false,true);return true;
    }
    later(){return this.state;} // Nenhuma instalação agendada ao sair.
}

export async function configuredUpdater(app, options) {
    let settings, metadata;
    try {
        settings=JSON.parse(await readFile(path.join(app.getPath('userData'),'updates.json'),'utf8'));
        if(!settings.enabled || !app.isPackaged)return null;
        const url=new URL(settings.url);
        if(url.protocol!=='https:' || url.username || url.password || url.search || url.hash)throw new Error('Provedor precisa usar HTTPS sem credenciais na URL.');
        // Identidade deve vir do pacote assinado, não de dados controlados pelo renderer.
        const yaml=await import('js-yaml');
        metadata=yaml.load(await readFile(path.join(process.resourcesPath,'app-update.yml'),'utf8'));
        const publishers=Array.isArray(metadata?.publisherName)?metadata.publisherName:[metadata?.publisherName];
        if(!publishers.length || publishers.some(p=>typeof p!=='string'||!p.trim()))return null;
        const module=await import('electron-updater'),{NsisUpdater}=module.default||module;
        const updater=new NsisUpdater({provider:'generic',url:url.href});
        // Não desabilitar a verificação Authenticode do electron-updater.
        updater.verifyUpdateCodeSignature=true;
        return new UpdateController(updater,options);
    }catch(error){if(error.code==='ENOENT')return null;throw error;}
}
