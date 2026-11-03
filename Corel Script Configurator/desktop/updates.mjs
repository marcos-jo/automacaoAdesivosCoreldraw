import {EventEmitter} from 'node:events';
import {readFile} from 'node:fs/promises';
import path from 'node:path';
import {verifyInstallerSignature} from './update-signature.mjs';

const notes=value=>(Array.isArray(value)?value.map(n=>n.note||'').join('\n'):String(value||'')).replace(/<[^>]*>/g,'').slice(0,4000);
export class UpdateController extends EventEmitter {
    constructor(updater,{canInstall,confirmInstall}) {
        super();this.updater=updater;this.canInstall=canInstall;this.confirmInstall=confirmInstall;this.state={phase:'idle'};
        updater.autoDownload=false;updater.autoInstallOnAppQuit=false;
        updater.on('checking-for-update',()=>{if(this.state.phase!=='checking')this.set({phase:'checking'});});
        updater.on('update-not-available',()=>this.set({phase:'current'}));
        updater.on('update-available',info=>this.set({phase:'available',version:info.version,notes:notes(info.releaseNotes)}));
        updater.on('download-progress',progress=>this.set({...this.state,phase:'downloading',percent:Math.max(0,Math.min(100,Number(progress.percent)||0))}));
        updater.on('update-downloaded',info=>this.set({phase:'ready',version:info.version,notes:notes(info.releaseNotes)}));
        updater.on('error',error=>this.failure(error));
    }
    set(state){this.state=state;this.emit('state',state);}
    failure(error){this.set({phase:'error',message:error.message,code:error.code||''});}
    async check(){if(['checking','downloading','ready','installing'].includes(this.state.phase))return this.state;this.set({phase:'checking'});try{await this.updater.checkForUpdates();}catch(error){this.failure(error);}return this.state;}
    async download(){if(this.state.phase!=='available')throw new Error('Nenhuma atualização selecionada.');this.set({...this.state,phase:'downloading',percent:0});try{await this.updater.downloadUpdate();}catch(error){this.failure(error);}return this.state;}
    async install(){
        if(this.state.phase!=='ready')throw new Error('Atualização ainda não está pronta.');
        if(!await this.canInstall())throw new Error('Exporte ou feche conscientemente as abas com alterações e conclua a operação antes de instalar.');
        if(!await this.confirmInstall())return false;
        // A confirmação pode ser assíncrona. Nunca quitar/descartar trabalho aqui.
        if(!await this.canInstall())throw new Error('O documento mudou durante a confirmação. Instalação cancelada.');
        if(this.state.phase!=='ready')throw new Error('A atualização deixou de estar pronta. Tente novamente.');
        this.set({...this.state,phase:'installing'});
        try{this.updater.quitAndInstall(false,true);}catch(error){this.failure(error);throw error;}return true;
    }
    later(){return this.state;} // Nenhuma instalação agendada ao sair.
}

export function updateSource(metadata, settings, lab = false) {
    const publishers=Array.isArray(metadata?.publisherName)?metadata.publisherName:[metadata?.publisherName];
    if(!publishers.length || publishers.some(p=>typeof p!=='string'||!p.trim()))throw new Error('Atualizações exigem identidade de editor no pacote.');
    if(metadata.verifyUpdateCodeSignature===false)throw new Error('Verificação de assinatura deve permanecer ativa.');
    const source=lab ? metadata : settings;
    if(!source?.url || !lab && !source.enabled)return null;
    const url=new URL(source.url);
    const local=lab && url.protocol==='http:' && url.hostname==='127.0.0.1';
    if((url.protocol!=='https:'&&!local)||url.username||url.password||url.search||url.hash)throw new Error('Provedor precisa usar HTTPS sem credenciais; HTTP apenas no laboratório compilado para 127.0.0.1.');
    const channel=source.channel||metadata.channel||'latest';
    if(!/^[a-z][a-z0-9-]{0,39}$/.test(channel))throw new Error('Canal de atualização inválido.');
    return {provider:'generic',url:url.href,channel};
}

export async function configuredUpdater(app, options) {
    try {
        if(!app.isPackaged)return null;
        const pkg=JSON.parse(await readFile(new URL('../package.json',import.meta.url),'utf8'));
        const lab=pkg.desktopUpdateLab===true && pkg.name==='4printshops-update-lab';
        let settings;
        if(!lab){settings=JSON.parse(await readFile(path.join(app.getPath('userData'),'updates.json'),'utf8'));if(!settings.enabled)return null;}
        // Identidade deve vir do pacote assinado, não de dados controlados pelo renderer.
        const yaml=await import('js-yaml');
        const metadata=yaml.load(await readFile(path.join(process.resourcesPath,'app-update.yml'),'utf8'));
        const source=updateSource(metadata,settings,lab);
        if(!source)return null;
        const module=await import('electron-updater'),{NsisUpdater}=module.default||module;
        const updater=new NsisUpdater(source);
        // Na API 6.8.11 essa propriedade é uma FUNÇÃO, não um booleano.
        // Preservar o verificador original, acrescentando prova que falha fechada.
        const nativeVerify=updater.verifyUpdateCodeSignature;
        updater.verifyUpdateCodeSignature=async(names,file)=>(await verifyInstallerSignature(names,file))??nativeVerify(names,file);
        updater.channel=source.channel;updater.allowDowngrade=false;
        updater.disableWebInstaller=true;
        return new UpdateController(updater,options);
    }catch(error){if(error.code==='ENOENT')return null;throw error;}
}
