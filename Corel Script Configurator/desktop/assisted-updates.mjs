import {EventEmitter} from 'node:events';
import {readFile,writeFile,rename,mkdir,unlink} from 'node:fs/promises';
import {randomUUID} from 'node:crypto';
import path from 'node:path';
import {authenticateEnvelope,validatePolicy,validateRelease,sourceURL} from './assisted-manifest.mjs';
import {getManifest,downloadInstaller,verifyCached} from './assisted-http.mjs';

export class AssistedUpdates extends EventEmitter{
    constructor({policy,source,currentVersion,folder,now=()=>Date.now(),networkOptions={}}){super();this.policy=validatePolicy(policy);this.source=sourceURL(source.url,{allowHttp:source.allowHttp===true});this.currentVersion=currentVersion;this.folder=folder;this.now=now;this.networkOptions=networkOptions;this.state={phase:'idle'};this.busy=false;}
    set(state){this.state=state;this.emit('state',state);}
    async loadFloor(){try{const floor=JSON.parse(await readFile(path.join(this.folder,'publication-state.json'),'utf8'));if(floor.product!==this.policy.product||floor.channel!==this.policy.channel||!Number.isSafeInteger(floor.sequence)||floor.sequence<1||!/^[a-f0-9]{64}$/.test(floor.digest))throw new Error('Estado local de publicação inválido.');return floor;}catch(error){if(error.code==='ENOENT')return {sequence:0};throw error;}}
    async persistFloor(release){await mkdir(this.folder,{recursive:true});const file=path.join(this.folder,'publication-state.json'),tmp=path.join(this.folder,randomUUID()+'.json.tmp');try{await writeFile(tmp,JSON.stringify({product:this.policy.product,channel:this.policy.channel,sequence:release.data.sequence,digest:release.digest}),{flag:'wx',mode:0o600});await rename(tmp,file);}finally{await unlink(tmp).catch(()=>{});}}
    validate(release,floor){return validateRelease(release,{currentVersion:this.currentVersion,floor,now:this.now()});}
    async check(){
        if(this.busy)return this.state;this.busy=true;this.set({phase:'checking'});
        try{const bytes=await getManifest(new URL('channels/'+this.policy.channel+'.json',this.source),this.networkOptions);const release=authenticateEnvelope(bytes,this.policy),floor=await this.loadFloor();const available=this.validate(release,floor);await this.persistFloor(release);this.release=release;this.file=null;
            this.set(available?{phase:'available',version:release.data.version,notes:release.data.notes}:{phase:'current'});
        }catch(error){this.release=null;this.file=null;this.set({phase:'error',message:error.message});}finally{this.busy=false;}
        return this.state;
    }
    async download(){
        if(this.busy||this.state.phase!=='available'||!this.release)throw new Error('Nenhuma publicação verificada disponível.');this.busy=true;
        try{this.validate(this.release,await this.loadFloor());const artifact=this.release.data.installer,url=new URL(artifact.path,this.source);if(url.origin!==this.source.origin||!url.pathname.startsWith(this.source.pathname))throw new Error('Caminho fora da origem autorizada.');
            this.set({phase:'downloading',version:this.release.data.version,percent:0});
            this.file=await downloadInstaller(url,artifact,path.join(this.folder,'installers'),percent=>this.set({phase:'downloading',version:this.release.data.version,percent}),this.networkOptions);
            this.set({phase:'ready',version:this.release.data.version,notes:this.release.data.notes});
        }catch(error){this.file=null;this.set({phase:'error',message:error.message});}finally{this.busy=false;}return this.state;
    }
    async preparedFile(){
        if(this.busy||this.state.phase!=='ready'||!this.release||!this.file)throw new Error('Instalador verificado ainda não está pronto.');
        try{
            this.validate(this.release,await this.loadFloor());
            if(!await verifyCached(this.file,this.release.data.installer))throw new Error('O instalador local foi modificado. Baixe novamente.');
            return this.file;
        }catch(error){this.file=null;this.set({phase:'error',message:error.message});throw error;}
    }
}

export async function configuredAssisted(app,pkg){
    if(!pkg.assistedUpdates)return null;
    const policy=validatePolicy(pkg.assistedUpdates);
    let source;
    if(pkg.name==='4printshops-assisted-lab'&&pkg.desktopAssistedLab===true)source={url:'http://127.0.0.1:18084/',allowHttp:true};
    else{try{source=JSON.parse(await readFile(path.join(app.getPath('userData'),'assisted-updates.json'),'utf8'));}catch(error){if(error.code==='ENOENT')return null;throw error;}if(!source.enabled)return null;}
    return new AssistedUpdates({policy,source,currentVersion:app.getVersion(),folder:path.join(app.getPath('userData'),'assisted-updates')});
}
