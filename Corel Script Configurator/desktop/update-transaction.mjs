// Installer-only transaction. The old .34 uninstaller is never executed.
// Node/libuv supplies long-path-aware Windows file operations; no force delete.
import {readFile as readEmbeddedFile} from 'node:fs/promises';
import {createRequire} from 'node:module';
import {createHash,randomUUID} from 'node:crypto';
import {spawn} from 'node:child_process';
import path from 'node:path';
import {registryRequest,registryScope} from './update-registry.mjs';

// Electron's patched fs treats app.asar as a virtual directory. Installation
// must inventory/copy its bytes, while embedded metadata still uses patched fs.
const require=createRequire(import.meta.url);
const physicalFS=process.versions.electron?require('original-fs'):require('node:fs');
const {open,readFile,writeFile,rename,copyFile,mkdir,readdir,lstat,stat,unlink,rmdir,mkdtemp,realpath}=physicalFS.promises;
const {createReadStream,mkdtempSync}=physicalFS;

export const ownershipName='4ps-installed-files.ini';
const text=buffer=>buffer[0]===255&&buffer[1]===254?buffer.toString('utf16le'):buffer.toString('utf8');
const same=(a,b)=>path.resolve(a).toLowerCase()===path.resolve(b).toLowerCase();
const sha=async file=>{const h=createHash('sha256');for await(const b of createReadStream(file))h.update(b);return h.digest('hex');};
const exists=async file=>{try{await lstat(file);return true;}catch(e){if(e.code==='ENOENT')return false;throw e;}};
function inside(root,relative){if(typeof relative!=='string'||!relative||relative.includes('\\')||relative.includes(':')||relative.split('/').some(p=>!p||p==='.'||p==='..'))throw Error('Caminho de arquivo inválido na transação.');const out=path.resolve(root,relative);if(!out.toLowerCase().startsWith(path.resolve(root).toLowerCase()+path.sep))throw Error('Arquivo fora do aplicativo.');return out;}
async function files(root,relative='',result=[]){if(!await exists(root))return result;const info=await lstat(path.join(root,relative));if(info.isSymbolicLink()||!info.isDirectory())throw Error('Pasta vinculada ou inválida na distribuição.');for(const e of await readdir(path.join(root,relative),{withFileTypes:true})){const n=relative?relative+'/'+e.name:e.name;if(e.isSymbolicLink())throw Error('Link não permitido na distribuição: '+n);if(e.isDirectory())await files(root,n,result);else if(e.isFile())result.push(n);else throw Error('Tipo de arquivo não permitido: '+n);}return result;}
async function inventory(root){const result=[];for(const name of await files(root))result.push({name,sha256:await sha(inside(root,name)),size:(await stat(inside(root,name))).size});return result;}
async function directories(root,relative='',out=[]){if(!await exists(root))return out;out.push(relative);for(const e of await readdir(path.join(root,relative),{withFileTypes:true}))if(e.isDirectory()&&!e.isSymbolicLink())await directories(root,relative?relative+'/'+e.name:e.name,out);return out;}
async function save(file,value){const temp=file+'.tmp';await writeFile(temp,JSON.stringify(value,null,2),'utf8');await rename(temp,file);}
async function move(source,dest){await mkdir(path.dirname(dest),{recursive:true});await rename(source,dest);}
async function removeEmpty(root){if(!await exists(root))return;for(const e of await readdir(root,{withFileTypes:true}))if(e.isDirectory()&&!e.isSymbolicLink())await removeEmpty(path.join(root,e.name));try{await rmdir(root);}catch(e){if(!['ENOTEMPTY','EEXIST','EBUSY'].includes(e.code))throw e;}}
export async function readAsarJSON(file,relative){
    const fd=await open(file,'r');try{const start=Buffer.alloc(16);await fd.read(start,0,16,0);const headerSize=start.readUInt32LE(4),jsonSize=start.readUInt32LE(12);if(start.readUInt32LE(0)!==4||headerSize>16*1024*1024||jsonSize>headerSize-8)throw Error('Cabeçalho ASAR inválido.');const data=Buffer.alloc(jsonSize);await fd.read(data,0,jsonSize,16);let node=JSON.parse(data.toString('utf8'));for(const part of relative.split('/'))node=node.files?.[part];if(!node||node.unpacked||node.link||!Number.isSafeInteger(node.size)||node.size>2*1024*1024||!/^\d+$/.test(node.offset))throw Error('Metadados ASAR indisponíveis.');const offset=8+headerSize+Number(node.offset);if(offset+node.size>(await fd.stat()).size)throw Error('Metadados ASAR truncados.');const body=Buffer.alloc(node.size);await fd.read(body,0,body.length,offset);return JSON.parse(body.toString('utf8'));}finally{await fd.close();}
}
function validateIdentity(id){registryScope(id);if(!id.product||id.exe!==id.product+'.exe'||/[\\/:]/.test(id.product)||!/^1\.0\.0-experimental\.\d+$/.test(id.version)||id.cacheDir&&(!/^[a-z0-9._-]+$/i.test(id.cacheDir)||id.cacheDir==='.'||id.cacheDir==='..'))throw Error('Identidade do instalador inválida.');}
export function parseIni(text){const values={};for(const line of text.replace(/^\uFEFF/,'').split(/\r?\n/)){const i=line.indexOf('=');if(i>0)values[line.slice(0,i).trim()]=line.slice(i+1).trim();}return values;}
function treeValue(snapshot,root,name){const tree=snapshot.trees.find(t=>t.path===root);return tree?.values?.find(v=>v.name===name)?.data;}
function ownership(entries,id){const dirs=new Set();for(const e of entries){if(e.name.length>900)throw Error('Caminho excede o limite do inventário de desinstalação.');let p=e.name.split('/');p.pop();while(p.length){dirs.add(p.join('/'));p.pop();}}const list=[...dirs].sort((a,b)=>b.length-a.length);return '[identity]\nappId='+id.appId+'\nguid='+id.guid+'\n[files]\ncount='+entries.length+'\n'+entries.map((e,i)=>i+'='+e.name).join('\n')+'\n[directories]\ndirCount='+list.length+'\n'+list.map((d,i)=>'d'+i+'='+d).join('\n')+'\n';}
async function oldOwned(target,id,newNames){
    const owned=new Set([...newNames,ownershipName,'Uninstall '+id.product+'.exe','uninstallerIcon.ico']);
    const ini=path.join(target,ownershipName);
    if(await exists(ini)){const old=parseIni(text(await readFile(ini)));if(old.appId!==id.appId||old.guid!==id.guid||!Number.isInteger(Number(old.count))||Number(old.count)<1||Number(old.count)>100000)throw Error('Inventário anterior inválido.');for(let i=0;i<Number(old.count);i++){inside(target,old[i]);owned.add(old[i]);}}
    // Only public resources listed by the trusted legacy Office manifest.
    const manifest=path.join(target,'resources/libreoffice/runtime-manifest.json');
    if(await exists(manifest)){
        const digest=await sha(manifest);
        if(['01f0286083b702da3c085ec7219c22b98db2d86aa8ae3f6c3633f6707537247e','ae42c73b898efeb80813b024634d91cfd0e64503702b3c67e1588ea4b31d5781'].includes(digest)){
            const parsed=JSON.parse(await readFile(manifest,'utf8'));for(const f of parsed.files){const n='resources/libreoffice/'+f.path;inside(target,n);owned.add(n);}owned.add('resources/libreoffice/runtime-manifest.json');
        } // Unknown legacy runtime is preserved as user/unmanaged data.
    }
    return owned;
}
async function recordFile(file,backup){const present=await exists(file);if(!present)return {file,exists:false};if(backup){await copyFile(file,backup);return {file,exists:true,backup,sha256:await sha(backup)};}if((await stat(file)).size>1024*1024)throw Error('Atalho inesperadamente grande.');return {file,exists:true,bytes:(await readFile(file)).toString('base64')};}
async function restoreFile(entry){if(entry.exists){await mkdir(path.dirname(entry.file),{recursive:true});if(entry.backup){if(await sha(entry.backup)!==entry.sha256)throw Error('Cache anterior alterado.');await copyFile(entry.backup,entry.file);}else await writeFile(entry.file,Buffer.from(entry.bytes,'base64'));}else if(await exists(entry.file))await unlink(entry.file);}
export class UpdateTransaction {
    constructor({identity,stage,target,configFile,registry=registryRequest,verify,installer,cache,desktopLink,startLink}){
        validateIdentity(identity);this.id=identity;this.stage=path.resolve(stage);this.target=path.resolve(target);this.configFile=configFile;this.registry=registry;this.verify=verify;this.installer=installer;this.cache=cache;this.links=[desktopLink,startLink].filter(Boolean);
        if(!path.isAbsolute(target)||path.basename(this.target).toLowerCase()!==identity.product.toLowerCase()||this.target.length<12||same(this.target,this.stage)||this.stage.toLowerCase().startsWith(this.target.toLowerCase()+path.sep))throw Error('Diretório de atualização inválido.');
        this.pointer=path.join(path.dirname(configFile),'fourps-journal.json');
    }
    async busy(){const result=await this.registry(this.id,{op:'busy',pid:process.pid,target:this.target,processName:path.basename(this.id.exe,'.exe')});if(result.processes.length)throw Error('Feche normalmente o aplicativo e seus auxiliares antes de atualizar.');}
    externalPaths(environment){const parent=path.dirname(this.target),label=this.id.shortcut||this.id.product;if(/[\\/:]/.test(label))throw Error('Nome de atalho inválido.');return this.id.lab?{links:[path.join(parent,'lab-desktop',label+'.lnk'),path.join(parent,'lab-programs',label+'.lnk')],cache:path.join(parent,'lab-cache','installer.exe')}:{links:[path.join(environment.desktop,label+'.lnk'),path.join(environment.programs,label+'.lnk')],cache:path.join(environment.local,this.id.cacheDir,'installer.exe')};}
    validateExternal(environment){const expected=this.externalPaths(environment);if(this.cache&&!same(this.cache,expected.cache)||this.links.length!==expected.links.length||this.links.some((v,i)=>!same(v,expected.links[i])))throw Error('Caminho externo do instalador fora do escopo.');}
    async load(){const pointer=JSON.parse(await readFile(this.pointer,'utf8'));const prefix=path.join(path.dirname(this.target),'.4ps-upgrade-');if(typeof pointer.journal!=='string'||!pointer.journal.startsWith(prefix)||(await stat(pointer.journal)).size>32*1024*1024)throw Error('Diário de recuperação inválido.');const j=JSON.parse(await readFile(pointer.journal,'utf8'));if(j.schema!==1||j.identity.appId!==this.id.appId||j.identity.guid!==this.id.guid||!same(j.target,this.target)||!same(j.stage,this.stage)||!same(j.folder,path.dirname(pointer.journal)))throw Error('Diário de outra instalação.');const live=await this.registry(this.id,{op:'snapshot'});this.validateExternal(live.environment);if(j.links.some((v,i)=>!same(v.file,this.links[i]))||j.cache&&!same(j.cache.file,this.cache)||j.cache?.backup&&!same(j.cache.backup,path.join(j.folder,'cache-before')))throw Error('Backup externo fora do escopo.');this.journal=pointer.journal;this.j=j;return j;}
    async persist(){await save(this.journal,this.j);}
    async prepare(){
        if(await exists(this.pointer))throw Error('Esta transação já foi preparada; não repetir automaticamente.');
        const parent=path.dirname(this.target);
        if(await exists(parent))for(const entry of await readdir(parent,{withFileTypes:true})){
            if(!entry.isDirectory()||!entry.name.startsWith('.4ps-upgrade-'))continue;
            const file=path.join(parent,entry.name,'journal.json');if(!await exists(file))continue;
            if((await stat(file)).size>32*1024*1024)throw Error('Diário anterior excessivo; confira a recuperação antes de instalar.');
            const prior=JSON.parse(await readFile(file,'utf8'));
            if(prior.target&&same(prior.target,this.target)&&!['committed','rolled-back'].includes(prior.phase))throw Error('Existe atualização pendente de recuperação: '+file);
        }
        await this.busy();await this.verify(this.stage); // All integrity checks before touching the old application.
        const stagePackage=await readAsarJSON(path.join(this.stage,'resources/app.asar'),'package.json');if(stagePackage.version!==this.id.version)throw Error('Versão do pacote não corresponde ao instalador.');
        const snapshot=await this.registry(this.id,{op:'snapshot'}),scope=registryScope(this.id),previous=treeValue(snapshot,scope.uninstall,'DisplayVersion');this.validateExternal(snapshot.environment);
        if(previous&&!/^1\.0\.0-experimental\.(34|35|36)$/.test(previous))throw Error('Versão anterior ainda não suportada por esta migração.');
        const previousFiles=await inventory(this.target);
        if(previous){const recorded=treeValue(snapshot,scope.uninstall,'UninstallString')||'',match=recorded.match(/^"([^"]+)"\s+\/currentuser/);if(!match||!same(path.dirname(match[1]),this.target))throw Error('Registro anterior aponta para outra instalação.');const pkg=await readAsarJSON(path.join(this.target,'resources/app.asar'),'package.json');if(pkg.version!==previous)throw Error('Versão anterior divergente do registro.');}
        else if(previousFiles.length)throw Error('Pasta existente sem registro próprio: instalação recusada.');
        const next=await inventory(this.stage),owned=await oldOwned(this.target,this.id,next.map(f=>f.name));const managed=previousFiles.filter(f=>owned.has(f.name)),unmanaged=previousFiles.filter(f=>!owned.has(f.name));
        await mkdir(path.dirname(this.target),{recursive:true});const folder=await mkdtemp(path.join(path.dirname(this.target),'.4ps-upgrade-'));this.journal=path.join(folder,'journal.json');
        this.j={schema:1,identity:this.id,target:this.target,stage:this.stage,folder,phase:'prepared',snapshot,previous,managed,unmanaged,next,links:await Promise.all(this.links.map(file=>recordFile(file))),cache:this.cache?await recordFile(this.cache,path.join(folder,'cache-before')):null,oldExists:await exists(this.target),directories:await directories(this.target),started:new Date().toISOString()};
        await this.persist();await save(this.pointer,{journal:this.journal});return {prepared:true,previous,managed:managed.length,preserved:unmanaged.length,journal:this.journal};
    }
    async install(){
        const j=await this.load();if(j.phase!=='prepared')throw Error('Operação já iniciada; não repetir automaticamente.');
        await this.busy();
        // Recheck every old managed file before moving even the first one.
        for(const f of j.managed)if(await sha(inside(this.target,f.name))!==f.sha256)throw Error('Arquivo anterior mudou após a conferência: '+f.name);
        try{
            j.phase='moving';await this.persist();
            for(const f of j.managed){const src=inside(this.target,f.name);if(await sha(src)!==f.sha256)throw Error('Arquivo anterior mudou após a conferência: '+f.name);await move(src,inside(path.join(j.folder,'old'),f.name));}
            j.phase='copying';await this.persist();
            for(const f of j.next){const dest=inside(this.target,f.name);await mkdir(path.dirname(dest),{recursive:true});await copyFile(inside(this.stage,f.name),dest);if(await sha(dest)!==f.sha256)throw Error('Falha de integridade na cópia: '+f.name);}
            j.phase='verifying';await this.persist();await this.verify(this.target);
            const own=[...j.next,{name:ownershipName}];await writeFile(path.join(this.target,ownershipName),Buffer.from('\uFEFF'+ownership(own,this.id),'utf16le'));
            j.phase='installed';await this.persist();return {installed:true,journal:this.journal};
        }catch(error){await this.rollback();throw error;}
    }
    async commit(){
        const j=await this.load();if(j.phase!=='installed')throw Error('Instalação incompleta; commit recusado.');
        try{
            const now=await this.registry(this.id,{op:'snapshot'}),scope=registryScope(this.id);if(treeValue(now,scope.uninstall,'DisplayVersion')!==this.id.version)throw Error('Registro da versão nova não foi concluído.');
            const command=now.trees.find(t=>t.path===scope.classes+'\\'+this.id.appId+'.PDF')?.children?.find(t=>t.path.endsWith('\\shell'))?.children?.find(t=>t.path.endsWith('\\open'))?.children?.find(t=>t.path.endsWith('\\command'))?.values?.find(v=>v.name==='')?.data;
            if(!command||!command.includes(path.join(this.target,this.id.exe)))throw Error('Associação própria não foi registrada corretamente.');
            for(const f of j.unmanaged)if(await sha(inside(this.target,f.name))!==f.sha256)throw Error('Arquivo de usuário foi alterado durante a atualização.');
            for(const f of j.next)if(await sha(inside(this.target,f.name))!==f.sha256)throw Error('Arquivo novo foi alterado antes da confirmação.');
            if(this.installer&&j.cache){await mkdir(path.dirname(j.cache.file),{recursive:true});await copyFile(this.installer,j.cache.file);}
            j.phase='committed';await this.persist();return {committed:true,journal:this.journal,preserved:j.unmanaged.length};
        }catch(error){await this.rollback();throw error;}
    }
    async rollback(){
        let j;try{j=await this.load();}catch(e){if(e.code==='ENOENT')return {nothingToRollback:true};throw e;}
        if(j.phase==='rolled-back')return {rolledBack:true,journal:this.journal};if(j.phase==='committed')throw Error('Atualização já confirmada: rollback automático recusado.');if(j.phase==='recovery-required')throw Error('Recuperação manual necessária. Preserve '+j.folder+'; não repetir automaticamente.');
        if(j.phase==='prepared'){j.phase='rolled-back';await this.persist();return {rolledBack:true,noFilesChanged:true,journal:this.journal};}
        const phase=j.phase;j.phase='rolling-back';await this.persist();
        try{
            if(['copying','verifying','installed','rolling-back'].includes(phase)){
                for(const f of [...j.next,{name:ownershipName}]){const dest=inside(this.target,f.name);if(await exists(dest))await unlink(dest);}
                const before=new Set([...j.managed,...j.unmanaged].map(f=>f.name));for(const n of await files(this.target)){if(!before.has(n))await move(inside(this.target,n),inside(path.join(j.folder,'failed-output'),n));}
            }
            for(const f of j.managed){const backup=inside(path.join(j.folder,'old'),f.name);if(await exists(backup)){if(await sha(backup)!==f.sha256)throw Error('Backup alterado; recuperação interrompida.');await move(backup,inside(this.target,f.name));}}
            await this.registry(this.id,{op:'restore',snapshot:j.snapshot});for(const link of j.links)await restoreFile(link);if(j.cache)await restoreFile(j.cache);
            for(const f of [...j.managed,...j.unmanaged])if(await sha(inside(this.target,f.name))!==f.sha256)throw Error('Recuperação de arquivo não confirmada.');
            const oldDirs=new Set(j.directories);for(const n of (await directories(this.target)).reverse())if(!oldDirs.has(n)){try{await rmdir(n?inside(this.target,n):this.target);}catch(e){if(!['ENOTEMPTY','EEXIST','EBUSY'].includes(e.code))throw e;}}j.phase='rolled-back';await this.persist();return {rolledBack:true,journal:this.journal};
        }catch(error){j.phase='recovery-required';j.recoveryError=error.message;await this.persist();throw Error('Recuperação incompleta. Preserve '+j.folder+': '+error.message);}
    }
}

export async function runUpgradeCommand({app,resourcesPath=process.resourcesPath,action,configFile,log=console.log,error=console.error}){
    try{
        if(!app.isPackaged||!['prepare','install','commit','rollback'].includes(action))throw Error('Comando interno de instalação inválido.');
        const temp=mkdtempSync(path.join(app.getPath('temp'),'4ps-installer-'));app.setPath('userData',temp);
        const stage=path.resolve(resourcesPath,'..'),cfg=path.resolve(configFile);if(!same(path.dirname(cfg),path.dirname(stage))||path.basename(cfg)!=='fourps-upgrade.ini')throw Error('Configuração fora do instalador.');
        const identity=JSON.parse(await readEmbeddedFile(new URL('../build/distribution/identity.json',import.meta.url),'utf8'));
        const values=parseIni(text(await readFile(cfg)));if(values.AppId!==identity.appId||values.Guid!==identity.guid||!same(values.Stage,stage))throw Error('Identidade/configuração divergentes.');
        const verify=root=>new Promise((resolve,reject)=>{const child=spawn(path.join(root,identity.exe),['--verify-distribution'],{windowsHide:true,stdio:['ignore','pipe','pipe']});let output='';child.stdout.on('data',d=>output+=d);child.stderr.on('data',d=>output+=d);const timer=setTimeout(()=>{child.kill();reject(Error('Tempo limite na verificação do pacote.'));},240000);child.once('error',e=>{clearTimeout(timer);reject(e);});child.once('close',code=>{clearTimeout(timer);code===0?resolve():reject(Error('Verificação do pacote recusada: '+output.slice(-1000)));});});
        const transaction=new UpdateTransaction({identity,stage,target:values.Target,configFile:cfg,verify,installer:values.Installer,cache:values.Cache,desktopLink:values.DesktopLink,startLink:values.StartLink});
        log(JSON.stringify(await transaction[action]()));app.exit(0);
    }catch(e){error(e.message);app.exit(2);}
}
