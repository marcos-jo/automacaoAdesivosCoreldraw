import test from'node:test';import assert from'node:assert/strict';import{mkdtemp,mkdir,writeFile,readFile,readdir,copyFile,symlink}from'node:fs/promises';import path from'node:path';import{tmpdir}from'node:os';import{randomUUID}from'node:crypto';import{createRequire}from'node:module';
import{UpdateTransaction,readAsarJSON,ownershipName}from'../desktop/update-transaction.mjs';import{registryScope,registryRequest}from'../desktop/update-registry.mjs';
import {spawn} from 'node:child_process';
import {unlink} from 'node:fs/promises';
import {hostname} from 'node:os';import {pathToFileURL} from 'node:url';import {createHash} from 'node:crypto';
import {residueName,moveResidue,recognizedResidue} from '../desktop/update-quarantine.mjs';
const require=createRequire(import.meta.url),asar=require('@electron/asar');
async function fixture({clean=false,fail=false,previousVersion="1.0.0-experimental.34",version="1.0.0-experimental.36"}={}){
 const lab=await mkdtemp(path.join(tmpdir(),'4ps-tx-test-')),id={appId:'br.com.4printshops.upgradelab.'+randomUUID(),guid:randomUUID(),product:'4PS Lab',exe:'4PS Lab.exe',version,lab:true},target=path.join(lab,'x'.repeat(172-lab.length-id.product.length-2),id.product),stage=path.join(lab,'stage'),cfg=path.join(lab,'fourps-upgrade.ini');assert.equal(target.length,172);await mkdir(stage);
 const temp=path.join(lab,'asar-input');await mkdir(temp);await writeFile(path.join(temp,'package.json'),JSON.stringify({name:'4ps-test',version:id.version}));await mkdir(path.join(stage,'resources'));await asar.createPackage(temp,path.join(stage,'resources/app.asar'));await writeFile(path.join(stage,id.exe),'NEW');
 const scope=registryScope(id),tree=(p,values=[])=>({path:p,exists:true,values,children:[]}),value=(name,data)=>({name,data,kind:'String'});const initial={trees:scope.allowedTrees.map(p=>tree(p)),values:scope.allowedValues.map(v=>({...v,exists:false})),environment:{},userChoice:{untouched:true}};
 if(!clean){await mkdir(path.join(target,'resources'),{recursive:true});await writeFile(path.join(temp,'package.json'),JSON.stringify({version:previousVersion}));await asar.createPackage(temp,path.join(target,'resources/app.asar'));await writeFile(path.join(target,id.exe),'OLD');const long='resources/libreoffice/help/media/icon-themes/chart2/res/valueaxisdirectpoints_52x60.svg';await mkdir(path.dirname(path.join(target,long)),{recursive:true});await writeFile(path.join(target,long),'OLD LONG');await writeFile(path.join(target,'cliente.pdf'),'SYNTHETIC USER FILE');await mkdir(path.join(target,'empty-user-directory'));await writeFile(path.join(target,ownershipName),'[identity]\nappId='+id.appId+'\nguid='+id.guid+'\n[files]\ncount=4\n0='+id.exe+'\n1=resources/app.asar\n2='+long+'\n3='+ownershipName+'\n');initial.trees.find(t=>t.path===scope.uninstall).values=[value('DisplayVersion',previousVersion),value('UninstallString','"'+path.join(target,'Uninstall '+id.product+'.exe')+'" /currentuser')];}
 else initial.trees.forEach(t=>t.exists=false);
 let state=structuredClone(initial);const registry=async(_id,r)=>{if(r.op==='snapshot')return structuredClone(state);if(r.op==='busy')return {processes:[]};if(r.op==='restore'){state=structuredClone(r.snapshot);return {restored:true};}throw Error('Unknown mock');};
 const links=[path.join(path.dirname(target),'lab-desktop',id.product+'.lnk'),path.join(path.dirname(target),'lab-programs',id.product+'.lnk')];for(const p of links){await mkdir(path.dirname(p),{recursive:true});await writeFile(p,'OLD LINK');}
 const tx=new UpdateTransaction({identity:id,stage,target,configFile:cfg,registry,verify:async root=>{assert.equal(await readFile(path.join(root,id.exe),'utf8'),'NEW');if(fail&&root===target)throw Error('Synthetic verification failure');},desktopLink:links[0],startLink:links[1]});
 const register=()=>{state.trees.find(t=>t.path===scope.uninstall).exists=true;state.trees.find(t=>t.path===scope.uninstall).values=[value('DisplayVersion',id.version)];const p=scope.classes+'\\'+id.appId+'.PDF';state.trees.find(t=>t.path===p).children=[{path:p+'\\shell',children:[{path:p+'\\shell\\open',children:[{path:p+'\\shell\\open\\command',values:[value('','"'+path.join(target,id.exe)+'" --open-pdf "%1"')]}]}]}];};
 return{lab,id,target,stage,tx,links,initial,register,state:()=>structuredClone(state)};
}
test('ASAR package reader is bounded and reads package version',async()=>{const f=await fixture();assert.equal((await readAsarJSON(path.join(f.stage,'resources/app.asar'),'package.json')).version,f.id.version);await assert.rejects(readAsarJSON(path.join(f.stage,'resources/app.asar'),'missing.json'));});
test('long-path migration preserves user files, empty folders and old metadata until commit',async()=>{const f=await fixture();await f.tx.prepare();await f.tx.install();assert.deepEqual(f.state(),f.initial);assert.equal(await readFile(path.join(f.target,'cliente.pdf'),'utf8'),'SYNTHETIC USER FILE');assert.ok((await readdir(f.target)).includes('empty-user-directory'));f.register();await f.tx.commit();assert.equal(await readFile(path.join(f.target,f.id.exe),'utf8'),'NEW');await assert.rejects(f.tx.rollback(),/confirmada/);});
test('failed new package integrity restores old long files and metadata',async()=>{const f=await fixture();await f.tx.prepare();await writeFile(path.join(f.stage,f.id.exe),'CORRUPTED AFTER PREPARE');await assert.rejects(f.tx.install(),/Falha de integridade na cópia/);assert.equal(await readFile(path.join(f.target,f.id.exe),'utf8'),'OLD');assert.equal((await readAsarJSON(path.join(f.target,'resources/app.asar'),'package.json')).version,'1.0.0-experimental.34');assert.equal(await readFile(path.join(f.target,'resources/libreoffice/help/media/icon-themes/chart2/res/valueaxisdirectpoints_52x60.svg'),'utf8'),'OLD LONG');assert.deepEqual(f.state(),f.initial);assert.equal(await readFile(path.join(f.target,'cliente.pdf'),'utf8'),'SYNTHETIC USER FILE');assert.equal(f.tx.j.phase,'rolled-back');assert.match(f.tx.j.failure.message,/Falha de integridade na cópia/);});
test('real copy error midway restores every previous file and metadata',async()=>{const f=await fixture();await f.tx.prepare();await unlink(path.join(f.stage,f.id.exe));await assert.rejects(f.tx.install(),/ENOENT/);assert.equal(await readFile(path.join(f.target,f.id.exe),'utf8'),'OLD');assert.equal((await readAsarJSON(path.join(f.target,'resources/app.asar'),'package.json')).version,'1.0.0-experimental.34');assert.deepEqual(f.state(),f.initial);assert.equal(f.tx.j.phase,'rolled-back');});
test('component startup refusal after copying restores the old installation and records its cause',async()=>{const f=await fixture();f.tx.verify=async root=>{if(root===f.target)throw Error('Private component startup refused');};await f.tx.prepare();await assert.rejects(f.tx.install(),/startup refused/);assert.equal(await readFile(path.join(f.target,f.id.exe),'utf8'),'OLD');assert.equal((await readAsarJSON(path.join(f.target,'resources/app.asar'),'package.json')).version,'1.0.0-experimental.34');assert.deepEqual(f.state(),f.initial);assert.equal(f.tx.j.phase,'rolled-back');const journal=JSON.parse(await readFile(f.tx.journal));assert.equal(journal.failure.phase,'verifying');assert.equal(journal.failure.message,'Private component startup refused');});
test('failure after metadata/shortcut changes restores previous state',async()=>{const f=await fixture();await f.tx.prepare();await f.tx.install();f.register();await writeFile(f.links[0],'NEW LINK');await f.tx.rollback();assert.deepEqual(f.state(),f.initial);assert.equal(await readFile(f.links[0],'utf8'),'OLD LINK');assert.equal(await readFile(path.join(f.target,f.id.exe),'utf8'),'OLD');});
test('rollback restores the previous installer cache bytes',async()=>{const f=await fixture(),cache=path.join(path.dirname(f.target),'lab-cache/installer.exe');await mkdir(path.dirname(cache));await writeFile(cache,'OLD VERIFIED CACHE');f.tx.cache=cache;await f.tx.prepare();await f.tx.install();f.register();await writeFile(cache,'PARTIAL NEW CACHE');await f.tx.rollback();assert.equal(await readFile(cache,'utf8'),'OLD VERIFIED CACHE');assert.deepEqual(f.state(),f.initial);});
test('clean installation uses same commit protocol and no second record',async()=>{const f=await fixture({clean:true});await f.tx.prepare();await f.tx.install();f.register();await f.tx.commit();assert.equal(f.state().trees.filter(t=>t.path.endsWith('\\Uninstall')).length,1);assert.equal(await readFile(path.join(f.target,f.id.exe),'utf8'),'NEW');});
test('pending sibling journal blocks another installer without touching files',async()=>{const f=await fixture();const recovery=path.join(path.dirname(f.target),'.4ps-upgrade-pending');await mkdir(recovery);await writeFile(path.join(recovery,'journal.json'),JSON.stringify({target:f.target,phase:'copying'}));await assert.rejects(f.tx.prepare(),/pendente de recuperação/);assert.equal(await readFile(path.join(f.target,f.id.exe),'utf8'),'OLD');});
test('invalid ownership cannot move files outside the application',async()=>{const f=await fixture();await writeFile(path.join(f.target,ownershipName),'appId='+f.id.appId+'\nguid='+f.id.guid+'\ncount=1\n0=../outside.txt');await assert.rejects(f.tx.prepare(),/Caminho/);assert.equal(await readFile(path.join(f.target,f.id.exe),'utf8'),'OLD');});
test('invalid staged ASAR metadata refuses before changing the previous version',async()=>{const f=await fixture();await writeFile(path.join(f.stage,'resources/app.asar'),'INVALID STAGED MANIFEST');await assert.rejects(f.tx.prepare(),/Cabeçalho ASAR inválido/);assert.equal(await readFile(path.join(f.target,f.id.exe),'utf8'),'OLD');assert.equal((await readAsarJSON(path.join(f.target,'resources/app.asar'),'package.json')).version,'1.0.0-experimental.34');assert.deepEqual(f.state(),f.initial);assert.equal(f.tx.j,undefined);});
test('links to an external directory refuse the transaction',async()=>{const f=await fixture(),outside=path.join(f.lab,'outside');await mkdir(outside);await writeFile(path.join(outside,'preserved.txt'),'EXTERNAL');await symlink(outside,path.join(f.target,'external-link'),'junction');await assert.rejects(f.tx.prepare(),/Link não permitido/);assert.equal(await readFile(path.join(outside,'preserved.txt'),'utf8'),'EXTERNAL');});
test('concurrent old-file change after prepare never starts moving files',async()=>{const f=await fixture();await f.tx.prepare();await writeFile(path.join(f.target,f.id.exe),'CHANGED');await assert.rejects(f.tx.install(),/mudou após/);await f.tx.rollback();assert.equal(await readFile(path.join(f.target,f.id.exe),'utf8'),'CHANGED');assert.deepEqual(f.state(),f.initial);});
test('real Windows file lock blocks migration and restores any earlier moves',{skip:process.platform!=='win32'},async()=>{
 const f=await fixture();await f.tx.prepare();const locked=path.join(f.target,'resources/app.asar').replaceAll("'","''");
 const holder=spawn('powershell.exe',['-NoProfile','-NonInteractive','-Command',`$f=[IO.File]::Open('${locked}',[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::Read);[Console]::WriteLine('READY');[Console]::Out.Flush();[Console]::ReadLine()|Out-Null;$f.Dispose()`],{windowsHide:true,stdio:['pipe','pipe','pipe']});
 try{await new Promise((resolve,reject)=>{const timer=setTimeout(()=>reject(Error('Lock helper readiness timeout')),20000);holder.stdout.once('data',d=>{clearTimeout(timer);String(d).includes('READY')?resolve():reject(Error(String(d)));});holder.once('error',e=>{clearTimeout(timer);reject(e);});});await assert.rejects(f.tx.install(),/EPERM|EBUSY|EACCES/);assert.equal(await readFile(path.join(f.target,f.id.exe),'utf8'),'OLD');assert.deepEqual(f.state(),f.initial);assert.equal(await readFile(path.join(f.target,'cliente.pdf'),'utf8'),'SYNTHETIC USER FILE');}
 finally{const stopped=new Promise(r=>holder.once('close',r));holder.stdin.end('\n');await stopped;}
});

for(const previousVersion of ['1.0.0-experimental.37','1.0.0-experimental.38','1.0.0-experimental.39'])test(`upgrade/reinstall ${previousVersion} to .39 preserves user files`,async()=>{
 const f=await fixture({previousVersion,version:'1.0.0-experimental.39'});await f.tx.prepare();await f.tx.install();f.register();await f.tx.commit();assert.equal(await readFile(path.join(f.target,'cliente.pdf'),'utf8'),'SYNTHETIC USER FILE');assert.equal((await readAsarJSON(path.join(f.target,'resources/app.asar'),'package.json')).version,f.id.version);
});
for(const previousVersion of ['1.0.0-experimental.40','1.0.0-experimental.33','1.0.0','other'])test(`unsupported/downgrade ${previousVersion} refuses without edits`,async()=>{
 const f=await fixture({previousVersion,version:'1.0.0-experimental.39'});await assert.rejects(f.tx.prepare(),/não pode ser atualizada/);assert.equal(await readFile(path.join(f.target,f.id.exe),'utf8'),'OLD');assert.deepEqual(f.state(),f.initial);assert.equal(f.tx.j,undefined);
});
test('clean reinstall after Windows uninstall preserves residual log and empty directories',async()=>{
 const f=await fixture({clean:true,version:'1.0.0-experimental.39'});await mkdir(path.join(f.target,'lo/empty'),{recursive:true});await writeFile(path.join(f.target,'debug.log'),'RESIDUAL LOG');await writeFile(path.join(f.target,'cliente.pdf'),'USER PDF');await f.tx.prepare();assert.equal(f.tx.j.unmanaged.length,2);await f.tx.install();f.register();await f.tx.commit();assert.equal(await readFile(path.join(f.target,'debug.log'),'utf8'),'RESIDUAL LOG');assert.equal(await readFile(path.join(f.target,'cliente.pdf'),'utf8'),'USER PDF');assert.ok((await readdir(path.join(f.target,'lo'))).includes('empty'));
});
test('unregistered colliding executable refuses instead of overwriting',async()=>{
 const f=await fixture({clean:true});await mkdir(f.target,{recursive:true});await writeFile(path.join(f.target,f.id.exe),'UNRELATED FILE');await assert.rejects(f.tx.prepare(),/arquivo em conflito/);assert.equal(await readFile(path.join(f.target,f.id.exe),'utf8'),'UNRELATED FILE');assert.equal(f.tx.j,undefined);
});
test('unregistered byte-identical remnants can be replaced; rollback preserves user files',async()=>{
 const f=await fixture({clean:true,fail:true});await mkdir(f.target,{recursive:true});await copyFile(path.join(f.stage,f.id.exe),path.join(f.target,f.id.exe));await writeFile(path.join(f.target,'debug.log'),'KEEP');await f.tx.prepare();await assert.rejects(f.tx.install(),/Synthetic verification failure/);assert.equal(await readFile(path.join(f.target,f.id.exe),'utf8'),'NEW');assert.equal(await readFile(path.join(f.target,'debug.log'),'utf8'),'KEEP');assert.deepEqual(f.state(),f.initial);
});
test('registered package version mismatch remains blocked',async()=>{
 const f=await fixture({previousVersion:'1.0.0-experimental.38',version:'1.0.0-experimental.39'});await copyFile(path.join(f.stage,'resources/app.asar'),path.join(f.target,'resources/app.asar'));await assert.rejects(f.tx.prepare(),/divergente do registro/);assert.equal(await readFile(path.join(f.target,f.id.exe),'utf8'),'OLD');
});
for(const rel of [ownershipName,'resources'])test(`unregistered generated metadata/file-directory collision ${rel} refuses`,async()=>{
 const f=await fixture({clean:true});await mkdir(f.target,{recursive:true});await writeFile(path.join(f.target,rel),'USER FILE');await assert.rejects(f.tx.prepare(),/arquivo em conflito/);assert.equal(await readFile(path.join(f.target,rel),'utf8'),'USER FILE');assert.equal(f.tx.j,undefined);
});
test('unregistered directory where incoming file belongs refuses before copying',async()=>{
 const f=await fixture({clean:true});await mkdir(path.join(f.target,f.id.exe),{recursive:true});await writeFile(path.join(f.target,f.id.exe,'cliente.pdf'),'KEEP');await assert.rejects(f.tx.prepare(),/arquivo em conflito/);assert.equal(await readFile(path.join(f.target,f.id.exe,'cliente.pdf'),'utf8'),'KEEP');
});

async function residueFixture(options={}){
 const f=await fixture(options),file=path.join(f.target,residueName),profile=path.join(tmpdir(),'4printshops-word-'+randomUUID().replaceAll('-',''),'profile');
 await mkdir(path.dirname(file),{recursive:true});const bytes=Buffer.from(',test/operator,'+hostname()+',16.11.2026 10:50,'+pathToFileURL(profile).href+';');await writeFile(file,bytes);
 const legit='lo/program/opencl/cl-test.ods';await writeFile(path.join(f.target,legit),'LEGIT OLD');await mkdir(path.dirname(path.join(f.stage,legit)),{recursive:true});await writeFile(path.join(f.stage,legit),'LEGIT NEW');
 const ini=path.join(f.target,ownershipName);await writeFile(ini,(await readFile(ini,'utf8')).replace('count=4','count=5')+'4='+legit+'\n');
 const verify=f.tx.verify;f.tx.verify=async root=>{await verify(root);await assert.rejects(readFile(path.join(root,residueName)),{code:'ENOENT'});assert.equal(await readFile(path.join(root,legit),'utf8'),'LEGIT NEW');};
 return {...f,file,profile,bytes,legit};
}
test('quarentena: .39 para .40 preserva bytes/hash fora de lo e mantém recurso legítimo',async()=>{
 const f=await residueFixture({previousVersion:'1.0.0-experimental.39',version:'1.0.0-experimental.40'});await f.tx.prepare();assert.equal(f.tx.j.quarantine.length,1);assert.equal(f.tx.j.quarantine[0].sha256,createHash('sha256').update(f.bytes).digest('hex'));assert.equal(f.tx.j.quarantine[0].name,residueName);
 await f.tx.install();f.register();const result=await f.tx.commit();assert.ok(result.quarantine);assert.deepEqual(await readFile(path.join(result.quarantine,'libreoffice-opencl-lock')),f.bytes);await assert.rejects(readFile(f.file),{code:'ENOENT'});assert.equal(await readFile(path.join(f.target,f.legit),'utf8'),'LEGIT NEW');assert.equal(await readFile(path.join(f.target,'cliente.pdf'),'utf8'),'SYNTHETIC USER FILE');
});
test('quarentena: falha de verificação restaura caminho, conteúdo, recursos e registro',async()=>{
 const f=await residueFixture({fail:true});await f.tx.prepare();await assert.rejects(f.tx.install(),/Synthetic verification failure/);assert.deepEqual(await readFile(f.file),f.bytes);assert.equal(await readFile(path.join(f.target,f.legit),'utf8'),'LEGIT OLD');assert.deepEqual(f.state(),f.initial);assert.equal(f.tx.j.phase,'rolled-back');
});
test('quarentena: falha após mover antes de confirmar não perde o resíduo',async()=>{
 const f=await residueFixture();await f.tx.prepare();let once=true;f.tx.quarantineMove=async args=>{await moveResidue(args);if(!args.restore&&once){once=false;throw Error('Falha simulada depois do move');}};
 await assert.rejects(f.tx.install(),/Falha simulada/);assert.deepEqual(await readFile(f.file),f.bytes);assert.equal(f.tx.j.phase,'rolled-back');
});
test('quarentena: interrupção após move é recuperável em nova instância pelo diário',async()=>{
 const f=await residueFixture();await f.tx.prepare();f.tx.j.phase='moving';await f.tx.persist();await moveResidue({target:f.target,folder:f.tx.j.folder,entry:f.tx.j.quarantine[0]});
 const resumed=new UpdateTransaction({identity:f.id,stage:f.stage,target:f.target,configFile:f.tx.configFile,registry:f.tx.registry,verify:f.tx.verify,desktopLink:f.links[0],startLink:f.links[1]});await resumed.rollback();assert.deepEqual(await readFile(f.file),f.bytes);assert.equal(resumed.j.phase,'rolled-back');
});
test('quarentena: colisão não substitui arquivo; restaurações recusam ambiguidade',async()=>{
 const f=await residueFixture();await f.tx.prepare();const stored=path.join(f.tx.j.folder,'quarantine/libreoffice-opencl-lock');await mkdir(path.dirname(stored));await writeFile(stored,'KEEP');
 await assert.rejects(moveResidue({target:f.target,folder:f.tx.j.folder,entry:f.tx.j.quarantine[0]}),/Quarentena recusada/);assert.equal(await readFile(stored,'utf8'),'KEEP');assert.deepEqual(await readFile(f.file),f.bytes);
 await assert.rejects(f.tx.install(),/Recuperação incompleta/);assert.equal(f.tx.j.phase,'recovery-required');assert.equal(await readFile(stored,'utf8'),'KEEP');assert.deepEqual(await readFile(f.file),f.bytes);assert.equal(await readFile(path.join(f.target,f.id.exe),'utf8'),'OLD');
});
test('quarentena: arquivo desconhecido em lo é recusado sem mover dados',async()=>{
 const f=await residueFixture();const unknown=path.join(f.target,'lo/unknown.txt');await writeFile(unknown,'KEEP');await assert.rejects(f.tx.prepare(),/Arquivo não gerenciado/);assert.equal(await readFile(unknown,'utf8'),'KEEP');assert.deepEqual(await readFile(f.file),f.bytes);assert.equal(f.tx.j,undefined);
});
test('quarentena: mesmo nome com conteúdo desconhecido ou perfil ainda existente não é reconhecido',async()=>{
 const f=await residueFixture();await writeFile(f.file,'USER DATA');assert.equal(await recognizedResidue(f.target,{name:residueName,size:9}),false);await writeFile(f.file,f.bytes);await mkdir(f.profile,{recursive:true});assert.equal(await recognizedResidue(f.target,{name:residueName,size:f.bytes.length}),false);await assert.rejects(f.tx.prepare(),/Arquivo não gerenciado/);assert.deepEqual(await readFile(f.file),f.bytes);
});
test('quarentena: caminhos adulterados no diário são rejeitados antes de mover',async()=>{
 const f=await residueFixture();await f.tx.prepare();f.tx.j.quarantine[0].name='../outside';await f.tx.persist();await assert.rejects(f.tx.install(),/fora do escopo/);assert.deepEqual(await readFile(f.file),f.bytes);assert.equal(await readFile(path.join(f.target,f.id.exe),'utf8'),'OLD');
});
test('quarentena: colisão ao restaurar preserva ambos os conteúdos e exige revisão',async()=>{
 const f=await residueFixture();await f.tx.prepare();const args={target:f.target,folder:f.tx.j.folder,entry:f.tx.j.quarantine[0]};await moveResidue(args);await writeFile(f.file,'NOVO');await assert.rejects(moveResidue({...args,restore:true}),/Quarentena recusada/);assert.equal(await readFile(f.file,'utf8'),'NOVO');assert.deepEqual(await readFile(path.join(f.tx.j.folder,'quarantine/libreoffice-opencl-lock')),f.bytes);
});
test('quarentena: conteúdo modificado após prepare não é movido',async()=>{
 const f=await residueFixture();await f.tx.prepare();await writeFile(f.file,'CHANGED');await assert.rejects(moveResidue({target:f.target,folder:f.tx.j.folder,entry:f.tx.j.quarantine[0]}),/Quarentena recusada/);assert.equal(await readFile(f.file,'utf8'),'CHANGED');
});
test('quarentena: raiz vinculada é recusada sem acessar destino externo',async()=>{
 const f=await residueFixture();await f.tx.prepare();const outside=await mkdtemp(path.join(tmpdir(),'4ps-outside-')),link=path.join(f.lab,'linked');await symlink(outside,link,'junction');await assert.rejects(moveResidue({target:link,folder:f.tx.j.folder,entry:f.tx.j.quarantine[0]}),/Raiz/);assert.deepEqual(await readFile(f.file),f.bytes);assert.deepEqual(await readdir(outside),[]);
});
test('quarentena: processo ocupado bloqueia instalação antes de qualquer movimento',async()=>{
 const f=await residueFixture();await f.tx.prepare();const registry=f.tx.registry;f.tx.registry=async(id,r)=>r.op==='busy'?{processes:[{pid:123,name:'soffice.bin'}]}:registry(id,r);await assert.rejects(f.tx.install(),/Feche normalmente/);assert.deepEqual(await readFile(f.file),f.bytes);assert.equal(await readFile(path.join(f.target,f.id.exe),'utf8'),'OLD');
});

test('quarentena: processo Windows com nome soffice bloqueia helper e detecção de ocupados',{skip:process.platform!=='win32'},async()=>{
 const f=await residueFixture();await f.tx.prepare();const exe=path.join(f.lab,'soffice.exe');await copyFile(path.join(process.env.SystemRoot,'System32/cmd.exe'),exe);
 const holder=spawn(exe,['/d','/q','/c','echo READY&set /p fourpsLabHold='],{windowsHide:true,stdio:['pipe','pipe','pipe']});
 try{await new Promise((resolve,reject)=>{const timer=setTimeout(()=>reject(Error('Process timeout')),10000);holder.stdout.once('data',()=>{clearTimeout(timer);resolve();});holder.once('error',reject);});
  const busy=await registryRequest(f.id,{op:'busy',target:f.target});assert.ok(busy.processes.some(p=>p.pid===holder.pid));
  await assert.rejects(moveResidue({target:f.target,folder:f.tx.j.folder,entry:f.tx.j.quarantine[0]}),/Quarentena recusada/);assert.deepEqual(await readFile(f.file),f.bytes);
 }finally{const closed=new Promise(r=>holder.once('close',r));holder.stdin.end('\n');await closed;}
});
test('quarentena: handle Windows ativo impede movimentação',{skip:process.platform!=='win32'},async()=>{
 const f=await residueFixture();await f.tx.prepare();const locked=f.file.replaceAll("'","''"),holder=spawn('powershell.exe',['-NoProfile','-NonInteractive','-Command',`$f=[IO.File]::Open('${locked}',[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::Read);[Console]::WriteLine('READY');[Console]::Out.Flush();[Console]::ReadLine()|Out-Null;$f.Dispose()`],{windowsHide:true,stdio:['pipe','pipe','pipe']});
 try{await new Promise((resolve,reject)=>{const t=setTimeout(()=>reject(Error('Lock timeout')),20000);holder.stdout.once('data',()=>{clearTimeout(t);resolve();});holder.once('error',reject);});await assert.rejects(moveResidue({target:f.target,folder:f.tx.j.folder,entry:f.tx.j.quarantine[0]}),/Quarentena recusada/);assert.deepEqual(await readFile(f.file),f.bytes);}finally{const closed=new Promise(r=>holder.once('close',r));holder.stdin.end('\n');await closed;}
});
