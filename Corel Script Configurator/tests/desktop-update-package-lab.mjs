// Opt-in real NSIS package test. Never execute with the Hold Test identity.
import assert from 'node:assert/strict';
import {mkdtemp,mkdir,writeFile,readFile,readdir,copyFile,stat,rm} from 'node:fs/promises';
import {createReadStream} from 'node:fs';import {createHash} from 'node:crypto';
import {tmpdir} from 'node:os';import path from 'node:path';import {spawn} from 'node:child_process';
import {createRequire} from 'node:module';import {registryRequest,registryScope} from '../desktop/update-registry.mjs';
const require=createRequire(import.meta.url),asar=require('@electron/asar');
const action=process.argv[2];assert.ok(['clean','migrate'].includes(action),'Specify clean or migrate explicitly');
const packageRoot=path.resolve('dist-desktop/upgrade-transaction-lab/win-unpacked');
const id=JSON.parse(asar.extractFile(path.join(packageRoot,'resources/app.asar'),path.join('build','distribution','identity.json')));
assert.equal(id.appId,'br.com.4printshops.upgradelab.transaction36');assert.equal(id.lab,true);assert.equal(id.product,'4PrintShops Upgrade Lab');
const installer=path.resolve('dist-desktop/upgrade-transaction-lab/4PrintShops-Upgrade-Lab-1.0.0-experimental.36-Setup.exe');
const lab=await mkdtemp(path.join(tmpdir(),'4ps-full-upgrade-')),scope=registryScope(id);
const root=path.join(lab,'x'.repeat(172-lab.length-id.product.length-2),id.product);assert.equal(root.length,172);
const hash=async p=>{const h=createHash('sha256');for await(const b of createReadStream(p))h.update(b);return h.digest('hex');};
async function inventory(folder,relative='',out=[]){for(const f of await readdir(path.join(folder,relative),{withFileTypes:true})){const n=relative?relative+'/'+f.name:f.name;if(f.isDirectory())await inventory(folder,n,out);else{assert.ok(f.isFile());out.push({name:n,sha256:await hash(path.join(folder,n))});}}return out;}
async function checkedRun(exe,args,{verbatim=false}={}){const began=Date.now(),child=spawn(exe,args,{windowsHide:true,windowsVerbatimArguments:verbatim,stdio:['ignore','pipe','pipe']});let out='';child.stdout.on('data',d=>out+=d);child.stderr.on('data',d=>out+=d);console.log('Started laboratory process',child.pid,path.basename(exe));const code=await new Promise((r,j)=>{child.once('error',j);child.once('close',r);});await writeFile(path.join(lab,'process-'+began+'.log'),out);console.log(JSON.stringify({exit:code,durationMs:Date.now()-began}));return{code,out};}
const before=await registryRequest(id,{op:'snapshot'});assert.ok(before.trees.every(t=>!t.exists),'Preserve an earlier lab and investigate before another run');
let seededSnapshot,seededFiles;
await writeFile(path.join(lab,'before-registry.json'),JSON.stringify(before,null,2));
await mkdir(path.join(lab,'profile'));await writeFile(path.join(lab,'profile/preferences.json'),'SYNTHETIC PREFERENCES');
await writeFile('.desktop-test/upgrade36-full-current.json',JSON.stringify({lab,root,installer,action},null,2));
if(action==='migrate'){
 const source=path.resolve('.desktop-test/graph34/setup-payload'),missing=new Set(JSON.parse(await readFile('.desktop-test/active-upgrade-integrity.json')).expectedMissing);
 assert.equal(JSON.parse(asar.extractFile(path.join(source,'resources/app.asar'),'package.json')).version,'1.0.0-experimental.34');
 const old=await inventory(source);let copied=0;for(const f of old){if(missing.has(f.name)||/^(?:4PrintShops Hold Test|Uninstall 4PrintShops Hold Test)\.exe$/.test(f.name))continue;const dest=path.join(root,f.name);await mkdir(path.dirname(dest),{recursive:true});await copyFile(path.join(source,f.name),dest);copied++;}
 // Inert markers: accidentally executing an old fixture cannot touch the real app.
 await writeFile(path.join(root,id.exe),'INERT OLD APPLICATION MARKER');
 await writeFile(path.join(root,'Uninstall '+id.product+'.exe'),'INERT OLD UNINSTALLER: MUST NEVER EXECUTE');
 await writeFile(path.join(root,'user-file.txt'),'SYNTHETIC USER DATA');await mkdir(path.join(root,'empty-user-folder'));
 const seed=structuredClone(before),v=(name,data)=>({name,data,kind:'String'});
 const install=seed.trees.find(t=>t.path===scope.install);Object.assign(install,{exists:true,values:[v('InstallLocation',root),v('KeepShortcuts','true'),v('ShortcutName',id.shortcut)],children:[]});
 const uninstall=seed.trees.find(t=>t.path===scope.uninstall);Object.assign(uninstall,{exists:true,values:[v('DisplayVersion','1.0.0-experimental.34'),v('UninstallString','"'+path.join(root,'Uninstall '+id.product+'.exe')+'" /currentuser')],children:[]});
 await registryRequest(id,{op:'restore',snapshot:seed});await writeFile(path.join(lab,'seeded.json'),JSON.stringify({copied,missing:missing.size,source,root,registry:seed},null,2));
 seededSnapshot=await registryRequest(id,{op:'snapshot'});seededFiles=await inventory(root);
}
const receipt={lab,root,action,installer,installerSha256:await hash(installer),started:new Date().toISOString()};await writeFile(path.join(lab,'receipt.json'),JSON.stringify(receipt,null,2));
const plans=action==='migrate'?[
 {name:'rollback-copy',file:path.join(path.dirname(installer),'4PrintShops-Upgrade-Lab-FailCopy-1.0.0-experimental.36-Setup.exe'),exit:2},
 {name:'rollback-registry',file:path.join(path.dirname(installer),'4PrintShops-Upgrade-Lab-FailRegistry-1.0.0-experimental.36-Setup.exe'),exit:2},
 {name:'migration',file:installer,exit:0}
]:[{name:'clean',file:installer,exit:0}];receipt.operations=[];
for(const plan of plans){
 const parent=path.dirname(root),prior=await readdir(parent).catch(e=>e.code==='ENOENT'?[]:Promise.reject(e));
 const result=await checkedRun(plan.file,['/S','/currentuser','/D='+root],{verbatim:true});receipt.operations.push({name:plan.name,exit:result.code,sha256:await hash(plan.file)});await writeFile(path.join(lab,'receipt.json'),JSON.stringify(receipt,null,2));assert.equal(result.code,plan.exit,'Unexpected result: stop and preserve lab');
 if(plan.exit!==0){
  assert.deepEqual(await registryRequest(id,{op:'snapshot'}),seededSnapshot,'Actual rollback must restore the seeded registry');
  assert.deepEqual(await inventory(root),seededFiles,'Actual rollback must restore every old file');
  assert.ok((await readdir(root)).includes('empty-user-folder'));assert.equal(await readFile(path.join(lab,'profile/preferences.json'),'utf8'),'SYNTHETIC PREFERENCES');
  const created=(await readdir(parent,{withFileTypes:true})).filter(e=>e.isDirectory()&&e.name.startsWith('.4ps-upgrade-')&&!prior.includes(e.name));assert.equal(created.length,1,'Preflight failure is not a successful rollback test');const journal=JSON.parse(await readFile(path.join(parent,created[0].name,'journal.json')));assert.equal(journal.phase,'rolled-back');assert.equal(journal.previous,'1.0.0-experimental.34');receipt.operations.at(-1).rollbackVerified=true;await writeFile(path.join(lab,'receipt.json'),JSON.stringify(receipt,null,2));
 }
}
const after=await registryRequest(id,{op:'snapshot'});assert.deepEqual(after.userChoice,before.userChoice);assert.equal(after.trees.find(t=>t.path===scope.uninstall).values.find(v=>v.name==='DisplayVersion').data,id.version);
const expected=await inventory(packageRoot),installed=await inventory(root);for(const file of expected)assert.equal(await hash(path.join(root,file.name)),file.sha256,file.name);
assert.equal((await checkedRun(path.join(root,id.exe),['--verify-distribution'])).code,0);
const word=await checkedRun(process.execPath,['scripts/test-private-office.mjs',root]);assert.equal(word.code,0,'Private DOC/DOCX conversion must pass');
assert.equal(await readFile(path.join(lab,'profile/preferences.json'),'utf8'),'SYNTHETIC PREFERENCES');
if(action==='migrate'){assert.equal(await readFile(path.join(root,'user-file.txt'),'utf8'),'SYNTHETIC USER DATA');assert.ok((await readdir(root)).includes('empty-user-folder'));}
Object.assign(receipt,{expected:expected.length,installed:installed.length,allExpectedHashes:true,privateOfficeStartup:true,profilePreserved:true,defaultPdfUnchanged:true});await writeFile(path.join(lab,'receipt.json'),JSON.stringify(receipt,null,2));await writeFile(path.join(lab,'installed-inventory.json'),JSON.stringify(installed));
if(action==='clean'){
 await writeFile(path.join(root,'user-file.txt'),'SYNTHETIC USER DATA');await mkdir(path.join(root,'empty-user-folder'));
 const runner=path.join(lab,'uninstaller-runner.exe');await copyFile(path.join(root,'Uninstall '+id.product+'.exe'),runner);
 // _? prevents NSIS's asynchronous self-copy launcher; wait for the real work.
 const un=await checkedRun(runner,['/S','/currentuser','_?='+root],{verbatim:true});assert.equal(un.code,0);
 assert.equal(await readFile(path.join(root,'user-file.txt'),'utf8'),'SYNTHETIC USER DATA');assert.ok((await readdir(root)).includes('empty-user-folder'));
 const removed=await registryRequest(id,{op:'snapshot'});assert.ok(removed.trees.every(t=>!t.exists));assert.ok(removed.values.every(v=>!v.exists));assert.deepEqual(removed.userChoice,before.userChoice);
 receipt.selectiveUninstall=true;
 // Discard only newly generated, verified package backups inside this lab.
 for(const e of await readdir(path.dirname(root),{withFileTypes:true}))if(e.isDirectory()&&e.name.endsWith('.tmp')){
  const backup=path.resolve(path.dirname(root),e.name);assert.ok(backup.startsWith(path.resolve(lab)+path.sep));
  const list=await inventory(backup);assert.equal(list.length,installed.length+1);for(const file of installed)assert.equal(await hash(path.join(backup,file.name)),file.sha256);
  await writeFile(path.join(lab,'uninstall-backup-proof.json'),JSON.stringify({backup,verified:list.length,sha256:await hash(path.join(backup,'ownership.ini'))},null,2));await rm(backup,{recursive:true});receipt.verifiedLabBackupReleased=true;
 }
 await writeFile(path.join(lab,'receipt.json'),JSON.stringify(receipt,null,2));
}
console.log(JSON.stringify(receipt,null,2));
