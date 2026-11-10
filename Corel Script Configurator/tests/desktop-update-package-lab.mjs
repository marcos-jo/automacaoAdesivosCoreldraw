// Opt-in real NSIS package test. Never execute with the Hold Test identity.
import assert from 'node:assert/strict';
import {mkdtemp,mkdir,writeFile,readFile,readdir,copyFile} from 'node:fs/promises';
import {createReadStream} from 'node:fs';import {createHash} from 'node:crypto';
import {tmpdir} from 'node:os';import path from 'node:path';import {spawn} from 'node:child_process';
import {createRequire} from 'node:module';import {registryRequest,registryScope} from '../desktop/update-registry.mjs';
const require=createRequire(import.meta.url),asar=require('@electron/asar');
const action=process.argv[2];assert.ok(['clean','migrate','startup'].includes(action),'Specify clean, migrate or startup explicitly');
const upgrading=action!=='clean';
const packageRoot=path.resolve('dist-desktop/upgrade-transaction-lab/win-unpacked');
const id=JSON.parse(asar.extractFile(path.join(packageRoot,'resources/app.asar'),path.join('build','distribution','identity.json')));
assert.equal(id.appId,'br.com.4printshops.upgradelab.transaction36');assert.equal(id.lab,true);assert.equal(id.product,'4PrintShops Upgrade Lab');
const installer=path.resolve('dist-desktop/upgrade-transaction-lab/4PrintShops-Upgrade-Lab-1.0.0-experimental.36-Setup.exe');
const lab=await mkdtemp(path.join(tmpdir(),'4ps-full-upgrade-')),scope=registryScope(id);
const rootLength=action==='startup'?196:172;
const root=path.join(lab,'x'.repeat(rootLength-lab.length-id.product.length-2),id.product);assert.equal(root.length,rootLength);
const hash=async p=>{const h=createHash('sha256');for await(const b of createReadStream(p))h.update(b);return h.digest('hex');};
async function inventory(folder,relative='',out=[]){for(const f of await readdir(path.join(folder,relative),{withFileTypes:true})){const n=relative?relative+'/'+f.name:f.name;if(f.isDirectory())await inventory(folder,n,out);else{assert.ok(f.isFile());out.push({name:n,sha256:await hash(path.join(folder,n))});}}return out;}
async function checkedRun(exe,args,{verbatim=false}={}){const began=Date.now(),child=spawn(exe,args,{windowsHide:true,windowsVerbatimArguments:verbatim,stdio:['ignore','pipe','pipe']});let out='';child.stdout.on('data',d=>out+=d);child.stderr.on('data',d=>out+=d);console.log('Started laboratory process',child.pid,path.basename(exe));const code=await new Promise((r,j)=>{child.once('error',j);child.once('close',r);});await writeFile(path.join(lab,'process-'+began+'.log'),out);console.log(JSON.stringify({exit:code,durationMs:Date.now()-began}));return{code,out};}
const before=await registryRequest(id,{op:'snapshot'});assert.ok(before.trees.every(t=>!t.exists),'Preserve an earlier lab and investigate before another run');
const normalBefore=await registryRequest({...id,lab:false},{op:'snapshot'});
let seededSnapshot,seededFiles,legacyStartup;
const option=name=>process.argv.find(v=>v.startsWith('--'+name+'='))?.slice(name.length+3);
async function startup(legacy=false){const result=await checkedRun(process.execPath,['tests/desktop-upgrade-startup.mjs',root,...(legacy?['--legacy-modules']:[])]);assert.equal(result.code,0,'Application must remain usable: '+result.out);const details=JSON.parse(result.out.split(/\r?\n/).find(line=>line.startsWith('{"passed":')));return details;}
await writeFile(path.join(lab,'before-registry.json'),JSON.stringify(before,null,2));
await mkdir(path.join(lab,'profile'));await writeFile(path.join(lab,'profile/preferences.json'),'SYNTHETIC PREFERENCES');
await writeFile('.desktop-test/upgrade36-full-current.json',JSON.stringify({lab,root,installer,action},null,2));
if(upgrading){
 const source=path.resolve(option('source')||'.desktop-test/graph34/setup-payload');
 const missingFile=option('missing');const missing=new Set(missingFile?JSON.parse(await readFile(missingFile)).expectedMissing:[]);
 assert.equal(JSON.parse(asar.extractFile(path.join(source,'resources/app.asar'),'package.json')).version,'1.0.0-experimental.34');
 const old=await inventory(source);let copied=0;for(const f of old){if(missing.has(f.name)||/^(?:4PrintShops Hold Test|Uninstall 4PrintShops Hold Test)\.exe$/.test(f.name))continue;const dest=path.join(root,f.name);await mkdir(path.dirname(dest),{recursive:true});await copyFile(path.join(source,f.name),dest);copied++;}
 // Inert markers: accidentally executing an old fixture cannot touch the real app.
 await writeFile(path.join(root,id.exe),'INERT OLD APPLICATION MARKER');
 await writeFile(path.join(root,'Uninstall '+id.product+'.exe'),'INERT OLD UNINSTALLER: MUST NEVER EXECUTE');
 await writeFile(path.join(root,'user-file.txt'),'SYNTHETIC USER DATA');await mkdir(path.join(root,'empty-user-folder'));
 const seed=structuredClone(before),v=(name,data)=>({name,data,kind:'String'});
 const install=seed.trees.find(t=>t.path===scope.install);Object.assign(install,{exists:true,values:[v('InstallLocation',root),v('KeepShortcuts','true'),v('ShortcutName',id.shortcut)],children:[]});
 const uninstall=seed.trees.find(t=>t.path===scope.uninstall);Object.assign(uninstall,{exists:true,values:[v('DisplayVersion','1.0.0-experimental.34'),v('UninstallString','"'+path.join(root,'Uninstall '+id.product+'.exe')+'" /currentuser')],children:[]});
 if(action==='migrate'){
  const tree=(p,values=[],children=[])=>({path:p,exists:true,values,children}),progid=id.appId+'.PDF',command='"'+path.join(root,id.exe)+'" --open-pdf "%1"';
  for(const [base,values] of [[scope.classes+'\\'+progid,[v('','OLD 34 PDF capability')]],[scope.classes+'\\Applications\\'+id.exe,[v('FriendlyAppName','OLD 34 Upgrade Lab')]]])Object.assign(seed.trees.find(t=>t.path===base),tree(base,values,[tree(base+'\\shell',[],[tree(base+'\\shell\\open',[],[tree(base+'\\shell\\open\\command',[v('',command)])])])]));
  Object.assign(seed.trees.find(t=>t.path===scope.capabilities),tree(scope.capabilities,[v('ApplicationName','OLD 34 Upgrade Lab')],[tree(scope.capabilities+'\\FileAssociations',[v('.pdf',progid)])]));
  for(const value of seed.values)Object.assign(value,{exists:true,data:value.name===id.appId?scope.capabilities:'',keyExists:true,parentExists:true});
  for(const location of ['lab-desktop','lab-programs']){const folder=path.join(path.dirname(root),location);await mkdir(folder,{recursive:true});await writeFile(path.join(folder,id.shortcut+'.lnk'),'INERT OLD 34 SHORTCUT '+location);}
  await mkdir(path.join(path.dirname(root),'lab-cache'),{recursive:true});await writeFile(path.join(path.dirname(root),'lab-cache/installer.exe'),'INERT OLD 34 INSTALLER CACHE');
 }
 await registryRequest(id,{op:'restore',snapshot:seed});await writeFile(path.join(lab,'seeded.json'),JSON.stringify({copied,missing:missing.size,source,root,registry:seed,legacyExecutablesInert:true,graphMissingInventoryAvailable:!!missingFile},null,2));
 legacyStartup=await startup(true);
  seededSnapshot=await registryRequest(id,{op:'snapshot'});seededFiles=await inventory(root);
  await writeFile(path.join(lab,'seeded-registry.json'),JSON.stringify(seededSnapshot,null,2));await writeFile(path.join(lab,'seeded-files.json'),JSON.stringify(seededFiles));
}
const receipt={lab,root,action,installer,installerSha256:await hash(installer),started:new Date().toISOString()};await writeFile(path.join(lab,'receipt.json'),JSON.stringify(receipt,null,2));
const plans=action==='startup'?[
 {name:'rollback-startup',file:installer,exit:2,detail:/conversor privado .*inicializou/}
]:upgrading?[
 {name:'rollback-copy',file:path.join(path.dirname(installer),'4PrintShops-Upgrade-Lab-FailCopy-1.0.0-experimental.36-Setup.exe'),exit:2,detail:/Falha simulada.*c.*pia/},
 {name:'rollback-registry',file:path.join(path.dirname(installer),'4PrintShops-Upgrade-Lab-FailRegistry-1.0.0-experimental.36-Setup.exe'),exit:2,detail:/Falha simulada.*registro/},
 {name:'rollback-locked',file:installer,exit:2,lock:true,detail:/EPERM|EACCES|EBUSY|operation not permitted|permission denied/i},
 {name:'migration',file:installer,exit:0}
]:[{name:'clean',file:installer,exit:0}];receipt.operations=[];
for(const plan of plans){
 const parent=path.dirname(root),prior=await readdir(parent).catch(e=>e.code==='ENOENT'?[]:Promise.reject(e));
 let holder,result;
 try{
  if(plan.lock){const file=path.join(root,'resources/app.asar').replaceAll("'","''");holder=spawn(path.join(process.env.SystemRoot,'System32/WindowsPowerShell/v1.0/powershell.exe'),['-NoProfile','-NonInteractive','-Command',`$f=[IO.File]::Open('${file}',[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::Read);[Console]::WriteLine('READY');[Console]::Out.Flush();[Console]::ReadLine()|Out-Null;$f.Dispose()`],{windowsHide:true,stdio:['pipe','pipe','pipe']});await new Promise((r,j)=>{holder.stdout.once('data',d=>String(d).includes('READY')?r():j(Error(String(d))));holder.once('error',j);});}
  result=await checkedRun(plan.file,['/S','/currentuser','/D='+root],{verbatim:true});
 }finally{if(holder){const stopped=new Promise(r=>holder.once('close',r));holder.stdin.end('\n');await stopped;}}
 receipt.operations.push({name:plan.name,exit:result.code,sha256:await hash(plan.file)});await writeFile(path.join(lab,'receipt.json'),JSON.stringify(receipt,null,2));assert.equal(result.code,plan.exit,'Unexpected result: stop and preserve lab');
 if(plan.exit!==0){
  const bytes=await readFile(path.join(parent,'.4ps-upgrade-last-error.ini'));let detail=bytes.toString(bytes[0]===255&&bytes[1]===254?'utf16le':'utf8');if(detail.includes('\uFFFD'))detail=new TextDecoder('windows-1252').decode(bytes);assert.match(detail,plan.detail,'A different failure must not pass the intended rollback case');receipt.operations.at(-1).failureDetail=detail;
  assert.deepEqual(await registryRequest(id,{op:'snapshot'}),seededSnapshot,'Actual rollback must restore the seeded registry');
  assert.deepEqual(await inventory(root),seededFiles,'Actual rollback must restore every old file');
  if(action==='migrate'){for(const location of ['lab-desktop','lab-programs'])assert.equal(await readFile(path.join(path.dirname(root),location,id.shortcut+'.lnk'),'utf8'),'INERT OLD 34 SHORTCUT '+location);assert.equal(await readFile(path.join(path.dirname(root),'lab-cache/installer.exe'),'utf8'),'INERT OLD 34 INSTALLER CACHE');}
  assert.ok((await readdir(root)).includes('empty-user-folder'));assert.equal(await readFile(path.join(lab,'profile/preferences.json'),'utf8'),'SYNTHETIC PREFERENCES');
  const created=(await readdir(parent,{withFileTypes:true})).filter(e=>e.isDirectory()&&e.name.startsWith('.4ps-upgrade-')&&!prior.includes(e.name));assert.equal(created.length,1,'Preflight failure is not a successful rollback test');const journal=JSON.parse(await readFile(path.join(parent,created[0].name,'journal.json')));assert.equal(journal.phase,'rolled-back');assert.equal(journal.previous,'1.0.0-experimental.34');receipt.operations.at(-1).rollbackVerified=true;await writeFile(path.join(lab,'receipt.json'),JSON.stringify(receipt,null,2));
  if(action==='startup'){assert.equal(journal.failure.phase,'verifying');assert.match(journal.failure.message,/conversor privado .*inicializou/);receipt.operations.at(-1).journalFailure=journal.failure;}
  const usable=await startup(true);assert.equal(usable.storageBefore['upgrade36.synthetic.preference'],'preserved');receipt.operations.at(-1).legacyModulesUsable=true;await writeFile(path.join(lab,'receipt.json'),JSON.stringify(receipt,null,2));
 }
}
if(action==='startup'){
 assert.deepEqual(await registryRequest({...id,lab:false},{op:'snapshot'}),normalBefore);
 Object.assign(receipt,{passed:true,startupFailureAfterCopy:true,oldFilesRestored:true,oldModulesUsable:true,legacyExecutablesInert:true,defaultPdfUnchanged:true,profilePreserved:true});
 await writeFile(path.join(lab,'receipt.json'),JSON.stringify(receipt,null,2));
 // Only release this synthetic registry fixture after the rollback and usable
 // old modules were proved. Keep every file, profile, journal and backup.
 await registryRequest(id,{op:'restore',snapshot:before});receipt.labRegistrationRestoredAfterProof=true;
 await writeFile(path.join(lab,'receipt.json'),JSON.stringify(receipt,null,2));console.log(JSON.stringify(receipt,null,2));process.exit(0);
}
const after=await registryRequest(id,{op:'snapshot'});assert.deepEqual(after.userChoice,before.userChoice);assert.deepEqual(after.pdfDefault,before.pdfDefault);assert.equal(after.trees.find(t=>t.path===scope.uninstall).values.find(v=>v.name==='DisplayVersion').data,id.version);
assert.deepEqual(await registryRequest({...id,lab:false},{op:'snapshot'}),normalBefore,'No duplicate/global registration is allowed');
function registryValue(snapshot,p,name=''){const visit=t=>t.path===p?t.values?.find(v=>v.name===name)?.data:t.children?.map(visit).find(v=>v!==undefined);return snapshot.trees.map(visit).find(v=>v!==undefined);}
const progid=id.appId+'.PDF',command='"'+path.join(root,id.exe)+'" --open-pdf "%1"';
assert.equal(registryValue(after,scope.classes+'\\'+progid+'\\shell\\open\\command'),command);
assert.equal(registryValue(after,scope.classes+'\\Applications\\'+id.exe+'\\shell\\open\\command'),command);
assert.equal(registryValue(after,scope.capabilities+'\\FileAssociations','.pdf'),progid);
assert.equal(after.values.find(v=>v.path===scope.registered&&v.name===id.appId)?.data,scope.capabilities);
assert.equal(after.values.find(v=>v.name===progid)?.exists,true);
for(const location of ['lab-desktop','lab-programs'])assert.ok((await readFile(path.join(path.dirname(root),location,id.shortcut+'.lnk'))).length>0);
assert.equal(await hash(path.join(path.dirname(root),'lab-cache/installer.exe')),await hash(installer),'Committed installer cache must match the current Setup');
const actualStartup=await startup();if(upgrading)assert.equal(actualStartup.storageBefore['upgrade36.synthetic.preference'],'preserved');
const expected=await inventory(packageRoot),installed=await inventory(root);for(const file of expected)assert.equal(await hash(path.join(root,file.name)),file.sha256,file.name);
assert.equal((await checkedRun(path.join(root,id.exe),['--verify-distribution'])).code,0);
const word=await checkedRun(process.execPath,['scripts/test-private-office.mjs',root]);assert.equal(word.code,0,'Private DOC/DOCX conversion must pass');
assert.equal(await readFile(path.join(lab,'profile/preferences.json'),'utf8'),'SYNTHETIC PREFERENCES');
if(action==='migrate'){assert.equal(await readFile(path.join(root,'user-file.txt'),'utf8'),'SYNTHETIC USER DATA');assert.ok((await readdir(root)).includes('empty-user-folder'));}
Object.assign(receipt,{expected:expected.length,installed:installed.length,allExpectedHashes:true,privateOfficeStartup:true,profilePreserved:true,defaultPdfUnchanged:true,userChoiceUnchanged:true,actualStartup,legacyStartup,legacyExecutablesInert:action==='migrate'});await writeFile(path.join(lab,'receipt.json'),JSON.stringify(receipt,null,2));await writeFile(path.join(lab,'installed-inventory.json'),JSON.stringify(installed));
if(action==='clean'){
 await writeFile(path.join(root,'user-file.txt'),'SYNTHETIC USER DATA');await mkdir(path.join(root,'empty-user-folder'));
 const runner=path.join(lab,'uninstaller-runner.exe');await copyFile(path.join(root,'Uninstall '+id.product+'.exe'),runner);
 // _? prevents NSIS's asynchronous self-copy launcher; wait for the real work.
 const un=await checkedRun(runner,['/S','/currentuser','_?='+root],{verbatim:true});assert.equal(un.code,0);
 assert.equal(await readFile(path.join(root,'user-file.txt'),'utf8'),'SYNTHETIC USER DATA');assert.ok((await readdir(root)).includes('empty-user-folder'));
 const removed=await registryRequest(id,{op:'snapshot'});assert.ok(removed.trees.every(t=>!t.exists));assert.ok(removed.values.every(v=>!v.exists));assert.deepEqual(removed.userChoice,before.userChoice);assert.deepEqual(removed.pdfDefault,before.pdfDefault);
 receipt.selectiveUninstall=true;
 // Keep the verified backup as evidence; do not release storage in this task.
 for(const e of await readdir(path.dirname(root),{withFileTypes:true}))if(e.isDirectory()&&e.name.endsWith('.tmp')){
  const backup=path.resolve(path.dirname(root),e.name);assert.ok(backup.startsWith(path.resolve(lab)+path.sep));
  const list=await inventory(backup);assert.equal(list.length,installed.length+1);for(const file of installed)assert.equal(await hash(path.join(backup,file.name)),file.sha256);
  await writeFile(path.join(lab,'uninstall-backup-proof.json'),JSON.stringify({backup,verified:list.length,sha256:await hash(path.join(backup,'ownership.ini'))},null,2));receipt.verifiedLabBackupPreserved=true;
 }
 await writeFile(path.join(lab,'receipt.json'),JSON.stringify(receipt,null,2));
}
console.log(JSON.stringify(receipt,null,2));
