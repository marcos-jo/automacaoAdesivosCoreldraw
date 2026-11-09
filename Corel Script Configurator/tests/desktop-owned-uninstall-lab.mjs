// Opt-in neutral NSIS laboratory; no real product registry or profile access.
import assert from 'node:assert/strict';
import {execFile,spawn} from 'node:child_process';
import {promisify} from 'node:util';
import {mkdtemp,mkdir,writeFile,readFile,readdir} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import path from 'node:path';
import {randomUUID} from 'node:crypto';
assert.equal(process.platform,'win32');
const run=promisify(execFile),lab=await mkdtemp(path.join(tmpdir(),'4ps-owned-uninstall-'));
const compiler=path.join(process.env.LOCALAPPDATA,'electron-builder/Cache/nsis-3.0.4.1/nsis-3.0.4.1-1mx3n/makensis.exe');
const include=path.resolve('build/nsis/uninstall-owned.nsh'),results=[];
for(const item of [{name:'long',count:19000},{name:'locked',lock:true},{name:'invalid',invalid:true}]){
 const folder=path.join(lab,item.name),root=path.join(folder,'r'.repeat(172-folder.length-1));await mkdir(root,{recursive:true});assert.equal(root.length,172);
 const names=['00-first.txt','01-locked.txt','resources/libreoffice/help/media/icon-themes/chart2/res/valueaxisdirectpoints_52x60.svg'];
 for(const n of names){await mkdir(path.dirname(path.join(root,n)),{recursive:true});await writeFile(path.join(root,n),'SYNTHETIC '+n);}
 await writeFile(path.join(root,'user-file.txt'),'KEEP');await mkdir(path.join(root,'empty-user-folder'));
 const id='br.com.4printshops.upgradelab.'+randomUUID(),guid=randomUUID(),entries=[...names,...Array.from({length:item.count||0},(_,i)=>'missing-public-'+i+'.txt'),'4ps-installed-files.ini'];if(item.invalid)entries.push('../outside.txt');
 await writeFile(path.join(root,'4ps-installed-files.ini'),Buffer.from('\uFEFF[identity]\nappId='+id+'\nguid='+guid+'\n[files]\ncount='+entries.length+'\n'+entries.map((n,i)=>i+'='+n).join('\n')+'\n[directories]\ndirCount=0\n','utf16le'));
 const creator=path.join(folder,'creator.exe'),uninstaller=path.join(folder,'neutral-uninstaller.exe'),status=path.join(folder,'exit.txt'),marker=path.join(folder,'association.txt');await writeFile(marker,'REGISTERED');
 const source=`Unicode true
Name "Neutral owned removal lab"
OutFile "${creator}"
RequestExecutionLevel user
SilentInstall silent
SilentUnInstall silent
!include LogicLib.nsh
!include FileFunc.nsh
!define BUILD_UNINSTALLER
!define APP_ID "${id}"
!define APP_GUID "${guid}"
!include "${include}"
Section
WriteUninstaller "${uninstaller}"
ExecWait '"${uninstaller}" /S _?=${root}' $0
FileOpen $1 "${status}" w
FileWrite $1 "$0"
FileClose $1
SectionEnd
Function un.onUninstSuccess
Delete "${marker}"
FunctionEnd
Section "Uninstall"
!insertmacro customRemoveFiles
SectionEnd
`;
 const script=path.join(folder,'test.nsi');await writeFile(script,source);await run(compiler,['-V2',script],{windowsHide:true});
 let holder;
 try{
  if(item.lock){holder=spawn('powershell.exe',['-NoProfile','-NonInteractive','-Command',`$f=[IO.File]::Open('${path.join(root,names[1]).replaceAll("'","''")}',[IO.FileMode]::Open,[IO.FileAccess]::ReadWrite,[IO.FileShare]::Read);[Console]::WriteLine('READY');[Console]::Out.Flush();[Console]::ReadLine()|Out-Null;$f.Dispose()`],{windowsHide:true,stdio:['pipe','pipe','pipe']});await new Promise((resolve,reject)=>{holder.stdout.once('data',d=>String(d).includes('READY')?resolve():reject(Error(String(d))));holder.once('error',reject);});}
  const start=performance.now();await run(creator,[],{windowsHide:true,timeout:240000});const exit=Number(await readFile(status,'utf8')),elapsedMs=Math.round(performance.now()-start);
  assert.equal(exit,item.lock||item.invalid?2:0);assert.equal(await readFile(path.join(root,'user-file.txt'),'utf8'),'KEEP');assert.ok((await readdir(root)).includes('empty-user-folder'));
  if(item.lock||item.invalid){for(const n of names)assert.equal(await readFile(path.join(root,n),'utf8'),'SYNTHETIC '+n);assert.equal(await readFile(marker,'utf8'),'REGISTERED');}
  else{await assert.rejects(readFile(path.join(root,names[2])));await assert.rejects(readFile(marker));}
  results.push({case:item.name,exit,elapsedMs,rootLength:root.length,manifestEntries:entries.length});
 }finally{if(holder){holder.stdin.end('\n');await new Promise(r=>holder.once('close',r));}}
}
await writeFile(path.join(lab,'results.json'),JSON.stringify(results,null,2));console.log(JSON.stringify({lab,results},null,2));
