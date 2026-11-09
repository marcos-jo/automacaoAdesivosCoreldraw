// Opt-in Windows/NSIS lab. Neutral synthetic uninstallers: no registry, real
// product GUID, shortcuts, profile, printer or installation is accessed.
import assert from 'node:assert/strict';
import {execFile,spawn} from 'node:child_process';
import {promisify} from 'node:util';
import {mkdtemp,mkdir,writeFile,readFile,readdir,realpath} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import path from 'node:path';
import {createHash} from 'node:crypto';
const run=promisify(execFile),lab=await mkdtemp(path.join(tmpdir(),'4ps-nsis-upgrade-'));
assert.equal(process.platform,'win32');
const metadata=JSON.parse(await readFile('node_modules/app-builder-lib/package.json','utf8'));assert.equal(metadata.version,'26.18.0');
const template=(await readFile('node_modules/app-builder-lib/templates/nsis/uninstaller.nsh','utf8')).replaceAll('\r\n','\n');
assert.equal(createHash('sha256').update(template).digest('hex'),'9ee2dac4593478083e8aa6f8487287ce9401006ccd50ecc538871d133ea4a42c','Revalidate the removal component before updating electron-builder');
const functions=template.slice(template.indexOf('Function un.atomicRMDir'),template.indexOf('!ifndef UNINSTALL_SECTION_NAME'));
const prefix=String.fromCharCode(92,92,63,92),compiler=path.join(process.env.LOCALAPPDATA,'electron-builder/Cache/nsis-3.0.4.1/nsis-3.0.4.1-1mx3n/makensis.exe');
const quote=s=>"'"+s.replaceAll("'","''")+"'",ps=async code=>(await run('powershell.exe',['-NoProfile','-NonInteractive','-Command',code],{windowsHide:true})).stdout.trim();
const names=['00-before.txt','01-locked.txt','4PrintShops Hold Test.exe','resources/app.asar','resources/acentos São José.txt','resources/libreoffice/help/media/icon-themes/chart2/res/valueaxisdirectpoints_52x60.svg'];
const cases=[{name:'short',length:70},{name:'legacy-long',length:172},{name:'alias-long',length:172,alias:true},{name:'extended-long',length:172,extended:true,defer:true},{name:'locked-extended',length:172,extended:true,defer:true,locked:true},{name:'alias-reselected-long',length:172,alias:true,reselect:true}];
const results=[];async function files(root,relative='',out=[]){for(const e of await readdir(path.join(root,relative),{withFileTypes:true})){const n=relative?relative+'/'+e.name:e.name;if(e.isDirectory())await files(root,n,out);else out.push(n);}return out;}
for(let index=0;index<cases.length;index++){
 const item=cases[index],padding=item.length-lab.length-1;assert.ok(padding>0);const root=path.join(lab,String.fromCharCode(97+index).repeat(padding));assert.equal(root.length,item.length);assert.ok(root.startsWith(lab+path.sep));
 for(const n of names){await mkdir(path.dirname(path.join(root,n)),{recursive:true});await writeFile(path.join(root,n),'SYNTHETIC '+n,{flag:'wx'});}
 const association=path.join(lab,item.name+'-association.txt');await writeFile(association,'REGISTERED'); // surrogate metadata, never Windows registry.
 let argument=root;if(item.alias){argument=await ps(`(New-Object -ComObject Scripting.FileSystemObject).GetFolder(${quote(root)}).ShortPath`);assert.ok(argument.length<root.length);assert.equal((await realpath(argument)).toLowerCase(),(await realpath(root)).toLowerCase());}
 let codeFunctions=functions;if(item.extended)codeFunctions=functions.replaceAll('"$INSTDIR','"'+prefix+'$INSTDIR').replaceAll('"$PLUGINSDIR','"'+prefix+'$PLUGINSDIR');
 const creator=path.join(lab,item.name+'-creator.exe'),uninstaller=path.join(lab,item.name+'-neutral-uninstaller.exe'),detail=path.join(lab,item.name+'-detail.txt'),status=path.join(lab,item.name+'-exit.txt'),quarantine=path.join(lab,item.name+'-quarantine');
 const script=`Unicode true
Name "4PS neutral migration laboratory"
OutFile "${creator}"
RequestExecutionLevel user
SilentInstall silent
SilentUnInstall silent
!include "LogicLib.nsh"
!include "FileFunc.nsh"
!define UNINSTALL_FILENAME "Unused-synthetic.exe"
Section
WriteUninstaller "${uninstaller}"
ExecWait '"${uninstaller}" /S /KEEP_APP_DATA /currentuser --updated _?=${argument}' $0
FileOpen $R8 "${status}" w
FileWrite $R8 "$0"
FileClose $R8
SectionEnd
${codeFunctions}
Function un.onUninstSuccess
${item.defer?'Delete "'+association+'"':''}
FunctionEnd
Section "Uninstall"
InitPluginsDir
FileOpen $R8 "${detail}" w
FileWrite $R8 "root=$INSTDIR$\\r$\\n"
\${GetParameters} $R9
FileWrite $R8 "args=$R9$\\r$\\n"
${item.reselect?'StrCpy $INSTDIR "'+root+'"\nFileWrite $R8 "selectedAgain=$INSTDIR$\\r$\\n"':''}
${item.defer?'':'Delete "'+association+'"'}
CreateDirectory "$PLUGINSDIR\\old-install"
Push ""
Call un.atomicRMDir
Pop $R0
FileWrite $R8 "atomic=$R0$\\r$\\n"
StrCmp $R0 "0" success
StrCpy $7 $R0
StrCpy $8 "$PLUGINSDIR\\native-move-check.tmp"
System::Call 'kernel32::MoveFileW(w r7, w r8) i.r9 ?e'
Pop $6
FileWrite $R8 "normalMove=$9 error=$6$\\r$\\n"
StrCpy $5 $7 4
StrCmp $5 "${prefix}" sourceExtended
StrCpy $5 "${prefix}$7"
Goto sourceReady
sourceExtended:
StrCpy $5 $7
sourceReady:
StrCpy $4 "${prefix}$8"
StrCmp $9 "0" tryExtended restoreNative
tryExtended:
System::Call 'kernel32::MoveFileW(w r5, w r4) i.r9 ?e'
Pop $6
FileWrite $R8 "extendedMove=$9 error=$6$\\r$\\n"
StrCmp $9 "0" restoreTree restoreNative
restoreNative:
System::Call 'kernel32::MoveFileW(w r4, w r5) i.r9 ?e'
Pop $6
FileWrite $R8 "restoreNative=$9$\\r$\\n"
restoreTree:
Push ""
Call un.restoreFiles
Pop $R0
FileClose $R8
Abort "Synthetic removal failed; fixtures restored"
success:
FileClose $R8
nsExec::ExecToStack '"$SYSDIR\\robocopy.exe" "$PLUGINSDIR\\old-install" "${quarantine}" /E /COPY:DAT /R:0 /W:0 /NFL /NDL /NJH /NJS /NP'
Pop $9
Pop $8
IntCmp $9 8 copyFailed done copyFailed
copyFailed:
Abort "Synthetic quarantine copy failed"
done:
SectionEnd
`;
 const nsi=path.join(lab,item.name+'.nsi');await writeFile(nsi,script);await run(compiler,['/V2',nsi],{windowsHide:true});
 let locker;
 try{
  if(item.locked){const locked=path.join(root,'01-locked.txt');locker=spawn('powershell.exe',['-NoProfile','-NonInteractive','-Command',`$f=[IO.File]::Open(${quote(locked)},[IO.FileMode]::Open,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None);[Console]::WriteLine('READY');[Console]::ReadLine()|Out-Null;$f.Dispose()`],{windowsHide:true,stdio:['pipe','pipe','pipe']});await new Promise((resolve,reject)=>{const timer=setTimeout(()=>reject(Error('Lock not ready')),15000);locker.stdout.on('data',d=>{if(String(d).includes('READY')){clearTimeout(timer);resolve();}});locker.once('error',reject);});}
  await run(creator,[],{windowsHide:true,timeout:60000});
 }finally{if(locker){const done=new Promise(resolve=>locker.once('exit',resolve));locker.stdin.end('release\n');await done;}}
 const exit=Number(await readFile(status,'utf8')),log=await readFile(detail,'utf8');assert.match(log,/\/KEEP_APP_DATA/);assert.match(log,/--updated/);assert.match(log,/\/currentuser/);const seen=log.split(/\r?\n/)[0].slice(5);assert.equal((await realpath(seen)).toLowerCase(),(await realpath(root)).toLowerCase());
 const failed=item.name==='legacy-long'||item.locked||item.reselect;assert.equal(exit,failed?2:0);const remaining=await files(root);const preserved=await readFile(association,'utf8').then(()=>true,e=>{if(e.code==='ENOENT')return false;throw e;});assert.equal(preserved,!!(failed&&item.defer));
 if(item.name==='legacy-long'){assert.match(log,/normalMove=0 error=3/);assert.match(log,/extendedMove=1 /);}if(item.locked){assert.match(log,/normalMove=0 error=32/);assert.match(log,/extendedMove=0 error=32/);}
 if(failed){assert.deepEqual(remaining.sort(),[...names].sort());for(const n of names)assert.equal(await readFile(path.join(root,n),'utf8'),'SYNTHETIC '+n);}
 else{assert.equal(remaining.length,0);assert.deepEqual((await files(quarantine)).sort(),[...names].sort());for(const n of names)assert.equal(await readFile(path.join(quarantine,n),'utf8'),'SYNTHETIC '+n);}
 const result={...item,exit,rootLength:root.length,argumentLength:argument.length,remaining:remaining.length,metadataPreserved:preserved,contentsVerified:true,log};results.push(result);console.log(JSON.stringify(result));
}
await writeFile(path.join(lab,'results.json'),JSON.stringify({passed:true,lab,results,noProductionInstallerExecuted:true,noRegistryWritten:true},null,2));console.log('PASSOU NSIS isolated upgrade; evidence: '+lab);
