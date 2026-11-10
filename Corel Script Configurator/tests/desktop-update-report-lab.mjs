// Real GUI child + nsExec: a blank console must never hide errors or approve a
// missing/stale/mismatched result. No registry, profile, Office or printing.
import assert from 'node:assert/strict';
import {mkdtemp,mkdir,writeFile,readFile} from 'node:fs/promises';
import {tmpdir} from 'node:os';import path from 'node:path';
import {execFile} from 'node:child_process';import {promisify} from 'node:util';
const run=promisify(execFile),lab=await mkdtemp(path.join(tmpdir(),'4ps-report-lab-')),stage=path.join(lab,'stage');await mkdir(stage);await mkdir(path.join(lab,'captured'));
const macro=(await readFile('build/nsis/upgrade-transaction.nsh','utf8')).match(/!macro fourpsCallUpgrade ACTION[\s\S]*?!macroend/)[0];
const code=`using System;using System.IO;using System.Text;using System.Runtime.InteropServices;
class ReportProbe {
[DllImport("kernel32.dll",CharSet=CharSet.Unicode)]static extern int GetPrivateProfileString(string s,string k,string d,StringBuilder b,int n,string p);
static int Main(string[] args){string cfg=null;foreach(string a in args)if(a.StartsWith("--4ps-config="))cfg=a.Substring(13);string mode=File.ReadAllText(Path.Combine(AppDomain.CurrentDomain.BaseDirectory,"mode.txt"));if(mode=="missing")return 0;var b=new StringBuilder(32);GetPrivateProfileString("Update","Request","",b,32,cfg);int request=int.Parse(b.ToString());int exit=mode=="failure"?2:0;string action=mode=="wrong"?"commit":"prepare";if(mode=="stale")request--;string report="[Result]\\r\\nAction="+action+"\\r\\nExit="+exit+"\\r\\nMessage="+(exit==2?"Conversor privado não inicializou: sintético São José":"Resultado sintético confirmado")+"\\r\\n";File.WriteAllText(Path.Combine(Path.GetDirectoryName(cfg),"fourps-result-prepare-"+request+".ini"),report,Encoding.Unicode);return exit;}}
`;
await writeFile(path.join(lab,'probe.cs'),code);
const csc=path.join(process.env.SystemRoot,'Microsoft.NET/Framework64/v4.0.30319/csc.exe');
await run(csc,['/nologo','/target:winexe','/platform:x64','/out:'+path.join(stage,'probe.exe'),path.join(lab,'probe.cs')],{windowsHide:true,timeout:30000});
const status=path.join(lab,'status.ini');await writeFile(status,Buffer.from('\uFEFF','utf16le'));
const cases=['missing','failure','wrong','stale','success'];
let source=`Unicode true\nName "Isolated GUI result lab"\nOutFile "${lab}\\report.exe"\nRequestExecutionLevel user\nSilentInstall silent\n!include LogicLib.nsh\n!define APP_EXECUTABLE_FILENAME "probe.exe"\nVar fourpsStage\nVar fourpsConfig\nVar fourpsSequence\n${macro}\nSection\nInitPluginsDir\nStrCpy $fourpsStage "${stage}"\nStrCpy $fourpsConfig "$PLUGINSDIR\\fourps-upgrade.ini"\nFileOpen $0 $fourpsConfig w\nFileWriteByte $0 255\nFileWriteByte $0 254\nFileClose $0\n`;
for(const mode of cases)source+=`FileOpen $0 "${stage}\\mode.txt" w\nFileWrite $0 "${mode}"\nFileClose $0\n!insertmacro fourpsCallUpgrade prepare\nWriteINIStr "${status}" "${mode}" "Exit" "$R2"\nWriteINIStr "${status}" "${mode}" "Message" "$R3"\nCopyFiles /SILENT "$PLUGINSDIR\\fourps-result-prepare-$fourpsSequence.ini" "${lab}\\captured\\${mode}.ini"\n`;
source+='SectionEnd\n';await writeFile(path.join(lab,'report.nsi'),source);
const compiler=path.join(process.env.LOCALAPPDATA,'electron-builder/Cache/nsis-3.0.4.1/nsis-3.0.4.1-1mx3n/makensis.exe');
await run(compiler,['-WX','-INPUTCHARSET','UTF8','-V2',path.join(lab,'report.nsi')],{windowsHide:true,timeout:30000});
await run(path.join(lab,'report.exe'),[],{windowsHide:true,timeout:60000});
const ini=(await readFile(status)).toString('utf16le').replace(/^\uFEFF/,'');
const results=Object.fromEntries(ini.split(/\r?\n(?=\[)/).map(section=>{const [header,...lines]=section.split(/\r?\n/);return[header.slice(1,-1),Object.fromEntries(lines.filter(Boolean).map(line=>{const n=line.indexOf('=');return[line.slice(0,n),line.slice(n+1)];}))];}));
for(const mode of cases)assert.equal(results[mode].Exit,mode==='success'?'0':'2',mode);
assert.match(results.failure.Message,/não inicializou.*São José/);
for(const mode of ['missing','wrong','stale'])assert.match(results[mode].Message,/relatório válido/);
const result={passed:true,lab,guiChildWithoutConsole:true,unicodeErrorPreserved:true,missingRejected:true,wrongActionRejected:true,staleRejected:true,successAccepted:true,results};await writeFile(path.join(lab,'result.json'),JSON.stringify(result,null,2));console.log(JSON.stringify(result));
