// Provisiona apenas identidade de teste. Leitura do registro não muda default/UserChoice.
import assert from 'node:assert/strict';
import {spawn,execFile} from 'node:child_process';
import {promisify} from 'node:util';
import {mkdtemp,readFile,writeFile,realpath} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import path from 'node:path';
import {PDFDocument} from '../public/vendor/pdf-lib/pdf-lib.js';
const exec=promisify(execFile),folder=await mkdtemp(path.join(tmpdir(),'4printshops-icons-installed-')),installation=path.join(folder,'installed');
const setup=path.resolve('dist-desktop/icon-test/4PrintShops-Icon-Test-1.0.0-experimental.7-Windows-x64-Setup.exe');
const ps=async script=>(await exec(path.join(process.env.SystemRoot,'System32','WindowsPowerShell','v1.0','powershell.exe'),['-NoProfile','-NonInteractive','-Command',script],{windowsHide:true,timeout:30000})).stdout.trim();
const quote=value=>"'"+value.replaceAll("'","''")+"'";
const snapshot=()=>ps("$r=[Microsoft.Win32.Registry]::CurrentUser;$out=@{};foreach($p in @('Software\\Classes\\.pdf','Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\FileExts\\.pdf\\UserChoice')){$k=$r.OpenSubKey($p);if($k){$out[$p]=@{};foreach($n in $k.GetValueNames()){$out[$p][$n]=$k.GetValue($n)};$k.Close()}};$out|ConvertTo-Json -Depth 5 -Compress");
const before=await snapshot();
assert.ok((await realpath(folder)).toLowerCase().startsWith((await realpath(tmpdir())).toLowerCase()+path.sep));
await new Promise((resolve,reject)=>{const p=spawn(setup,['/S','/D='+installation],{windowsHide:true,stdio:'ignore'});p.once('error',reject);p.once('exit',code=>code===0?resolve():reject(new Error('NSIS '+code)));});
assert.deepEqual(JSON.parse(await snapshot()),JSON.parse(before),'Instalador nao altera leitor padrao ou UserChoice');
const helper=await readFile(new URL('./helpers/windows-icon-resources.ps1',import.meta.url),'utf8');
console.log(await ps(`& {${helper}} -Installation ${quote(installation)} -Installer ${quote(setup)} -Output ${quote(folder)} -ProjectRoot ${quote(process.cwd())}`));
const pdf=await PDFDocument.create();for(let n=0;n<3;n++){const page=pdf.addPage([420,595]);page.drawText('4PrintShops - Teste de icones - Pagina '+(n+1),{x:25,y:550,size:13});}
const document=path.join(folder,'Conferencia visual São José.pdf');await writeFile(document,await pdf.save());
const exe=path.join(installation,'4PrintShops Icon Test.exe'),env={...process.env};delete env.ELECTRON_RUN_AS_NODE;
const app=spawn(exe,['--open-pdf',document],{env,stdio:'ignore',windowsHide:true,detached:true});await new Promise((resolve,reject)=>{app.once('spawn',resolve);app.once('error',reject);});app.unref();
await writeFile(path.join(folder,'result.json'),JSON.stringify({installation,setup,document,pid:app.pid,defaultPreserved:true,resourcesMatchedAt:[16,32],visualInspectionPending:true},null,2));
console.log('PASSOU: NSIS instalado sem mudar default; aplicativo iniciado com PDF sintetico para conferencia visual.');console.log('Artefatos: '+folder);
