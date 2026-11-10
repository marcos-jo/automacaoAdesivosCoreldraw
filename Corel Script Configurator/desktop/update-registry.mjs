// Used only by the isolated installer process. No UserChoice/default writes.
import {spawn} from 'node:child_process';
import path from 'node:path';
const script=String.raw`
$ErrorActionPreference='Stop'
[Console]::InputEncoding=[Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
$r=[Console]::In.ReadToEnd()|ConvertFrom-Json
$base=[Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser,[Microsoft.Win32.RegistryView]::Registry64)
function ReadTree($name){
 $key=$base.OpenSubKey($name,$false);if($null -eq $key){return @{path=$name;exists=$false}}
 try{$values=@();foreach($n in $key.GetValueNames()){$kind=$key.GetValueKind($n);$v=$key.GetValue($n,$null,[Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames);if($kind -eq 'Binary'){$v=[Convert]::ToBase64String($v)}elseif($kind -eq 'DWord' -or $kind -eq 'QWord'){$v=[string]$v};$values+=@{name=$n;kind=[string]$kind;data=$v}};$children=@();foreach($n in $key.GetSubKeyNames()){$children+=ReadTree ($name+'\'+$n)};return @{path=$name;exists=$true;values=$values;children=$children}}finally{$key.Dispose()}
}
function WriteTree($tree,$root){
 if($tree.path -ne $root -and -not $tree.path.StartsWith($root+'\',[StringComparison]::OrdinalIgnoreCase)){throw 'Registry subtree scope refused'}
 if($tree.exists){$key=$base.CreateSubKey($tree.path);try{foreach($v in $tree.values){$kind=[Enum]::Parse([Microsoft.Win32.RegistryValueKind],[string]$v.kind);$value=$v.data;if($kind -eq 'Binary'){$value=[Convert]::FromBase64String($value)}elseif($kind -eq 'DWord'){$value=[int]::Parse($value)}elseif($kind -eq 'QWord'){$value=[long]::Parse($value)}elseif($kind -eq 'MultiString'){$value=[string[]]@($value)};$key.SetValue($v.name,$value,$kind)}}finally{$key.Dispose()};foreach($child in $tree.children){WriteTree $child $root}}
}
function CheckPath($p){if($p -notin @($r.allowedTrees)){throw 'Registry scope refused'}}
try{
 if($r.op -eq 'snapshot'){
  $trees=@();foreach($p in $r.allowedTrees){$trees+=ReadTree $p};$values=@();foreach($v in $r.allowedValues){$key=$base.OpenSubKey($v.path,$false);$parentPath=$v.path.Substring(0,$v.path.LastIndexOf('\'));$parent=$base.OpenSubKey($parentPath,$false);try{$has=$null -ne $key -and $key.GetValueNames() -contains $v.name;$values+=@{path=$v.path;name=$v.name;exists=$has;keyExists=($null -ne $key);parentExists=($null -ne $parent);data=$(if($has){[string]$key.GetValue($v.name)}else{$null})}}finally{if($key){$key.Dispose()};if($parent){$parent.Dispose()}}}
  $choice=ReadTree 'Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.pdf\UserChoice'
  $pdf=$base.OpenSubKey('Software\Classes\.pdf',$false);try{$hasDefault=$null -ne $pdf -and $pdf.GetValueNames() -contains ''; $pdfDefault=@{exists=$hasDefault;data=$(if($hasDefault){$pdf.GetValue('',$null,[Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)}else{$null});kind=$(if($hasDefault){[string]$pdf.GetValueKind('')}else{$null})}}finally{if($pdf){$pdf.Dispose()}}
  @{trees=$trees;values=$values;userChoice=$choice;pdfDefault=$pdfDefault;environment=@{desktop=[Environment]::GetFolderPath('DesktopDirectory');programs=[Environment]::GetFolderPath('Programs');local=[Environment]::GetFolderPath('LocalApplicationData')}}|ConvertTo-Json -Depth 40 -Compress
 }elseif($r.op -eq 'restore'){
  foreach($tree in $r.snapshot.trees){CheckPath $tree.path;$base.DeleteSubKeyTree($tree.path,$false);WriteTree $tree $tree.path}
  foreach($v in $r.snapshot.values){$permitted=@($r.allowedValues|Where-Object {$_.path -eq $v.path -and $_.name -eq $v.name});if($permitted.Count -ne 1){throw 'Registry value scope refused'};if($v.exists){$key=$base.CreateSubKey($v.path);try{$key.SetValue($v.name,[string]$v.data,[Microsoft.Win32.RegistryValueKind]::String)}finally{$key.Dispose()}}else{$key=$base.OpenSubKey($v.path,$true);$empty=$false;if($key){try{$key.DeleteValue($v.name,$false);$empty=($key.GetValueNames().Count -eq 0 -and $key.GetSubKeyNames().Count -eq 0)}finally{$key.Dispose()}};if(-not $v.keyExists -and $empty){$base.DeleteSubKeyTree($v.path,$false)};if(-not $v.parentExists -and $v.path.EndsWith('\.pdf\OpenWithProgids')){$p=$v.path.Substring(0,$v.path.LastIndexOf('\'));$parent=$base.OpenSubKey($p,$false);if($parent){$empty=($parent.GetValueNames().Count -eq 0 -and $parent.GetSubKeyNames().Count -eq 0);$parent.Dispose();if($empty){$base.DeleteSubKeyTree($p,$false)}}}}}
  @{restored=$true}|ConvertTo-Json -Compress
 }elseif($r.op -eq 'busy'){
  $busy=@();foreach($p in Get-Process){if($p.Id -eq $r.pid){continue};try{$file=$p.MainModule.FileName;if($file -and $file.StartsWith($r.target+'\',[StringComparison]::OrdinalIgnoreCase)){$busy+=@{pid=$p.Id;name=$p.ProcessName}}}catch{if($p.ProcessName -eq $r.processName){$busy+=@{pid=$p.Id;name=$p.ProcessName}}}};@{processes=$busy}|ConvertTo-Json -Depth 5 -Compress
 }else{throw 'Unknown registry action'}
}finally{$base.Dispose()}
`;
export function registryScope(identity){
    if(!/^br\.com\.4printshops\.[a-z0-9.-]+$/i.test(identity.appId)||!/^[a-f0-9-]{36}$/i.test(identity.guid))throw Error('Identidade de distribuição inválida.');
    const lab=identity.lab===true;
    if(lab&&!identity.appId.startsWith('br.com.4printshops.upgradelab.'))throw Error('Identidade de laboratório inválida.');
    const prefix=lab?'Software\\4PrintShops\\UpgradeLab\\'+identity.guid:'Software';
    const install=lab?prefix+'\\Install':'Software\\'+identity.guid;
    const uninstall=lab?prefix+'\\Uninstall':'Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\'+identity.guid;
    const classes=lab?prefix+'\\Classes':'Software\\Classes';
    const capabilities=lab?prefix+'\\Capabilities':'Software\\'+identity.appId+'\\Capabilities';
    const registered=lab?prefix+'\\RegisteredApplications':'Software\\RegisteredApplications';
    return {install,uninstall,classes,capabilities,registered,
        allowedTrees:[install,uninstall,classes+'\\'+identity.appId+'.PDF',classes+'\\Applications\\'+identity.exe,capabilities],
        allowedValues:[{path:classes+'\\.pdf\\OpenWithProgids',name:identity.appId+'.PDF'},{path:registered,name:identity.appId}]};
}
export function registryRequest(identity,request){
    const scope=registryScope(identity);
    return new Promise((resolve,reject)=>{
        const exe=path.join(process.env.SystemRoot||'C:\\Windows','System32/WindowsPowerShell/v1.0/powershell.exe');
        const child=spawn(exe,['-NoProfile','-NonInteractive','-Command',script],{windowsHide:true,shell:false,stdio:['pipe','pipe','pipe']});
        let out='',err='';const timer=setTimeout(()=>{child.kill();reject(Error('Tempo limite na operação de registro do instalador.'));},45000);
        child.stdout.on('data',d=>out+=d);child.stderr.on('data',d=>err+=d);
        child.once('error',e=>{clearTimeout(timer);reject(e);});
        child.once('close',code=>{clearTimeout(timer);if(code!==0)return reject(Error('Não foi possível conferir/restaurar o registro próprio: '+err.slice(0,1200)));try{resolve(JSON.parse(out.replace(/^\uFEFF/,'')));}catch(e){reject(Error('Resposta inválida do registro: '+e.message));}});
        child.stdin.end(JSON.stringify({...request,...scope}));
    });
}
