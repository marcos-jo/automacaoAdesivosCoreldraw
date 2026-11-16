import {readFile,lstat,realpath} from 'node:fs/promises';
import {hostname,tmpdir} from 'node:os';
import {fileURLToPath} from 'node:url';
import path from 'node:path';
import {spawn} from 'node:child_process';

export const residueName='lo/program/opencl/.~lock.cl-test.ods#';
export function validateQuarantine(entries=[],unmanaged=[],next=[]){
    if(!Array.isArray(entries)||entries.length>1)throw Error('Quarentena inválida.');
    for(const e of entries){const old=unmanaged.find(f=>f.name===residueName);
        if(e.name!==residueName||!old||e.sha256!==old.sha256||e.size!==old.size||!Number.isInteger(e.size)||e.size<1||e.size>4096||!/^[a-f0-9]{64}$/.test(e.sha256)||next.some(f=>f.name.toLowerCase()===e.name.toLowerCase()))throw Error('Entrada de quarentena fora do escopo.');}
}
export async function recognizedResidue(target,entry){
    if(entry.name!==residueName||entry.size<1||entry.size>4096)return false;
    const bytes=await readFile(path.join(target,entry.name)),text=bytes.toString('utf8');
    if(!Buffer.from(text,'utf8').equals(bytes)||text.includes('\0'))return false;
    const fields=text.trim().split(',');
    if(fields.length!==5||fields[0]!==''||!fields[1]||fields[2].toLowerCase()!==hostname().toLowerCase()||!/^\d{2}\.\d{2}\.\d{4} \d{2}:\d{2}$/.test(fields[3])||!fields[4].endsWith(';'))return false;
    let profile;try{profile=fileURLToPath(fields[4].slice(0,-1));}catch{return false;}
    if(path.basename(profile).toLowerCase()!=='profile'||!/^4printshops-word-[a-z0-9]+$/i.test(path.basename(path.dirname(profile))))return false;
    // Windows TEMP can use an 8.3 alias while the lock records the long path.
    try{if((await realpath(path.dirname(path.dirname(profile)))).toLowerCase()!==(await realpath(tmpdir())).toLowerCase())return false;}catch(e){if(e.code==='ENOENT')return false;throw e;}
    try{await lstat(profile);return false;}catch(e){if(e.code!=='ENOENT')throw e;}
    return true;
}

// FileShare.Delete refuses existing/future readers and writers while allowing
// our atomic move with its read handle held. File.Move never replaces a target.
const script=String.raw`
$ErrorActionPreference='Stop'
$r=[Console]::In.ReadToEnd()|ConvertFrom-Json
function NoLinks($p){while($p){if(Test-Path -LiteralPath $p){if(([IO.File]::GetAttributes($p) -band [IO.FileAttributes]::ReparsePoint) -ne 0){throw 'Link recusado na quarentena'}};$next=[IO.Path]::GetDirectoryName($p);if($next -eq $p){break};$p=$next}}
$target=[IO.Path]::GetFullPath($r.target);$folder=[IO.Path]::GetFullPath($r.folder)
if([IO.Path]::GetDirectoryName($target) -ne [IO.Path]::GetDirectoryName($folder) -or [IO.Path]::GetFileName($folder) -notmatch '^\.4ps-upgrade-[a-zA-Z0-9]+$' -or $r.sha256 -notmatch '^[a-f0-9]{64}$'){throw 'Escopo da quarentena inválido'}
if(@(Get-Process|Where-Object {$_.ProcessName -match '^(soffice(\.bin)?|opencltest|oosplash)$'}).Count){throw 'Feche o LibreOffice antes de tratar resíduos'}
$original=Join-Path $target 'lo\program\opencl\.~lock.cl-test.ods#'
$qdir=Join-Path $folder 'quarantine';$stored=Join-Path $qdir 'libreoffice-opencl-lock'
NoLinks $original;NoLinks $stored
if($r.restore){$src=$stored;$dst=$original}else{$src=$original;$dst=$stored}
if(-not (Test-Path -LiteralPath $src)){
 if(-not (Test-Path -LiteralPath $dst)){throw 'Resíduo ausente nos dois locais'}
 $src=$dst;$already=$true
}else{$already=$false;if(Test-Path -LiteralPath $dst){throw 'Colisão na quarentena; nenhum arquivo substituído'}}
$stream=[IO.File]::Open($src,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::Delete)
try{
 $sha=[Security.Cryptography.SHA256]::Create();try{$hash=[BitConverter]::ToString($sha.ComputeHash($stream)).Replace('-','').ToLowerInvariant()}finally{$sha.Dispose()}
 if($hash -ne $r.sha256){throw 'Conteúdo do resíduo alterado'}
 if(-not $already){[IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($dst))|Out-Null;NoLinks $dst;[IO.File]::Move($src,$dst)}
}finally{$stream.Dispose()}
@{preserved=$true}|ConvertTo-Json -Compress
`;
export async function moveResidue({target,folder,entry,restore=false}){
    validateQuarantine([entry],[entry],[]);
    // Reject redirected roots before passing any path to the native helper.
    for(const p of [target,folder]){const s=await lstat(p);if(!s.isDirectory()||s.isSymbolicLink())throw Error('Raiz de quarentena inválida.');}
    return new Promise((resolve,reject)=>{
        const exe=path.join(process.env.SystemRoot||'C:\\Windows','System32/WindowsPowerShell/v1.0/powershell.exe');
        const child=spawn(exe,['-NoProfile','-NonInteractive','-Command',script],{windowsHide:true,shell:false,stdio:['pipe','pipe','pipe']});let out='',err='';
        // Do not kill a mover mid-operation; caller awaits completion and the
        // recorded intent makes recovery independent of a result acknowledgement.
        child.stdout.on('data',b=>out+=b);child.stderr.on('data',b=>err+=b);child.once('error',reject);child.once('close',code=>{if(code!==0)return reject(Error('Quarentena recusada: '+err.slice(-1200)));try{const r=JSON.parse(out.replace(/^\uFEFF/,''));if(r.preserved!==true)throw Error('Resultado inválido');resolve(r);}catch(e){reject(e);}});
        child.stdin.on('error',()=>{});child.stdin.end(JSON.stringify({target,folder,sha256:entry.sha256,restore}));
    });
}
