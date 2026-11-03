// Barreiras contra uso acidental no host; não substituem o isolamento da VM.
const {execFileSync}=require('node:child_process');
const {readFileSync}=require('node:fs');
const path=require('node:path');

function inspectMachine(){
    const script="$ErrorActionPreference='Stop';$id=(Get-ItemProperty -LiteralPath 'HKLM:\\SOFTWARE\\Microsoft\\Cryptography').MachineGuid;$sha=[System.Security.Cryptography.SHA256]::Create();$hash=[BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($id))).Replace('-','').ToLowerInvariant();$c=Get-CimInstance Win32_ComputerSystem;[pscustomobject]@{fingerprint=$hash;model=$c.Model;manufacturer=$c.Manufacturer}|ConvertTo-Json -Compress";
    return JSON.parse(execFileSync(path.join(process.env.SystemRoot||'C:\\Windows','System32/WindowsPowerShell/v1.0/powershell.exe'),['-NoProfile','-NonInteractive','-Command',script],{encoding:'utf8',windowsHide:true,timeout:15000}));
}
function validateGuest(manifest,machine,confirmed){
    if(confirmed!==true)throw new Error('Confirme explicitamente que esta é uma VM Windows descartável.');
    if(manifest?.schema!=='4printshops-authenticode-kit-v1'||!/^([a-f0-9]{64})$/.test(manifest.hostFingerprint||''))throw new Error('Manifesto do kit ausente ou inválido.');
    if(!/^[a-f0-9]{64}$/.test(machine?.fingerprint||''))throw new Error('Não foi possível identificar a máquina; operação bloqueada.');
    if(machine.fingerprint===manifest.hostFingerprint)throw new Error('HOST PRINCIPAL: operação de laboratório bloqueada.');
    if(!/virtual machine|virtualbox|vmware|kvm|qemu|parallels|hvm domu/i.test(machine.model||''))throw new Error('Modelo de VM não reconhecido; não contornar esta barreira.');
    return machine;
}
function assertGuest(manifestPath,confirmed){
    if(!confirmed)throw new Error('Falta confirmação explícita de VM descartável.');
    if(!manifestPath)throw new Error('Informe o manifesto do kit criado no host.');
    return validateGuest(JSON.parse(readFileSync(manifestPath,'utf8')),inspectMachine(),confirmed);
}
module.exports={inspectMachine,validateGuest,assertGuest};
