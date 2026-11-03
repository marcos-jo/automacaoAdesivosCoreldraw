import {execFile} from 'node:child_process';
import {promisify} from 'node:util';
import path from 'node:path';
import {parseDn} from 'builder-util-runtime';

// O verificador 6.8.11 pode aceitar quando PowerShell/ConvertTo-Json faltam.
// Esta prova adicional deve passar ANTES dele; nunca confiar em PATH/shell.
export async function verifyInstallerSignature(publishers,file,{execute=promisify(execFile),systemRoot=process.env.SystemRoot||'C:\\Windows'}={}) {
    try{
        if(!Array.isArray(publishers)||!publishers.length||publishers.some(p=>typeof p!=='string'||!p.trim()))return 'Identidade de editor ausente.';
        const powershell=path.join(systemRoot,'System32','WindowsPowerShell','v1.0','powershell.exe');
        const literal="'"+file.replaceAll("'","''")+"'";
        const script=`[Console]::OutputEncoding=[System.Text.UTF8Encoding]::new($false);$s=Get-AuthenticodeSignature -LiteralPath ${literal};[pscustomobject]@{status=[int]$s.Status;embedded=($s.SignatureType -eq 'Authenticode');subject=$s.SignerCertificate.Subject;path=$s.Path}|ConvertTo-Json -Compress`;
        const {stdout}=await execute(powershell,['-NoLogo','-NoProfile','-NonInteractive','-Command',script],{windowsHide:true,timeout:20000,maxBuffer:1024*1024,
            env:{...process.env,PSModulePath:path.join(systemRoot,'System32','WindowsPowerShell','v1.0','Modules')}});
        const result=JSON.parse(stdout.replace(/^\uFEFF/,''));
        if(result.status!==0||result.embedded!==true||typeof result.subject!=='string')return 'Assinatura Authenticode incorporada ausente, inválida ou não confiável.';
        if(typeof result.path!=='string'||path.resolve(result.path).toLowerCase()!==path.resolve(file).toLowerCase())return 'A assinatura não corresponde ao instalador solicitado.';
        const subject=parseDn(result.subject);
        const matches=publishers.some(p=>{const expected=parseDn(p);return expected.size?[...expected].every(([key,value])=>subject.get(key)===value):subject.get('CN')===p;});
        return matches?null:'Assinatura pertence a outro editor.';
    }catch{return 'Não foi possível comprovar a assinatura do instalador. Instalação bloqueada.';}
}
