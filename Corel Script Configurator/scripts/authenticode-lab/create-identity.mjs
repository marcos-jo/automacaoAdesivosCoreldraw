import {execFileSync} from 'node:child_process';
import {mkdirSync,existsSync,writeFileSync} from 'node:fs';
import path from 'node:path';
import guard from './guest-guard.cjs';

// Executar SOMENTE no guest descartável; não exporta chave privada/PFX/senha.
guard.assertGuest(process.env.PRINTSHOPS_VM_MANIFEST,process.argv.includes('--confirm-disposable-guest'));
const folder=path.join(process.env.LOCALAPPDATA,'4PrintShops Authenticode Lab');
const output=path.join(folder,'identities.json');
if(existsSync(output))throw new Error('Identidades já existem. Reutilize-as nos dois builds ou restaure o snapshot; não gerar outra chave silenciosamente.');
mkdirSync(folder,{recursive:true});
const literal="'"+folder.replaceAll("'","''")+"'";
const script=`$ErrorActionPreference='Stop';[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false);$folder=${literal};$result=@();
foreach($item in @(@{id='primary';name='4PrintShops Update Lab (TEST ONLY)'},@{id='foreign';name='4PrintShops Other Publisher (TEST ONLY)'})){
  $cert=New-SelfSignedCertificate -Type CodeSigningCert -Subject ('CN='+$item.name) -FriendlyName $item.name -CertStoreLocation 'Cert:\\CurrentUser\\My' -KeyAlgorithm RSA -KeyLength 3072 -HashAlgorithm SHA256 -KeyExportPolicy NonExportable -NotAfter (Get-Date).AddDays(30);
  $file=Join-Path $folder ($item.id+'.cer');Export-Certificate -Cert $cert -FilePath $file | Out-Null;
  Import-Certificate -FilePath $file -CertStoreLocation 'Cert:\\CurrentUser\\Root' | Out-Null;
  Import-Certificate -FilePath $file -CertStoreLocation 'Cert:\\CurrentUser\\TrustedPublisher' | Out-Null;
  $result+=@{id=$item.id;publisher=$item.name;thumbprint=$cert.Thumbprint;subject=$cert.Subject;expires=$cert.NotAfter.ToString('o')};
};ConvertTo-Json -InputObject $result -Compress`;
const data=JSON.parse(execFileSync(path.join(process.env.SystemRoot,'System32/WindowsPowerShell/v1.0/powershell.exe'),['-NoProfile','-Command',script],{encoding:'utf8',windowsHide:true,timeout:120000}));
writeFileSync(output,JSON.stringify(data,null,2)+'\n');
console.log('Identidades públicas registradas em '+output+'. Chaves NÃO exportáveis permanecem somente no perfil do guest. Descartar a VM ao terminar.');
