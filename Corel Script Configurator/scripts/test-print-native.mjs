// Testes de estado e política. Não abrem impressoras nem criam trabalhos.
import {execFile} from 'node:child_process';
import {promisify} from 'node:util';
import {mkdir} from 'node:fs/promises';
import path from 'node:path';
import {fileURLToPath} from 'node:url';

const exec=promisify(execFile);
const root=fileURLToPath(new URL('../',import.meta.url));
if(process.platform!=='win32')throw new Error('Os testes C# requerem o compilador do Windows.');
const framework=path.join(process.env.SystemRoot,'Microsoft.NET/Framework64/v4.0.30319');
const folder=path.join(root,'.desktop-test/hold-lab');
await mkdir(folder,{recursive:true});
for(const target of ['PrintConfigStateTests','HoldPrintTicketTests','CaptureEnvironmentPolicyTests','WindowsPrintPolicyTests']){
    const names=target==='CaptureEnvironmentPolicyTests'?['CaptureEnvironmentPolicy.cs']:target==='PrintConfigStateTests'?['DevModeState.cs']:target==='WindowsPrintPolicyTests'?['DevModeState.cs','PrintConfigNative.cs','HoldPrintTicket.cs','WindowsPrintJob.cs']:['DevModeState.cs','PrintConfigNative.cs','HoldPrintTicket.cs'];
    const output=path.join(folder,target+'.exe');
    const sources=names.map(name=>path.join(root,'desktop/native',name));
    sources.push(path.join(root,'tests/native',target+'.cs'));
    await exec(path.join(framework,'csc.exe'),[
        '/nologo','/target:exe','/platform:x64','/main:'+target,'/out:'+output,
        '/reference:System.Xml.Linq.dll','/reference:System.Web.Extensions.dll',
        '/reference:'+path.join(framework,'WPF/System.Printing.dll'),
        '/reference:'+path.join(framework,'WPF/ReachFramework.dll'),...sources
    ],{windowsHide:true,timeout:30000});
    const result=await exec(output,[],{windowsHide:true,timeout:10000});
    console.log(result.stdout.trim());
}
