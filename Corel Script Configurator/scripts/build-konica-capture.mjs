import {execFile} from 'node:child_process';
import {promisify} from 'node:util';
import {mkdir,readFile,cp} from 'node:fs/promises';
import {createHash} from 'node:crypto';
import path from 'node:path';
import {fileURLToPath} from 'node:url';

const exec=promisify(execFile);
const root=fileURLToPath(new URL('../',import.meta.url));
const folder=path.join(root,'build/hold-capture');
const lab=path.join(root,'build/print-lab');
const manifest=JSON.parse(await readFile(path.join(lab,'manifest.json'),'utf8'));
for(const file of ['pdfium.dll','synthetic.pdf']){
    if(createHash('sha256').update(await readFile(path.join(lab,file))).digest('hex')!==manifest.files[file])throw new Error('Recurso do laboratório anterior incompatível.');
}
await mkdir(folder,{recursive:true});
const framework=path.join(process.env.SystemRoot,'Microsoft.NET/Framework64/v4.0.30319');
const sources=['DevModeState.cs','PrintConfigNative.cs','HoldPrintTicket.cs','PdfiumVirtualLab.cs','PdfiumFixture.cs','KonicaPostScriptCapture.cs','CaptureEnvironmentPolicy.cs'].map(name=>path.join(root,'desktop/native',name));
await exec(path.join(framework,'csc.exe'),['/nologo','/target:exe','/platform:x64','/optimize+',
    '/main:FourPrintShops.VirtualPrintLab.KonicaCaptureHost','/out:'+path.join(folder,'4PrintShops.KonicaCapture.exe'),
    '/reference:System.Xml.Linq.dll','/reference:System.Web.Extensions.dll',
    '/reference:'+path.join(framework,'WPF/System.Printing.dll'),'/reference:'+path.join(framework,'WPF/ReachFramework.dll'),...sources
],{windowsHide:true,timeout:30000});
for(const file of ['pdfium.dll','synthetic.pdf','PDFium-LICENSE.txt','Liberation-LICENSE.txt','licenses'])await cp(path.join(lab,file),path.join(folder,file),{recursive:true});
console.log('Capturador isolado compilado. Destino fixo FILE:, sem endpoint de envio físico.');
