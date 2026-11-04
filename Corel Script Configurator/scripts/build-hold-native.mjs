import {execFile} from 'node:child_process';
import {promisify} from 'node:util';
import {mkdir,readFile,writeFile,cp} from 'node:fs/promises';
import {createHash} from 'node:crypto';
import path from 'node:path';
import {fileURLToPath} from 'node:url';

const root=fileURLToPath(new URL('../',import.meta.url));
const exec=promisify(execFile);
export async function ghostscriptTrust({configured=process.env.PRINT_GHOSTSCRIPT,programFiles=process.env.ProgramFiles}={}){
    if(configured==='none')return {path:'',sha256:''};
    if(configured&&!path.isAbsolute(configured))throw new Error('PRINT_GHOSTSCRIPT exige caminho absoluto ou none.');
    const file=configured||(programFiles?path.join(programFiles,'gs/gs10.01.2/bin/gswin64c.exe'):null);
    if(!file)return {path:'',sha256:''};
    try{return {path:file,sha256:createHash('sha256').update(await readFile(file)).digest('hex')};}
    catch(error){if(!configured&&error.code==='ENOENT')return {path:'',sha256:''};throw error;}
}
export async function buildHoldNative(){
    const folder=path.join(root,'build/hold-native'),lab=path.join(root,'build/print-lab');
    await mkdir(folder,{recursive:true});
    const ghostscript=await ghostscriptTrust(),gs=ghostscript.path,gsHash=ghostscript.sha256;
    if(!gs)console.log('Ghostscript não configurado. Impressão Windows disponível; diagnóstico PS/Hold bloqueado.');
    const trust=path.join(folder,'HoldBuildTrust.cs');
    await writeFile(trust,'namespace FourPrintShops.PrintConfig { internal static class HoldBuildTrust { internal const string Ghostscript=@"'+gs.replaceAll('"','""')+'"; internal const string GhostscriptSHA256="'+gsHash+'"; } }\n');
    const framework=path.join(process.env.SystemRoot,'Microsoft.NET/Framework64/v4.0.30319');
    const sources=['DevModeState.cs','PrintConfigNative.cs','PrintConfigHost.cs','HoldPrintTicket.cs','PdfiumVirtualLab.cs','PdfiumFixture.cs','KonicaPostScriptCapture.cs','PdfiumSyntheticGuard.cs','HoldExperimentalJob.cs','CaptureEnvironmentPolicy.cs','WindowsPrintJob.cs'].map(name=>path.join(root,'desktop/native',name));
    const output=path.join(folder,'4PrintShops.PrintConfig.exe');
    await exec(path.join(framework,'csc.exe'),['/nologo','/target:exe','/platform:x64','/optimize+','/define:HOLD_EXPERIMENTAL;WINDOWS_PRINT',
        '/main:FourPrintShops.PrintConfig.Host','/out:'+output,'/reference:System.Xml.Linq.dll','/reference:System.Web.Extensions.dll',
        '/reference:'+path.join(framework,'WPF/System.Printing.dll'),'/reference:'+path.join(framework,'WPF/ReachFramework.dll'),...sources,trust
    ],{windowsHide:true,timeout:30000});
    const files={};
    for(const name of ['pdfium.dll','synthetic.pdf','PDFium-LICENSE.txt','Liberation-LICENSE.txt','licenses'])await cp(path.join(lab,name),path.join(folder,name),{recursive:true});
    for(const name of ['4PrintShops.PrintConfig.exe','pdfium.dll','synthetic.pdf'])files[name]=createHash('sha256').update(await readFile(path.join(folder,name))).digest('hex');
    await writeFile(path.join(folder,'manifest.json'),JSON.stringify({schema:1,file:'4PrintShops.PrintConfig.exe',sha256:files['4PrintShops.PrintConfig.exe'],files,ghostscript:{path:gs,sha256:gsHash},physicalSendRequiresConfirmation:true})+'\n');
    console.log('Auxiliar experimental compilado. Build sem envio de trabalhos.');
}
if(process.argv[1]&&path.resolve(process.argv[1])===fileURLToPath(import.meta.url))await buildHoldNative();
