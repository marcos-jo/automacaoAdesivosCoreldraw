import {execFile} from 'node:child_process';
import {promisify} from 'node:util';
import {mkdir,readFile,writeFile,cp,rename} from 'node:fs/promises';
import {createHash,randomUUID} from 'node:crypto';
import path from 'node:path';
import {fileURLToPath,pathToFileURL} from 'node:url';

const exec=promisify(execFile);
const root=fileURLToPath(new URL('../',import.meta.url));
export async function buildPrintLab(){
    if(process.platform!=='win32')throw new Error('O laboratório PDFium requer Windows x64.');
    const folder=path.join(root,'build/print-lab');
    const pdfium=path.join(root,'build/pdfium');
    const lock=JSON.parse(await readFile(path.join(root,'build/pdfium-lock.json'),'utf8'));
    const archive=await readFile(path.join(pdfium,lock.asset));
    if(createHash('sha256').update(archive).digest('hex')!==lock.sha256)throw new Error('Prepare o PDFium fixado com scripts/prepare-pdfium.ps1.');
    await mkdir(folder,{recursive:true});
    const compiler=path.join(process.env.SystemRoot,'Microsoft.NET/Framework64/v4.0.30319/csc.exe');
    await exec(compiler,['/nologo','/target:exe','/platform:x64','/optimize+','/out:'+path.join(folder,'4PrintShops.PrintLab.exe'),
        path.join(root,'desktop/native/PdfiumVirtualLab.cs'),path.join(root,'desktop/native/PdfiumFixture.cs'),path.join(root,'desktop/native/CaptureEnvironmentPolicy.cs')],{windowsHide:true,timeout:60000});
    await cp(path.join(pdfium,'bin/pdfium.dll'),path.join(folder,'pdfium.dll'));
    await cp(path.join(pdfium,'LICENSE'),path.join(folder,'PDFium-LICENSE.txt'));
    await cp(path.join(pdfium,'licenses'),path.join(folder,'licenses'),{recursive:true});
    const fonts=path.join(root,'node_modules/pdfjs-dist/standard_fonts');
    await cp(path.join(fonts,'LICENSE_LIBERATION'),path.join(folder,'Liberation-LICENSE.txt'));
    const synthetic=path.join(folder,'synthetic-'+randomUUID()+'.pdf');
    await exec(path.join(folder,'4PrintShops.PrintLab.exe'),['--create-synthetic',path.join(fonts,'LiberationSans-Regular.ttf'),synthetic],{windowsHide:true,timeout:30000});
    await rename(synthetic,path.join(folder,'synthetic.pdf'));
    const files={};
    for(const name of ['4PrintShops.PrintLab.exe','pdfium.dll','synthetic.pdf']){
        files[name]=createHash('sha256').update(await readFile(path.join(folder,name))).digest('hex');
    }
    await writeFile(path.join(folder,'manifest.json'),JSON.stringify({schema:1,pdfium:lock.tag,allowedQueue:'Microsoft Print to PDF',files})+'\n');
    console.log('Laboratório PDFium compilado com destino virtual fixo.');
}
if(process.argv[1]&&import.meta.url===pathToFileURL(path.resolve(process.argv[1])).href)await buildPrintLab();
