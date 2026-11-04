import {execFile} from 'node:child_process';
import {promisify} from 'node:util';
import {readFile,mkdir,writeFile} from 'node:fs/promises';
import {createHash,randomUUID} from 'node:crypto';
import path from 'node:path';
import {pathToFileURL} from 'node:url';

const exec=promisify(execFile);
let running=false;
export async function runVirtualPrintLab(resourceFolder=path.resolve('build/print-lab')){
    if(running)throw new Error('Aguarde o teste virtual atual.');
    if(process.platform!=='win32')throw new Error('O teste virtual requer Windows.');
    running=true;
    try {
        const manifest=JSON.parse(await readFile(path.join(resourceFolder,'manifest.json'),'utf8'));
        if(manifest.schema!==1||manifest.allowedQueue!=='Microsoft Print to PDF')throw new Error('Manifesto de laboratório inválido.');
        for(const filename of ['4PrintShops.PrintLab.exe','pdfium.dll','synthetic.pdf']){
            const bytes=await readFile(path.join(resourceFolder,filename));
            if(createHash('sha256').update(bytes).digest('hex')!==manifest.files[filename])throw new Error('Recurso de laboratório incompatível.');
        }
        const folder=path.resolve('.desktop-test/hold-lab');
        await mkdir(folder,{recursive:true});
        const output=path.join(folder,'virtual-'+randomUUID()+'.pdf');
        const result=await exec(path.join(resourceFolder,'4PrintShops.PrintLab.exe'),[
            '--virtual-file',path.join(resourceFolder,'synthetic.pdf'),output
        ],{windowsHide:true,timeout:60000,maxBuffer:32768});
        const report={queue:manifest.allowedQueue,pdfium:manifest.pdfium,output,stdout:result.stdout.trim(),physicalDestinationAvailable:false};
        await writeFile(path.join(folder,'virtual-latest.json'),JSON.stringify(report,null,2)+'\n');
        return report;
    } finally {
        running=false;
    }
}
if(process.argv[1]&&import.meta.url===pathToFileURL(path.resolve(process.argv[1])).href){
    console.log(JSON.stringify(await runVirtualPrintLab(process.argv[2]?path.resolve(process.argv[2]):undefined),null,2));
}
