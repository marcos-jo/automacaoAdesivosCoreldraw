import {execFile} from 'node:child_process';
import {promisify} from 'node:util';
import {mkdir,readFile,writeFile} from 'node:fs/promises';
import {createHash} from 'node:crypto';
import path from 'node:path';
import {fileURLToPath,pathToFileURL} from 'node:url';

const exec=promisify(execFile),root=fileURLToPath(new URL('../',import.meta.url));
export async function buildPrintHelper(){
    if(process.platform!=='win32')throw new Error('O auxiliar de configuração requer compilação no Windows.');
    const folder=path.join(root,'build/native'),output=path.join(folder,'4PrintShops.PrintConfig.exe');
    await mkdir(folder,{recursive:true});
    const compiler=path.join(process.env.SystemRoot,'Microsoft.NET/Framework64/v4.0.30319/csc.exe');
    const sources=['DevModeState.cs','PrintConfigNative.cs','PrintConfigHost.cs'].map(name=>path.join(root,'desktop/native',name));
    await exec(compiler,['/nologo','/target:exe','/platform:x64','/optimize+','/out:'+output,'/reference:System.Web.Extensions.dll',...sources],{windowsHide:true,timeout:60000});
    const sha256=createHash('sha256').update(await readFile(output)).digest('hex');
    await writeFile(path.join(folder,'manifest.json'),JSON.stringify({schema:1,file:'4PrintShops.PrintConfig.exe',sha256})+'\n');
    console.log('Auxiliar de configuração compilado; sem API de envio de trabalhos.');
    return output;
}
if(process.argv[1]&&import.meta.url===pathToFileURL(path.resolve(process.argv[1])).href)await buildPrintHelper();
