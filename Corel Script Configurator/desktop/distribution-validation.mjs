import {mkdtempSync} from 'node:fs';
import path from 'node:path';
import {fileURLToPath,pathToFileURL} from 'node:url';
import {verifyOfficeRuntime,privateOfficeEnvironment} from './libreoffice-runtime.mjs';
import {runOffice} from './word-import.mjs';

async function probeOffice(result,temporaryRoot){
    const root=path.dirname(path.dirname(result.executable));
    try{await runOffice(result.executable,['-env:UserInstallation='+pathToFileURL(path.join(temporaryRoot,'office')).href,'--version'],{cwd:temporaryRoot,env:privateOfficeEnvironment(root)});}
    catch(cause){throw Error('O conversor privado não inicializou neste diretório. A instalação não foi validada. Detalhe: '+cause.message,{cause});}
}

// Entrada interna, sem janela, single-instance, PDF, impressão ou perfil real.
// Não recebe caminhos pela linha de comando: verifica somente o próprio pacote.
export async function validateDistribution({app,resourcesPath=process.resourcesPath,
    trustFile=fileURLToPath(new URL('../build/libreoffice/trust.json',import.meta.url)),
    verify=verifyOfficeRuntime,probe=probeOffice,temporary=()=>mkdtempSync(path.join(app.getPath('temp'),'4ps-verify-')),
    log=console.log,error=console.error,timeout=180000}={}){
    let timer;
    try{
        if(!app.isPackaged)throw Error('Verificação de instalação requer aplicativo empacotado.');
        const temporaryRoot=temporary();app.setPath('userData',temporaryRoot); // Antes do primeiro await; perfil do operador não é aberto.
        const checking=(async()=>{const result=await verify(path.resolve(resourcesPath,'../lo'),trustFile);await probe(result,temporaryRoot);return result;})();
        const result=await Promise.race([checking,new Promise((_,reject)=>{timer=setTimeout(()=>reject(Error('Tempo limite na verificação integral da instalação.')),timeout);})]);
        log(JSON.stringify({verified:true,version:result.version,files:result.files}));app.exit(0);
    }catch(failure){error('Instalação não validada: '+failure.message);app.exit(2);}
    finally{clearTimeout(timer);}
}
