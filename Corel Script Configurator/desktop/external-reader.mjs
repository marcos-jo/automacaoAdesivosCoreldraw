import {mkdir,open,lstat} from 'node:fs/promises';
import path from 'node:path';
import {randomUUID} from 'node:crypto';
import {spawn} from 'node:child_process';
import {pdfName,pdfData} from './security.mjs';

export async function findMicrosoftEdge(){
    const bases=[process.env['ProgramFiles(x86)'],process.env.ProgramFiles,process.env.LOCALAPPDATA].filter(Boolean);
    for(const base of bases){
        const candidate=path.join(base,'Microsoft/Edge/Application/msedge.exe');
        try {
            const info=await lstat(candidate);
            if(info.isFile()&&!info.isSymbolicLink())return candidate;
        } catch {}
    }
    return null;
}

async function launchReader(executable,file){
    await new Promise((resolve,reject)=>{
        const child=spawn(executable,[file],{shell:false,windowsHide:false,detached:true,stdio:'ignore'});
        child.once('error',()=>reject(new Error('Não foi possível iniciar o leitor escolhido.')));
        child.once('spawn',()=>{child.unref();resolve();});
    });
}

export async function openExternalCopy(window,dialog,bytes,name,tempRoot,{findEdge=findMicrosoftEdge,launch=launchReader}={}){
    pdfData(bytes);
    pdfName(name);
    const edge=await findEdge();
    const buttons=edge?['Microsoft Edge','Escolher outro leitor','Cancelar']:['Escolher leitor','Cancelar'];
    const choice=await dialog.showMessageBox(window,{
        type:'question',title:'Abrir cópia em leitor externo',
        message:'Escolha o leitor para abrir uma cópia do documento atual.',
        detail:'A cópia inclui as edições atuais. As alterações continuam pendentes de exportação no 4PrintShops.',
        buttons,defaultId:0,cancelId:buttons.length-1,noLink:true
    });
    if(choice.response===buttons.length-1)return {opened:false};
    if(!Number.isInteger(choice.response)||choice.response<0||choice.response>=buttons.length)throw new Error('Escolha de leitor inválida.');
    let executable=edge;
    if(!edge||choice.response!==0){
        const selected=await dialog.showOpenDialog(window,{
            title:'Escolher leitor PDF instalado',properties:['openFile'],filters:[{name:'Aplicativos Windows',extensions:['exe']}]
        });
        if(selected.canceled)return {opened:false};
        executable=selected.filePaths?.[0];
        if(typeof executable!=='string'||!path.isAbsolute(executable)||executable.startsWith('\\\\')||path.extname(executable).toLowerCase()!=='.exe'||/^4PrintShops.*\.exe$/i.test(path.basename(executable)))throw new Error('Escolha um leitor PDF externo instalado neste computador.');
        const info=await lstat(executable);
        if(!info.isFile()||info.isSymbolicLink())throw new Error('O leitor selecionado é inválido.');
    }
    const folder=path.join(tempRoot,'4PrintShops-external-copies',randomUUID());
    await mkdir(folder,{recursive:true});
    const filename=path.join(folder,name);
    const handle=await open(filename,'wx');
    try {await handle.writeFile(bytes);await handle.sync();}
    finally {await handle.close();}
    await launch(executable,filename);
    return {opened:true,reader:edge&&choice.response===0?'Microsoft Edge':path.basename(executable)};
}
