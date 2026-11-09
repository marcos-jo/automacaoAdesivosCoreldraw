import {lstat} from 'node:fs/promises';
import path from 'node:path';
import {convertWord,readImportFile} from './word-import.mjs';

export const IMPORT_TOTAL=512*1024*1024;
export async function readImportedFiles(paths,{signal,convert=convertWord,wordOptions={},onSelection=()=>{},onProgress=()=>{},onEntry=()=>{}}={}) {
    if(!Array.isArray(paths)||!paths.length||paths.length>128)throw Error('Selecione de 1 a 128 arquivos.');
    if(paths.some(file=>typeof file!=='string'||!path.isAbsolute(file)))throw Error('Seleção inválida.');
    if(signal?.aborted)throw Error('Importação cancelada.');
    await onSelection(paths.map((file,index)=>({index,name:path.basename(file),type:path.extname(file).slice(1).toUpperCase()})));
    let inputTotal=0;
    for(const [index,file] of paths.entries()) {
        if(signal?.aborted)throw Error('Importação cancelada.');
        try{
            if(! /\.(pdf|jpe?g|png|docx?)$/i.test(file))throw Error('Selecione PDF, JPG/JPEG, PNG, DOCX ou DOC.');
            const info=await lstat(file);
            if(!info.isFile()||info.isSymbolicLink())throw Error('Selecione somente arquivos comuns.');
            inputTotal+=info.size;
            if(inputTotal>IMPORT_TOTAL)throw Error('A seleção ultrapassa o limite de 512 MiB.');
        }catch(error){if(!signal?.aborted)await onProgress(index,'failed');throw error;}
    }
    const result=[];let total=0;
    for(const [index,file] of paths.entries()) {
        if(signal?.aborted)throw Error('Importação cancelada.');
        try {
            let entry;
            await onProgress(index,/\.docx?$/i.test(file)?'word':/\.pdf$/i.test(file)?'pdf':'image');
            if(/\.docx?$/i.test(file)) {
                const converted=await convert(file,{...wordOptions,signal});
                entry={name:converted.name,bytes:converted.bytes,wordReport:{...converted.report,originalName:path.basename(file)}};
            }else entry={name:path.basename(file),bytes:new Uint8Array(await readImportFile(file,IMPORT_TOTAL))};
            total+=entry.bytes.length;
            if(total>IMPORT_TOTAL)throw Error('Os arquivos preparados ultrapassam 512 MiB. Divida a seleção.');
            if(signal?.aborted)throw Error('Importação cancelada.');
            await onEntry(index,entry);
            result.push(entry);
        }catch(error){if(!signal?.aborted)await onProgress(index,'failed');throw Error(path.basename(file)+' — '+error.message);}
    }
    if(signal?.aborted)throw Error('Importação cancelada.');
    return result;
}

export async function openImportedFiles(window,dialog,multiple=false,options={}) {
    if(typeof multiple!=='boolean')throw Error('Seleção inválida.');
    const choice=await dialog.showOpenDialog(window,{title:'Importar arquivos',filters:[
        {name:'PDF, imagens e Word',extensions:['pdf','jpg','jpeg','png','docx','doc']},
        {name:'Documentos PDF',extensions:['pdf']},{name:'Imagens',extensions:['jpg','jpeg','png']},{name:'Documentos Word',extensions:['docx','doc']}
    ],properties:multiple?['openFile','multiSelections']:['openFile']});
    if(choice.canceled)return [];
    if(!multiple&&choice.filePaths?.length!==1)throw Error('Selecione somente um arquivo.');
    return readImportedFiles(choice.filePaths,options);
}
