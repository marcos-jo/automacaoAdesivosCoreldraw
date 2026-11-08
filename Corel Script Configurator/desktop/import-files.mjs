import {lstat} from 'node:fs/promises';
import path from 'node:path';
import {convertWord,readImportFile} from './word-import.mjs';

export const IMPORT_TOTAL=512*1024*1024;
export async function readImportedFiles(paths,{signal,convert=convertWord,wordOptions={}}={}) {
    if(!Array.isArray(paths)||!paths.length||paths.length>128)throw Error('Selecione de 1 a 128 arquivos.');
    let inputTotal=0;
    for(const file of paths) {
        if(typeof file!=='string'||!path.isAbsolute(file)||! /\.(pdf|jpe?g|png|docx?)$/i.test(file))throw Error('Selecione PDF, JPG/JPEG, PNG, DOCX ou DOC.');
        const info=await lstat(file);
        if(!info.isFile()||info.isSymbolicLink())throw Error('Selecione somente arquivos comuns.');
        inputTotal+=info.size;
        if(inputTotal>IMPORT_TOTAL)throw Error('A seleção ultrapassa o limite de 512 MiB.');
    }
    const result=[];let total=0;
    for(const file of paths) {
        if(signal?.aborted)throw Error('Importação cancelada.');
        try {
            let entry;
            if(/\.docx?$/i.test(file)) {
                const converted=await convert(file,{...wordOptions,signal});
                entry={name:converted.name,bytes:converted.bytes,wordReport:{...converted.report,originalName:path.basename(file)}};
            }else entry={name:path.basename(file),bytes:new Uint8Array(await readImportFile(file,IMPORT_TOTAL))};
            total+=entry.bytes.length;
            if(total>IMPORT_TOTAL)throw Error('Os arquivos preparados ultrapassam 512 MiB. Divida a seleção.');
            result.push(entry);
        }catch(error){throw Error(path.basename(file)+' — '+error.message);}
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
