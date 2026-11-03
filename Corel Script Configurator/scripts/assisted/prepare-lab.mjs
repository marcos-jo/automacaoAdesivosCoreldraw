import {mkdir,readFile,writeFile} from 'node:fs/promises';
import path from 'node:path';
import {randomUUID} from 'node:crypto';
import {generatePublishingKey} from './keys.mjs';

const folder=path.resolve('dist-desktop/assisted-lab');await mkdir(folder,{recursive:true});
const pointer=path.join(folder,'lab-local.json');
try{await readFile(pointer);throw new Error('Laboratório já preparado. Reutilize lab-local.json; não gere outra identidade sobre os pacotes existentes.');}
catch(error){if(error.code!=='ENOENT')throw error;}
if(!process.env.LOCALAPPDATA)throw new Error('LOCALAPPDATA ausente.');
const keys=await generatePublishingKey(path.join(process.env.LOCALAPPDATA,'4PrintShops-Publishing-Lab-'+randomUUID()),{laboratory:true});
await writeFile(pointer,JSON.stringify({...keys,key:path.join(keys.folder,'publication.dpapi'),repository:path.join(folder,'repository')},null,2)+'\n',{flag:'wx'});
console.log('Laboratório preparado. Referências locais (sem segredos): '+pointer);
