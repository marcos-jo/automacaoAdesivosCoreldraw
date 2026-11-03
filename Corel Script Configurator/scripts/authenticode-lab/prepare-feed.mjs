import {readFile,writeFile,mkdir,copyFile} from 'node:fs/promises';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {createHash} from 'node:crypto';
import yaml from 'js-yaml';
import guard from './guest-guard.cjs';

guard.assertGuest(process.env.PRINTSHOPS_VM_MANIFEST,process.argv.includes('--confirm-disposable-guest'));
const root=fileURLToPath(new URL('../../',import.meta.url)),base=path.join(root,'dist-desktop/authenticode-lab');
const hash=bytes=>createHash('sha512').update(bytes).digest('base64');
for(const variant of ['valid','foreign','hash-mismatch','signature-tampered']){
    const src=path.join(base,variant==='foreign'?'foreign':'primary','1.0.0-experimental.3');
    const manifest=await readFile(path.join(src,'experimental.yml'),'utf8'),info=yaml.load(manifest);
    const name=info.files[0].url;
    if(!/^4PrintShops-Update-Lab-1\.0\.0-experimental\.3-Setup\.exe$/.test(name))throw new Error('Artefato inesperado.');
    const bytes=await readFile(path.join(src,name));if(hash(bytes)!==info.files[0].sha512)throw new Error('SHA512 diverge do artefato gerado.');
    const dest=path.join(base,'feeds',variant);await mkdir(dest,{recursive:true});
    if(variant==='hash-mismatch'||variant==='signature-tampered')bytes[Math.floor(bytes.length/2)]^=1;
    await writeFile(path.join(dest,name),bytes);
    if(variant==='signature-tampered'){info.files[0].sha512=hash(bytes);info.sha512=info.files[0].sha512;}
    await writeFile(path.join(dest,'experimental.yml'),variant==='signature-tampered'?yaml.dump(info):manifest);
    // Sem blockmap nos casos adulterados: forçar fallback completo, não editar a origem.
    if(variant==='valid'||variant==='foreign')await copyFile(path.join(src,name+'.blockmap'),path.join(dest,name+'.blockmap'));
}
console.log('Feeds valid/foreign/hash-mismatch/signature-tampered preparados somente no guest. Origens preservadas.');
