import {readdir,readFile,mkdir,copyFile,writeFile} from 'node:fs/promises';
import {createHash} from 'node:crypto';
import {fileURLToPath} from 'node:url';
import path from 'node:path';
import guard from './guest-guard.cjs';

const root=fileURLToPath(new URL('../../',import.meta.url));
const output=path.join(root,'dist-desktop','authenticode-vm-kit');
const machine=guard.inspectMachine();
await mkdir(path.dirname(output),{recursive:true});
await mkdir(output); // Recusar kit prévio; nunca misturar sobras ou sobrescrever.
const allowed=[/^desktop\/[^/]+\.(mjs|cjs)$/, /^public\/index\.html$/, /^public\/css\/[^/]+\.css$/, /^public\/img\/logo\.png$/,
    /^public\/js\/main\.js$/, /^public\/js\/app\/(shell|tabs|menus|export-dialog|save-file|desktop-client|desktop-updates)\.js$/,
    /^public\/js\/tools\/[^/]+\.js$/, /^public\/vendor\/(pdfjs|pdf-lib)\/.+\.js$/];
const files=['package.json','package-lock.json','electron-builder.config.cjs','electron-builder.update-lab.cjs','electron-builder.authenticode-lab.cjs',
    'scripts/serve-update-lab.mjs','scripts/authenticode-lab/guest-guard.cjs','scripts/authenticode-lab/create-identity.mjs',
    'scripts/authenticode-lab/prepare-feed.mjs','docs/LABORATORIO_AUTHENTICODE.md'];
async function visit(folder){for(const entry of await readdir(path.join(root,folder),{withFileTypes:true})){const name=folder+'/'+entry.name;
    if(entry.isDirectory())await visit(name);else if(entry.isFile()&&allowed.some(rule=>rule.test(name))&&!/backup/i.test(name))files.push(name);}}
await visit('desktop');await visit('public');
const hashes=[];
for(const file of files){const dest=path.join(output,'source',file);await mkdir(path.dirname(dest),{recursive:true});await copyFile(path.join(root,file),dest);
    hashes.push({file,sha256:createHash('sha256').update(await readFile(dest)).digest('hex')});}
const versions={};for(const name of ['electron-builder','electron-updater','electron'])versions[name]=JSON.parse(await readFile(path.join(root,'node_modules',name,'package.json'),'utf8')).version;
const manifest={schema:'4printshops-authenticode-kit-v1',hostFingerprint:machine.fingerprint,host:{model:machine.model,manufacturer:machine.manufacturer},
    created:new Date().toISOString(),versions,files:hashes};
await writeFile(path.join(output,'lab-kit.json'),JSON.stringify(manifest,null,2)+'\n');
console.log('Kit de fontes selecionadas: '+output+' ('+files.length+' arquivos). Sem dependências, certificados, Git ou documentos de clientes. Provisionar Node/dependências somente na VM.');
