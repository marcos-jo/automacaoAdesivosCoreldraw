import {readFile} from 'node:fs/promises';
import {createRequire} from 'node:module';
import {createHash} from 'node:crypto';
import path from 'node:path';
const require=createRequire(import.meta.url);
// A ordem de includes e o hook são conferidos antes de depender do adaptador.
export async function checkNsisDistribution(){
    const pkg=require.resolve('app-builder-lib/package.json');
    const metadata=JSON.parse(await readFile(pkg,'utf8'));
    if(metadata.version!=='26.18.0')throw Error('Revalidar adaptador NSIS antes de atualizar electron-builder.');
    const source=await readFile(path.join(path.dirname(pkg),'templates/nsis/include/extractAppPackage.nsh'));
    const digest=createHash('sha256').update(source.toString('utf8').replaceAll('\r\n','\n')).digest('hex');
    if(digest!==EXPECTED_TEMPLATE)throw Error('Template NSIS diferente do auditado; build interrompido.');
}
const EXPECTED_TEMPLATE='e4174388a0f7a1df0b85a0742aa1ea7a4b2b18f9f29dccd6ef10a66212f68148';
