import test from 'node:test';
import assert from 'node:assert/strict';
import {mkdtemp,mkdir,writeFile,readFile,rm,unlink} from 'node:fs/promises';
import path from 'node:path';
import {tmpdir} from 'node:os';
import {createHash} from 'node:crypto';
import {verifyOfficeRuntime,createOfficeResolver,privateOfficeEnvironment} from '../desktop/libreoffice-runtime.mjs';
import {verifyPreparedLibreOffice} from '../scripts/prepare-libreoffice.mjs';
import {convertWord} from '../desktop/word-import.mjs';
import {readImportedFiles} from '../desktop/import-files.mjs';
import {createWordFixture} from './helpers/word-fixtures.mjs';
import {PDFDocument} from '../public/vendor/pdf-lib/pdf-lib.js';
const digest=bytes=>createHash('sha256').update(bytes).digest('hex');
const required=['program/soffice.com','program/soffice.bin','program/mergedlo.dll','program/vcruntime140.dll','program/vcruntime140_1.dll','program/msvcp140.dll','LICENSE.html','license.txt','NOTICE','4PRINTSHOPS-NOTICE.txt','SOURCE-CODE.json','share/fonts/truetype/test.ttf'];
const temp=async fn=>{const folder=await mkdtemp(path.join(tmpdir(),'4printshops-private-office-'));try{return await fn(folder);}finally{await rm(folder,{recursive:true,force:true});}};
async function fixture(project,{omit}={}) {
    const root=path.join(project,'build/libreoffice/runtime'),trustFile=path.join(project,'build/libreoffice/trust.json'),files=[];
    await mkdir(root,{recursive:true});
    for(const relative of required.filter(p=>p!==omit)) {
        const bytes=Buffer.from('synthetic '+relative);await mkdir(path.dirname(path.join(root,relative)),{recursive:true});await writeFile(path.join(root,relative),bytes);
        files.push({path:relative,size:bytes.length,sha256:digest(bytes)});
    }
    // Same traversal ordering as the preparer, without needing any Office binary.
    files.sort((a,b)=>a.path<b.path?-1:1);
    const manifest={schema:1,version:'26.8.2.2',msiSha256:'fixture',files};
    const bytes=Buffer.from(JSON.stringify(manifest));await writeFile(path.join(root,'runtime-manifest.json'),bytes);
    await writeFile(trustFile,JSON.stringify({schema:1,version:manifest.version,manifestSha256:digest(bytes)}));
    return {root,trustFile,manifest,bytes};
}
test('motor privado verifica manifesto, todos os arquivos e componentes obrigatórios',()=>temp(async project=>{
    const f=await fixture(project),result=await verifyOfficeRuntime(f.root,f.trustFile);
    assert.equal(result.files,required.length);assert.equal(result.executable,path.join(f.root,'program/soffice.com'));
}));
test('manifesto alterado e arquivo ausente/adulterado são recusados antes da execução',()=>temp(async project=>{
    const f=await fixture(project);await writeFile(path.join(f.root,'runtime-manifest.json'),'{}');await assert.rejects(verifyOfficeRuntime(f.root,f.trustFile),/manifesto/);
    await writeFile(path.join(f.root,'runtime-manifest.json'),f.bytes);await writeFile(path.join(f.root,'program/mergedlo.dll'),'alterado');await assert.rejects(verifyOfficeRuntime(f.root,f.trustFile),/incompleto|alterado/);
    await unlink(path.join(f.root,'program/soffice.com'));await assert.rejects(verifyOfficeRuntime(f.root,f.trustFile));
}));
test('arquivos extras e componente obrigatório ausente são recusados',()=>temp(async project=>{
    const f=await fixture(project);await writeFile(path.join(f.root,'program/unlisted.dll'),'extra');await assert.rejects(verifyOfficeRuntime(f.root,f.trustFile),/não autorizado/);
    await unlink(path.join(f.root,'program/unlisted.dll'));const other=path.join(project,'other');const g=await fixture(other,{omit:'LICENSE.html'});await assert.rejects(verifyOfficeRuntime(g.root,g.trustFile),/obrigatório/);
}));
test('caminho que escapa do motor é recusado mesmo em manifesto de teste assinado pelo seu próprio hash',()=>temp(async project=>{
    const f=await fixture(project);f.manifest.files[0].path='../outside.dll';const bytes=Buffer.from(JSON.stringify(f.manifest));await writeFile(path.join(f.root,'runtime-manifest.json'),bytes);await writeFile(f.trustFile,JSON.stringify({version:'26.8.2.2',manifestSha256:digest(bytes)}));await assert.rejects(verifyOfficeRuntime(f.root,f.trustFile),/Caminho inválido/);
}));
test('aplicativo empacotado sem motor não recorre ao Office do Windows ou variável de ambiente',()=>temp(async project=>{
    const system=path.join(project,'soffice.com');await writeFile(system,'fake');
    const resolve=createOfficeResolver({packaged:true,resourcesPath:project,environment:{PRINTSHOPS_LIBREOFFICE:system}});
    await assert.rejects(resolve(),/conversor interno não foi encontrado/);
}));
test('desenvolvimento usa cópia privada antes de qualquer instalação e falha de integridade não faz fallback',()=>temp(async project=>{
    const f=await fixture(project),environment={PRINTSHOPS_LIBREOFFICE:path.join(project,'global-soffice.com'),SystemRoot:'C:\\Windows',PATH:'unsafe'};
    const resolve=createOfficeResolver({projectRoot:project,environment});const result=await resolve();assert.equal(result.mode,'bundled');assert.equal(result.executable,path.join(f.root,'program/soffice.com'));assert.equal(result.env.PATH,path.join(environment.SystemRoot,'System32'));
    await writeFile(path.join(f.root,'program/mergedlo.dll'),'broken');await assert.rejects(createOfficeResolver({projectRoot:project,environment})(),/alterado|incompleto/);
}));
test('desenvolvimento sem bundle preserva seleção explícita do motor externo',()=>temp(async project=>{
    const file=path.join(project,'soffice.com');await writeFile(file,'synthetic');const result=await createOfficeResolver({projectRoot:project,environment:{PRINTSHOPS_LIBREOFFICE:file}})();assert.equal(result.mode,'system-development');assert.equal(result.executable,file);
}));
test('ambiente do motor privado usa fontes próprias e remove caminhos de outros Offices/Python',()=>{
    const source={SystemRoot:'C:\\Windows',PATH:'global',URE_BOOTSTRAP:'global',Uno_Path:'global',PYTHONHOME:'global',SAL_FONTPATH_PRIVATE:'global',TEMP:'temp'};
    const env=privateOfficeEnvironment('C:\\private',source);assert.equal(env.URE_BOOTSTRAP,undefined);assert.equal(env.Uno_Path,undefined);assert.equal(env.PYTHONHOME,undefined);assert.equal(env.PYTHONDONTWRITEBYTECODE,'1');assert.equal(env.SAL_FONTPATH_PRIVATE,path.join('C:\\private','share/fonts/truetype'));assert.equal(source.PATH,'global');assert.equal(env.TEMP,'temp');
});
test('conversor recebe executável/ambiente privado e lote encaminha a resolução sem expor caminhos no renderer',()=>temp(async project=>{
    const file=path.join(project,'a.docx');await writeFile(file,createWordFixture());const pdf=await PDFDocument.create();pdf.addPage();const bytes=await pdf.save();const env={PRIVATE_ENGINE_TEST:'1'};let resolverCalls=0;
    const options={resolveExecutable:async()=>{resolverCalls++;return {executable:'private-soffice.com',env};},runner:async(exe,_args,request)=>{assert.equal(exe,'private-soffice.com');assert.equal(request.env.PRIVATE_ENGINE_TEST,'1');assert.equal(request.env.PYTHONPYCACHEPREFIX,path.join(request.cwd,'python-cache'));assert.deepEqual(env,{PRIVATE_ENGINE_TEST:'1'});await writeFile(path.join(request.cwd,'out/documento.pdf'),bytes);}};
    const result=await convertWord(file,options);assert.equal(result.report.pages,1);assert.equal(resolverCalls,1);
    let received;await readImportedFiles([file],{wordOptions:options,convert:async(_file,value)=>{received=value;return result;}});assert.equal(received.resolveExecutable,options.resolveExecutable);
}));
test('builder recusa preparação ausente; pins oficiais e fontes correspondentes são explícitos',()=>temp(async project=>{
    await assert.rejects(verifyPreparedLibreOffice(project));
    const lock=JSON.parse(await readFile(new URL('../build/libreoffice-lock.json',import.meta.url),'utf8'));assert.equal(lock.arch,'x64');assert.match(lock.msi.sha256,/^[a-f0-9]{64}$/);assert.equal(lock.sourceArchives.length,4);for(const source of lock.sourceArchives)assert.match(source.sha256,/^[a-f0-9]{64}$/);
}));
