// Actual packaged modules + private runtime. No windows, system Office or printing.
import assert from 'node:assert/strict';
import * as electron from 'electron';
import {readFile,writeFile,mkdtemp,mkdir,rm} from 'node:fs/promises';
import path from 'node:path';
import {tmpdir} from 'node:os';
import {pathToFileURL} from 'node:url';
import {createWordFixture} from './helpers/word-fixtures.mjs';
import * as pdfjs from '../public/vendor/pdfjs/pdf.js';
pdfjs.GlobalWorkerOptions.workerSrc=new URL('../public/vendor/pdfjs/pdf.worker.js',import.meta.url).href;
Math.sumPrecise??=values=>Array.from(values).reduce((a,b)=>a+b,0);
Uint8Array.prototype.toHex??=function(){return Buffer.from(this).toString('hex');};
async function checkContent(bytes,report){
    for(const d of report.dimensions){assert.ok(Math.abs(d.widthMm-210)<.1);assert.ok(Math.abs(d.heightMm-297)<.1);}
    assert.ok(report.fonts.length>0,'PDF convertido precisa conservar fontes/texto');
    const task=pdfjs.getDocument({data:new Uint8Array(bytes)});
    try{const doc=await task.promise;let text='';for(let i=1;i<=doc.numPages;i++)text+=(await(await doc.getPage(i)).getTextContent()).items.map(item=>item.str).join(' ')+' ';text=text.normalize('NFC').replace(/\s+/g,' ');for(const phrase of ['4PrintShops','Gráfica São José','Tabela: acentos çãé','Quantidade 100','Fim do documento'])assert.ok(text.includes(phrase),phrase);return{characters:text.length,dimensions:report.dimensions,textPreserved:true};}
    finally{await task.destroy();}
}
const packageRoot=path.resolve(process.argv[2]||'dist-desktop/hold-native-test/win-unpacked');
const resources=path.join(packageRoot,'resources');
const folder=await mkdtemp(path.join(tmpdir(),'4printshops-private-package-'));
try {
    const api=await import(pathToFileURL(path.join(resources,'app.asar/desktop/word-import.mjs')).href);
    const {createOfficeResolver,verifyOfficeRuntime}=await import(pathToFileURL(path.join(resources,'app.asar/desktop/libreoffice-runtime.mjs')).href);
    const blocked=path.join(folder,'no-system-office');
    process.env.ProgramFiles=blocked;process.env['ProgramFiles(x86)']=blocked;process.env.PRINTSHOPS_LIBREOFFICE=path.join(blocked,'soffice.com');process.env.PATH=blocked;
    const resolver=createOfficeResolver({packaged:true,resourcesPath:resources});
    const engine=await resolver();assert.equal(engine.mode,'bundled');assert.equal(engine.executable,path.join(packageRoot,'lo/program/soffice.com'));
    const source=path.join(folder,'São José & Word privado.docx'),original=createWordFixture();await writeFile(source,original);
    const a=await api.convertWord(source,{resolveExecutable:resolver});assert.equal(a.report.pages,2);await writeFile(path.join(folder,'docx.pdf'),a.bytes);assert.deepEqual(await readFile(source),original);
    const profile=path.join(folder,'generator');await mkdir(path.join(profile,'user'),{recursive:true});await writeFile(path.join(profile,'user/registrymodifications.xcu'),api.safeProfile);
    await api.runOffice(engine.executable,['-env:UserInstallation='+pathToFileURL(profile).href,'--headless','--convert-to','doc:MS Word 97','--outdir',folder,source],{env:engine.env,cwd:folder});
    const legacy=source.replace(/docx$/,'doc'),legacyBytes=await readFile(legacy);
    const b=await api.convertWord(legacy,{resolveExecutable:resolver});assert.equal(b.report.pages,2);await writeFile(path.join(folder,'doc.pdf'),b.bytes);assert.deepEqual(await readFile(legacy),legacyBytes);
    const content={docx:await checkContent(a.bytes,a.report),doc:await checkContent(b.bytes,b.report)};
    const timeout=new AbortController();timeout.abort();await assert.rejects(api.convertWord(source,{resolveExecutable:resolver,signal:timeout.signal}),/cancelada/);
    await assert.rejects(createOfficeResolver({packaged:true,resourcesPath:path.join(folder,'missing')})(),/conversor interno não foi encontrado/);
    // Copy neither global preferences nor licensing changes into the runtime.
    const integrity=await verifyOfficeRuntime(path.join(packageRoot,'lo'),path.join(resources,'app.asar/build/libreoffice/trust.json'));
    const result={passed:true,packageRoot,privateExecutable:engine.executable,version:engine.version,verifiedFiles:integrity.files,docxPages:a.report.pages,docPages:b.report.pages,content,systemOfficeLookupDisabled:true,originalsPreserved:true,requiresCleanWindowsValidation:true};
    await writeFile(path.join(folder,'result.json'),JSON.stringify(result,null,2));console.log(JSON.stringify({...result,folder}));
}catch(error){console.error(error.stack);process.exitCode=1;}
finally{electron.app.exit(process.exitCode||0);}
