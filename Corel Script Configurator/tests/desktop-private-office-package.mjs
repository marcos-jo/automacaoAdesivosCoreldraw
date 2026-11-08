// Actual packaged modules + private runtime. No windows, system Office or printing.
import assert from 'node:assert/strict';
import * as electron from 'electron';
import {readFile,writeFile,mkdtemp,mkdir,rm} from 'node:fs/promises';
import path from 'node:path';
import {tmpdir} from 'node:os';
import {pathToFileURL} from 'node:url';
import {createWordFixture} from './helpers/word-fixtures.mjs';
const packageRoot=path.resolve(process.argv[2]||'dist-desktop/hold-native-test/win-unpacked');
const resources=path.join(packageRoot,'resources');
const folder=await mkdtemp(path.join(tmpdir(),'4printshops-private-package-'));
try {
    const api=await import(pathToFileURL(path.join(resources,'app.asar/desktop/word-import.mjs')).href);
    const {createOfficeResolver,verifyOfficeRuntime}=await import(pathToFileURL(path.join(resources,'app.asar/desktop/libreoffice-runtime.mjs')).href);
    const blocked=path.join(folder,'no-system-office');
    process.env.ProgramFiles=blocked;process.env['ProgramFiles(x86)']=blocked;process.env.PRINTSHOPS_LIBREOFFICE=path.join(blocked,'soffice.com');process.env.PATH=blocked;
    const resolver=createOfficeResolver({packaged:true,resourcesPath:resources});
    const engine=await resolver();assert.equal(engine.mode,'bundled');assert.equal(engine.executable,path.join(resources,'libreoffice/program/soffice.com'));
    const source=path.join(folder,'São José & Word privado.docx'),original=createWordFixture();await writeFile(source,original);
    const a=await api.convertWord(source,{resolveExecutable:resolver});assert.equal(a.report.pages,2);await writeFile(path.join(folder,'docx.pdf'),a.bytes);assert.deepEqual(await readFile(source),original);
    const profile=path.join(folder,'generator');await mkdir(path.join(profile,'user'),{recursive:true});await writeFile(path.join(profile,'user/registrymodifications.xcu'),api.safeProfile);
    await api.runOffice(engine.executable,['-env:UserInstallation='+pathToFileURL(profile).href,'--headless','--convert-to','doc:MS Word 97','--outdir',folder,source],{env:engine.env,cwd:folder});
    const legacy=source.replace(/docx$/,'doc'),legacyBytes=await readFile(legacy);
    const b=await api.convertWord(legacy,{resolveExecutable:resolver});assert.equal(b.report.pages,2);await writeFile(path.join(folder,'doc.pdf'),b.bytes);assert.deepEqual(await readFile(legacy),legacyBytes);
    const timeout=new AbortController();timeout.abort();await assert.rejects(api.convertWord(source,{resolveExecutable:resolver,signal:timeout.signal}),/cancelada/);
    await assert.rejects(createOfficeResolver({packaged:true,resourcesPath:path.join(folder,'missing')})(),/conversor interno não foi encontrado/);
    // Copy neither global preferences nor licensing changes into the runtime.
    const integrity=await verifyOfficeRuntime(path.join(resources,'libreoffice'),path.join(resources,'app.asar/build/libreoffice/trust.json'));
    const result={passed:true,packageRoot,privateExecutable:engine.executable,version:engine.version,verifiedFiles:integrity.files,docxPages:a.report.pages,docPages:b.report.pages,systemOfficeLookupDisabled:true,originalsPreserved:true,requiresCleanWindowsValidation:true};
    await writeFile(path.join(folder,'result.json'),JSON.stringify(result,null,2));console.log(JSON.stringify({...result,folder}));
}catch(error){console.error(error.stack);process.exitCode=1;}
finally{electron.app.exit(process.exitCode||0);}
