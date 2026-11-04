// Integração real até a preparação. Este arquivo nunca solicita send-hold.
import assert from 'node:assert/strict';
import {spawn,execFile} from 'node:child_process';
import {promisify} from 'node:util';
import {readFile,writeFile} from 'node:fs/promises';
import path from 'node:path';
import {setTimeout as pause} from 'node:timers/promises';
import {PrintConfigClient} from '../desktop/print-config.mjs';
import {PDFDocument,degrees} from '../public/vendor/pdf-lib/pdf-lib.js';

const exec=promisify(execFile);
const client=new PrintConfigClient({helper:path.resolve('build/hold-native/4PrintShops.PrintConfig.exe'),manifest:path.resolve('build/hold-native/manifest.json'),experimental:true,
    launch:(file,args,options)=>spawn(file,args,{...options,windowsHide:false})});
const ps=path.join(process.env.SystemRoot,'System32/WindowsPowerShell/v1.0/powershell.exe');
const script=await readFile(new URL('./helpers/windows-print-properties.ps1',import.meta.url),'utf8');
const native=action=>exec(ps,['-NoProfile','-NonInteractive','-Command',`& {${script}} -DriverProcess ${client.child.pid} -Action '${action}'`],{windowsHide:true,timeout:30000});
try {
    let state=await client.request({action:'select',printer:'KONICA MINOLTA C6100PS'});
    const pending=client.request({action:'properties',session:state.session,revision:state.revision});
    await native('a3-sizes');await native('workflow');await pause(250);await native('hold');await native('accept');state=await pending;
    const original=await readFile('build/hold-native/synthetic.pdf');
    const prepare=bytes=>client.request({action:'prepare-hold',session:state.session,revision:state.revision,pdf:Buffer.from(bytes).toString('base64')});
    const first=await prepare(original);assert.equal(first.outputValidated,true);assert.equal(first.scale,120);assert.equal(first.copies,1);
    const document=await PDFDocument.load(original);document.getPage(0).setRotation(degrees(180));
    const rotated=await prepare(await document.save());assert.equal(rotated.outputValidated,true);assert.notEqual(first.token,rotated.token);
    const invalid=await PDFDocument.create();invalid.addPage([595.2756,841.8898]).drawText('Documento diferente');
    await assert.rejects(prepare(await invalid.save()),/sintético|Conteúdo|vetorial/);
    await writeFile('.desktop-test/hold-native/prepare-result.json',JSON.stringify({passed:true,originalValidated:true,editedRotation180Validated:true,differentDocumentRejected:true,output:'Hold A3 120%',physicalSendRequested:false},null,2));
    console.log('PASSOU: preparação real, PS interpretado, giro de 180° do documento serializado, sintético exclusivo e nenhuma chamada de envio.');
} finally {client.destroy();}
