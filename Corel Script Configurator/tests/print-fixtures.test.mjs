import test from 'node:test';
import assert from 'node:assert/strict';
import {mkdtemp,writeFile,readFile} from 'node:fs/promises';
import {execFile} from 'node:child_process';
import {promisify} from 'node:util';
import {tmpdir} from 'node:os';
import path from 'node:path';
import {PDFDocument} from '../public/vendor/pdf-lib/pdf-lib.js';
import {createPaperFixture,createMixedPrintFixture,createPrintFormFixture,printScales} from './helpers/windows-print-fixtures.mjs';
import {ghostscriptTrust} from '../scripts/build-hold-native.mjs';

test('geradores portáteis criam A4/A3, rotação e metadados determinísticos',async()=>{
    const bytes=await createPaperFixture();assert.deepEqual(bytes,await createPaperFixture());
    const a4=await PDFDocument.load(bytes),a3=await PDFDocument.load(await createPaperFixture({width:297,height:420}));
    assert.ok(Math.abs(a4.getPage(0).getWidth()*25.4/72-210)<1e-6);
    assert.ok(Math.abs(a3.getPage(0).getHeight()*25.4/72-420)<1e-6);
    assert.equal(a4.getCreationDate().toISOString(),'2026-11-04T00:00:00.000Z');
    assert.equal((await PDFDocument.load(await createPaperFixture({rotation:90}))).getPage(0).getRotation().angle,90);
});
test('fixture mista conserva ordem, giro e dimensões; formulário tem valor e aparência',async()=>{
    const bytes=await createMixedPrintFixture();assert.deepEqual(bytes,await createMixedPrintFixture());
    const pdf=await PDFDocument.load(bytes);assert.equal(pdf.getPageCount(),3);assert.equal(pdf.getPage(0).getRotation().angle,180);
    assert.ok(Math.abs(pdf.getPage(1).getWidth()*25.4/72-300)<1e-6);
    assert.ok(Math.abs(pdf.getPage(2).getHeight()*25.4/72-150)<1e-6);
    const formBytes=await createPrintFormFixture();assert.deepEqual(formBytes,await createPrintFormFixture());
    const form=(await PDFDocument.load(formBytes)).getForm();assert.equal(form.getTextField('test').getText(),'PRINTABLE FORM');
    assert.ok(form.getTextField('test').acroField.getWidgets()[0].getNormalAppearance());assert.deepEqual(printScales,[100,120,87.5]);
});
test('build sem Ghostscript conserva bloqueio do diagnóstico e aceita configuração explícita com hash',async()=>{
    assert.deepEqual(await ghostscriptTrust({configured:'none'}),{path:'',sha256:''});
    const folder=await mkdtemp(path.join(tmpdir(),'4ps-gs-test-'));
    assert.deepEqual(await ghostscriptTrust({configured:'',programFiles:folder}),{path:'',sha256:''});
    const file=path.join(folder,'gswin64c.exe');await writeFile(file,'test bytes, never executed');
    const trust=await ghostscriptTrust({configured:file});assert.equal(trust.path,file);assert.match(trust.sha256,/^[a-f0-9]{64}$/);
    await assert.rejects(ghostscriptTrust({configured:'relative.exe'}),/absoluto/);
    await assert.rejects(ghostscriptTrust({configured:path.join(folder,'missing.exe')}),/ENOENT/);
});
test('CLI preserva PDF e manifesto existentes com conteúdo diferente',async()=>{
    const run=promisify(execFile),folder=await mkdtemp(path.join(tmpdir(),'4ps-fixtures-test-'));
    const script=new URL('../scripts/generate-print-fixtures.mjs',import.meta.url);
    const {fileURLToPath}=await import('node:url');
    const args=[fileURLToPath(script),folder];
    await writeFile(path.join(folder,'a4.pdf'),'preservar este arquivo');
    await assert.rejects(run(process.execPath,args));assert.equal(await readFile(path.join(folder,'a4.pdf'),'utf8'),'preservar este arquivo');
    const next=await mkdtemp(path.join(tmpdir(),'4ps-manifest-test-'));
    await run(process.execPath,[fileURLToPath(script),next]);
    await writeFile(path.join(next,'manifest.json'),'preservar manifesto');
    await assert.rejects(run(process.execPath,[fileURLToPath(script),next]));
    assert.equal(await readFile(path.join(next,'manifest.json'),'utf8'),'preservar manifesto');
});
