import test from 'node:test';
import assert from 'node:assert/strict';
import path from 'node:path';
import { mkdtemp, readFile, readdir, mkdir, rm, realpath } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { assetPath, isAppURL, validSender, pdfName, pdfData } from '../desktop/security.mjs';
import { savePDF, openPDFs } from '../desktop/files.mjs';

async function removeFixture(root) {
    const verified=await realpath(root),temp=await realpath(tmpdir()),relative=path.relative(temp,verified);
    assert.ok(relative && !relative.startsWith('..') && !path.isAbsolute(relative) && path.basename(verified).startsWith('4printshops-native-'));
    await rm(verified,{recursive:true,force:true});
}

test('origem fixa e allowlist não expõem caminhos, legado ou recursos externos', () => {
    const root = path.resolve('public');
    for (const url of ['app://4printshops/index.html','app://4printshops/js/tools/pdfs.js','app://4printshops/vendor/pdfjs/pdf.worker.js']) assert.ok(assetPath(root,url));
    for (const url of ['https://4printshops/index.html','app://user@4printshops/index.html','app://4printshops:99/index.html','app://other/index.html','app://4printshops/js/toolbar.js','app://4printshops/js/tools/pdfs%20backup.js','app://4printshops/%2e%2e%5cpackage.json','app://4printshops/js/%00.js','app://4printshops/api/scripts','app://4printshops/desktop/preload.cjs']) assert.equal(assetPath(root,url),null,url);
    assert.equal(isAppURL('not a url'),false);
});
test('IPC exige janela, frame principal e origem próprios', () => {
    const frame = {url:'app://4printshops/index.html'}, sender = {mainFrame:frame}, window = {webContents:sender,isDestroyed:()=>false};
    assert.equal(validSender({sender,senderFrame:frame},window),true);
    assert.equal(validSender({sender,senderFrame:{...frame}},window),false);
    assert.equal(validSender({sender:{},senderFrame:frame},window),false);
    frame.url='https://example.com'; assert.equal(validSender({sender,senderFrame:frame},window),false);
});
test('nomes e payload PDF rejeitam caminhos, nomes reservados e dados arbitrários', () => {
    assert.equal(pdfName('Cópia (A4).pdf'),'Cópia (A4).pdf');
    for(const name of ['../a.pdf','C:\\a.pdf','CON.pdf','x.exe','a.pdf ','x|y.pdf']) assert.throws(()=>pdfName(name));
    assert.throws(()=>pdfData(new Uint8Array([1,2,3]))); assert.throws(()=>pdfData('%PDF-1.7'));
});
test('salvar nativo é transacional; cancelar não escreve e falhar não trunca um destino', async () => {
    const root=await mkdtemp(path.join(tmpdir(),'4printshops-native-save-'));
    const bytes=Buffer.from('%PDF-1.7\nexemplo'), file=path.join(root,'copia.pdf');
    try {
        assert.deepEqual(await savePDF({}, {showSaveDialog:async()=>({canceled:true})},bytes,'copia.pdf'),{saved:false});
        assert.deepEqual(await readdir(root),[]);
        assert.deepEqual(await savePDF({}, {showSaveDialog:async()=>({canceled:false,filePath:file})},bytes,'copia.pdf'),{saved:true});
        assert.deepEqual(await readFile(file),bytes);
        const dir=path.join(root,'diretorio.pdf'); await mkdir(dir);
        await assert.rejects(savePDF({}, {showSaveDialog:async()=>({filePath:dir})},bytes,'copia.pdf'));
        assert.deepEqual(await readFile(file),bytes); assert.equal((await readdir(root)).some(f=>f.endsWith('.tmp')),false);
    } finally { await removeFixture(root); }
});
test('importação lê somente a seleção do diálogo, preserva bytes e não devolve caminhos', async () => {
    const root=await mkdtemp(path.join(tmpdir(),'4printshops-native-open-')),file=path.join(root,'origem.pdf'),bytes=Buffer.from('%PDF-1.7\noriginal');
    try {
        const {writeFile}=await import('node:fs/promises');await writeFile(file,bytes);
        const dialog={showOpenDialog:async()=>({canceled:false,filePaths:[file]})};
        const result=await openPDFs({},dialog,false);assert.deepEqual(result,[{name:'origem.pdf',bytes}]);assert.deepEqual(await readFile(file),bytes);
        assert.deepEqual(await openPDFs({}, {showOpenDialog:async()=>({canceled:true})}),[]);
        await assert.rejects(openPDFs({},dialog,'../origem.pdf'),/inválida/);
    } finally { await removeFixture(root); }
});
