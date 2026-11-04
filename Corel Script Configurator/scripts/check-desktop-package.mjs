import {createRequire} from 'node:module';
import assert from 'node:assert/strict';
import path from 'node:path';
import {readFileSync,existsSync} from 'node:fs';
import {fileURLToPath} from 'node:url';
import {createHash} from 'node:crypto';
const require=createRequire(import.meta.url), asar=require('@electron/asar');
const file=path.resolve(process.argv[2] || 'dist-desktop/win-unpacked/resources/app.asar');
const files=asar.listPackage(file).map(f=>f.replaceAll('\\','/'));
for(const needed of ['/desktop/main.mjs','/desktop/preload.cjs','/build/icons/app.png','/public/index.html','/public/js/tools/pdfs.js','/public/vendor/pdfjs/pdf.js','/public/vendor/pdfjs/pdf.worker.js','/public/vendor/pdf-lib/pdf-lib.js']) assert.ok(files.includes(needed),'Falta '+needed);
const prohibited=files.filter(f=>/^\/(tests|docs|modelos|src)(\/|$)/.test(f)||/\.(pdf|pem|key|pfx|p12|log|tmp)$/i.test(f)||/backup/i.test(f)||/^\/node_modules\/(electron|electron-builder|pdfjs-dist|pdf-lib)(\/|$)/.test(f));
assert.deepEqual(prohibited,[],'Pacote contém arquivo fora da distribuição');
const root=fileURLToPath(new URL('../',import.meta.url));
const outdated=files.filter(f=>/^\/(desktop|public)\//.test(f)&&/\.(mjs|cjs|js|css|html|png)$/.test(f))
    .filter(f=>!asar.extractFile(file,path.normalize(f.slice(1))).equals(readFileSync(path.join(root,f.slice(1)))));
assert.deepEqual(outdated,[],'Pacote desatualizado em relação ao código atual');
assert.ok(asar.extractFile(file,path.normalize('build/icons/app.png')).equals(readFileSync(path.join(root,'build/icons/app.png'))),'Ícone da janela desatualizado');
assert.ok(readFileSync(path.join(path.dirname(file),'icons/pdf.ico')).equals(readFileSync(path.join(root,'build/icons/pdf.ico'))),'Ícone PDF instalado desatualizado');
const nativeManifest=JSON.parse(asar.extractFile(file,path.normalize('build/native/manifest.json')).toString('utf8'));
const nativeHelper=readFileSync(path.join(path.dirname(file),'native/4PrintShops.PrintConfig.exe'));
assert.equal(createHash('sha256').update(nativeHelper).digest('hex'),nativeManifest.sha256,'Auxiliar nativo difere do manifesto incorporado ao ASAR');
const labFolder=path.join(path.dirname(file),'print-lab');
if(existsSync(labFolder)){
    const metadata=JSON.parse(asar.extractFile(file,'package.json').toString('utf8'));
    assert.equal(metadata.desktopHoldTest,true,'O laboratório virtual pertence somente à identidade Hold Test');
    const manifest=JSON.parse(readFileSync(path.join(labFolder,'manifest.json'),'utf8'));
    assert.equal(manifest.allowedQueue,'Microsoft Print to PDF');
    for(const name of ['4PrintShops.PrintLab.exe','pdfium.dll','synthetic.pdf']){
        assert.equal(createHash('sha256').update(readFileSync(path.join(labFolder,name))).digest('hex'),manifest.files[name]);
    }
    assert.ok(readFileSync(path.join(labFolder,'PDFium-LICENSE.txt')).length>0);
    assert.ok(readFileSync(path.join(labFolder,'Liberation-LICENSE.txt')).length>0);
}
const holdFolder=path.join(path.dirname(file),'hold-native');
if(existsSync(holdFolder)){
    const metadata=JSON.parse(asar.extractFile(file,'package.json').toString('utf8'));
    assert.equal(metadata.desktopHoldNativeTest,true);
    assert.equal(metadata.name,'4printshops-hold-test');
    const manifest=JSON.parse(asar.extractFile(file,path.normalize('build/hold-native/manifest.json')).toString('utf8'));
    for(const name of ['4PrintShops.PrintConfig.exe','pdfium.dll','synthetic.pdf']){
        assert.equal(createHash('sha256').update(readFileSync(path.join(holdFolder,name))).digest('hex'),manifest.files[name]);
    }
    assert.equal(manifest.physicalSendRequiresConfirmation,true);
    assert.ok(!existsSync(path.join(holdFolder,'gswin64c.exe')),'Ghostscript de laboratório não pertence ao pacote');
}
console.log('PASSOU: allowlist ASAR, frontend/bibliotecas PDF presentes e testes/legado/documentos/credenciais/dependências de desenvolvimento ausentes. '+files.length+' entradas.');
