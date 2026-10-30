import {createRequire} from 'node:module';
import assert from 'node:assert/strict';
import path from 'node:path';
import {readFileSync} from 'node:fs';
import {fileURLToPath} from 'node:url';
const require=createRequire(import.meta.url), asar=require('@electron/asar');
const file=path.resolve(process.argv[2] || 'dist-desktop/win-unpacked/resources/app.asar');
const files=asar.listPackage(file).map(f=>f.replaceAll('\\','/'));
for(const needed of ['/desktop/main.mjs','/desktop/preload.cjs','/public/index.html','/public/js/tools/pdfs.js','/public/vendor/pdfjs/pdf.js','/public/vendor/pdfjs/pdf.worker.js','/public/vendor/pdf-lib/pdf-lib.js']) assert.ok(files.includes(needed),'Falta '+needed);
const prohibited=files.filter(f=>/^\/(tests|docs|modelos|src)(\/|$)/.test(f)||/\.(pdf|pem|key|pfx|p12|log|tmp)$/i.test(f)||/backup/i.test(f)||/^\/node_modules\/(electron|electron-builder|pdfjs-dist|pdf-lib)(\/|$)/.test(f));
assert.deepEqual(prohibited,[],'Pacote contém arquivo fora da distribuição');
const root=fileURLToPath(new URL('../',import.meta.url));
const outdated=files.filter(f=>/^\/(desktop|public)\//.test(f)&&/\.(mjs|cjs|js|css|html|png)$/.test(f))
    .filter(f=>!asar.extractFile(file,path.normalize(f.slice(1))).equals(readFileSync(path.join(root,f.slice(1)))));
assert.deepEqual(outdated,[],'Pacote desatualizado em relação ao código atual');
console.log('PASSOU: allowlist ASAR, frontend/bibliotecas PDF presentes e testes/legado/documentos/credenciais/dependências de desenvolvimento ausentes. '+files.length+' entradas.');
