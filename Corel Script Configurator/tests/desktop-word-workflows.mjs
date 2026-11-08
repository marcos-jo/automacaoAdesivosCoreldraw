// Compatibilidade opt-in dos PDFs sintéticos convertidos; não imprime.
import assert from 'node:assert/strict';
import {readFile} from 'node:fs/promises';
import {PDFDocument} from '../public/vendor/pdf-lib/pdf-lib.js';
import {gerarMontagem} from './helpers/pdf-imposition-modules.mjs';
const files=process.argv.slice(2);if(!files.length)throw Error('Informe os PDFs sintéticos convertidos pelo ensaio Word.');
for(const file of files) {
    const bytes=new Uint8Array(await readFile(file)),original=bytes.slice();
    for(const config of [
        {modo:'repeat',largura:500,altura:500,automatico:true,quantidade:2,marcas:false,paginas:[1]},
        {modo:'nup',largura:500,altura:500,colunas:2,linhas:1,escalaModo:'reduzir',marcas:false,paginas:[1,2]},
        {modo:'livreto',escalaModo:'ajustar'}
    ]) {
        const result=await gerarMontagem(bytes,config).catch(error=>{throw Error(config.modo+': '+error.message);});
        assert.ok((await PDFDocument.load(result.bytes)).getPageCount()>0);
        assert.deepEqual(bytes,original);
        console.log(config.modo+' aprovado: '+file);
    }
}
