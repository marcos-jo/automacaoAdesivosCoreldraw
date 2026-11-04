import {mkdir,readFile,writeFile} from 'node:fs/promises';
import {createHash} from 'node:crypto';
import path from 'node:path';
import {createPaperFixture,createMixedPrintFixture,createPrintFormFixture,printScales} from '../tests/helpers/windows-print-fixtures.mjs';

const folder=path.resolve(process.argv[2]||'.desktop-test/fixtures/windows-print');
await mkdir(folder,{recursive:true});
const fixtures={
    'a4.pdf':await createPaperFixture(),
    'a3.pdf':await createPaperFixture({width:297,height:420}),
    'a4-rotated-90.pdf':await createPaperFixture({rotation:90}),
    'mixed-pages.pdf':await createMixedPrintFixture(),
    'form-appearance.pdf':await createPrintFormFixture()
};
const files=[];
for(const [name,bytes]of Object.entries(fixtures)){
    const destination=path.join(folder,name),sha256=createHash('sha256').update(bytes).digest('hex');
    try{await writeFile(destination,bytes,{flag:'wx'});}
    catch(error){
        if(error.code!=='EEXIST')throw error;
        if(createHash('sha256').update(await readFile(destination)).digest('hex')!==sha256)throw new Error('Arquivo diferente preservado em '+destination+'. Escolha outra pasta.');
    }
    files.push({name,sha256});
}
const manifestFile=path.join(folder,'manifest.json'),manifest=JSON.stringify({schema:1,synthetic:true,scales:printScales,files},null,2)+'\n';
try{await writeFile(manifestFile,manifest,{flag:'wx'});}
catch(error){
    if(error.code!=='EEXIST')throw error;
    if(await readFile(manifestFile,'utf8')!==manifest)throw new Error('Manifesto diferente preservado em '+manifestFile+'. Escolha outra pasta.');
}
console.log('Cinco PDFs sintéticos gerados/conferidos em '+folder+'. Nenhum trabalho de impressão foi enviado.');
