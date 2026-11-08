import test from 'node:test';
import assert from 'node:assert/strict';
import {mkdtemp,writeFile,readFile,readdir,rm,open} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import path from 'node:path';
import {PDFDocument,StandardFonts} from '../public/vendor/pdf-lib/pdf-lib.js';
import {convertWord,inspectWord,inspectPDF,isWordPath,findLibreOffice,runOffice,importWord,WORD_LIMIT} from '../desktop/word-import.mjs';
import {createWordFixture,wordZip,wordEntries} from './helpers/word-fixtures.mjs';

const temp=async fn=>{const folder=await mkdtemp(path.join(tmpdir(),'4printshops-word-test-'));try{return await fn(folder);}finally{await rm(folder,{recursive:true,force:true});}};
async function pdfBytes(){const pdf=await PDFDocument.create();const font=await pdf.embedFont(StandardFonts.Helvetica);pdf.addPage().drawText('Word seguro',{font});return Buffer.from(await pdf.save());}
test('DOCX: metadados de paginação e fontes são sugestões; DOC binário reconhecido',()=>{
    assert.deepEqual(inspectWord(createWordFixture(),'.docx'),{sourcePages:2,requestedFonts:['Arial'],legacy:false});
    const doc=Buffer.concat([Buffer.from('d0cf11e0a1b11ae1','hex'),Buffer.from('WordDocument','utf16le')]);
    assert.equal(inspectWord(doc,'.doc').legacy,true);
    assert.throws(()=>inspectWord(Buffer.from('RTF não é DOC'),'.doc'),/binário/);
});
test('Word recusa pacote falso, truncado, macros, objetos ativos e vínculos externos',()=>{
    assert.throws(()=>inspectWord(Buffer.alloc(40),'.docx'),/pacote/);
    assert.throws(()=>inspectWord(createWordFixture().subarray(0,-5),'.docx'),/incompleto/);
    for(const name of ['word/vbaProject.bin','word/activeX/x.xml','word/embeddings/object.bin'])assert.throws(()=>inspectWord(wordZip({...wordEntries(),[name]:'x'}),'.docx'),/macros/);
    assert.throws(()=>inspectWord(wordZip({...wordEntries(),'word/_rels/document.xml.rels':'<Relationships><Relationship TargetMode="External" Target="http://127.0.0.1/x"/></Relationships>'}),'.docx'),/vínculos externos/);
    assert.throws(()=>inspectWord(wordZip({...wordEntries(),'../outside':'x'}),'.docx'),/Caminho/);
    assert.throws(()=>inspectWord(wordZip({...wordEntries(),'docProps/app.xml':'<!DOCTYPE x><Pages>1</Pages>'}),'.docx'),/XML/);
});
test('PDF validado por parsing, páginas, dimensões e incorporação de fontes',async()=>{
    const report=await inspectPDF(await pdfBytes());assert.equal(report.pages,1);assert.equal(report.fonts[0].name,'Helvetica');assert.equal(report.fonts[0].embedded,false);assert.ok(report.dimensions[0].widthMm>0);
    await assert.rejects(inspectPDF(Buffer.from('nada')),/PDF válido/);
    await assert.rejects(inspectPDF(Buffer.from('%PDF-1.7 inválido')),/inválido/);
});
test('caminhos e descoberta não aceitam URL, extensão falsa ou executável relativo',async()=>{
    assert.equal(isWordPath('https://example.com/a.docx'),false);assert.equal(isWordPath('a.docx'),false);assert.equal(isWordPath(path.join(tmpdir(),'a.docx.exe')),false);
    if(process.platform==='win32')for(const file of ['C:\\a:evil.docx','\\\\server\\share\\a.docx','\\\\.\\a.doc'])assert.equal(isWordPath(file),false);
    await assert.rejects(findLibreOffice({PRINTSHOPS_LIBREOFFICE:'soffice.com'}),/não encontrado/);
});
test('conversão usa cópia/perfil isolados, argumentos sem shell e limpa tudo preservando o original',()=>temp(async folder=>{
    const file=path.join(folder,'São José & orçamento.docx'),original=createWordFixture(),pdf=await pdfBytes();await writeFile(file,original);
    const result=await convertWord(file,{tempRoot:folder,executable:'fake',runner:async(_exe,args,{cwd})=>{
        assert.ok(args[0].startsWith('-env:UserInstallation=file:'));assert.ok(args.includes('--headless'));assert.equal(args.at(-1),path.join(cwd,'documento.docx'));
        assert.deepEqual(await readFile(args.at(-1)),original);
        const profile=await readFile(path.join(cwd,'profile/user/registrymodifications.xcu'),'utf8');assert.match(profile,/DisableMacrosExecution/);assert.match(profile,/DisableActiveContent/);assert.match(profile,/BlockUntrustedRefererLinks/);
        await writeFile(path.join(cwd,'out/documento.pdf'),pdf);
    }});
    assert.equal(result.name,'São José & orçamento_WORD.pdf');assert.equal(result.report.pages,1);assert.equal(result.report.sourcePages,2);
    assert.deepEqual(await readFile(file),original);assert.deepEqual(await readdir(folder),[path.basename(file)]);
}));
test('erro, saída ausente e cancelamento limpam temporários e permitem repetir',()=>temp(async folder=>{
    const file=path.join(folder,'teste.docx');await writeFile(file,createWordFixture());
    for(const runner of [async()=>{throw Error('falha do processo');},async()=>{}])await assert.rejects(convertWord(file,{tempRoot:folder,executable:'fake',runner}));
    const controller=new AbortController();
    await assert.rejects(convertWord(file,{tempRoot:folder,executable:'fake',signal:controller.signal,runner:async()=>{controller.abort();}}),/cancelada/);
    assert.deepEqual(await readdir(folder),['teste.docx']);
    let called=false;await assert.rejects(convertWord(file,{tempRoot:folder,signal:controller.signal,runner:async()=>{called=true;}}),/cancelada/);assert.equal(called,false);
}));
test('arquivo alterado durante conversão, grande, renomeado e inexistente são recusados',()=>temp(async folder=>{
    const file=path.join(folder,'teste.docx');await writeFile(file,createWordFixture());
    await assert.rejects(convertWord(file,{tempRoot:folder,executable:'fake',runner:async(_exe,_args,{cwd})=>{await writeFile(path.join(cwd,'out/documento.pdf'),await pdfBytes());await writeFile(file,createWordFixture({font:'Outra'}));}}),/mudou/);
    const handle=await open(file,'w');await handle.truncate(WORD_LIMIT+1);await handle.close();await assert.rejects(convertWord(file),/limite/);
    await assert.rejects(convertWord(path.join(folder,'ausente.docx')),/não foi possível/i);
    assert.deepEqual(await readdir(folder),['teste.docx']);
}));
test('cancelar diálogo não executa conversor; seleção múltipla é recusada',async()=>{
    assert.equal(await importWord(null,{showOpenDialog:async()=>({canceled:true})}),null);
    await assert.rejects(importWord(null,{showOpenDialog:async()=>({filePaths:['a.docx','b.docx']})}),/somente um/);
});
test('processo real sem shell tem deadline e cancelamento, sem matar processos alheios',async()=>{
    await runOffice(process.execPath,['-e','process.exit(0)']);
    await assert.rejects(runOffice(process.execPath,['-e','setInterval(()=>{},1000)'],{timeoutMs:100}),/excedeu/);
    const controller=new AbortController(),running=runOffice(process.execPath,['-e','setInterval(()=>{},1000)'],{signal:controller.signal});setTimeout(()=>controller.abort(),100);
    await assert.rejects(running,/cancelada/);
});
