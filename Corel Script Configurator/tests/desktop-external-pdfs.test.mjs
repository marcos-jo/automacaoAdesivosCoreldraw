import test from 'node:test';
import assert from 'node:assert/strict';
import {mkdtemp,writeFile,unlink,open,mkdir} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import path from 'node:path';
import {pdfArguments,isPDFPath,readExternalPDFs,ExternalPDFQueue,IMPORT_LIMIT} from '../desktop/external-pdfs.mjs';
import {PDFDocument} from '../public/vendor/pdf-lib/pdf-lib.js';

test('argv inicial/segunda instância aceita apenas PDFs posicionais ou após marcador, sem parâmetros arbitrários',()=>{
    const a='C:\\Gráfica São José\\proposta final.pdf',b='D:\\outro.PDF';
    assert.deepEqual(pdfArguments(['app.exe','--open-pdf',a,b,a]),[a,b]);
    assert.deepEqual(pdfArguments(['electron.exe','main.mjs','--remote-debugging-port=0',a],{packaged:false}),[a]);
    assert.deepEqual(pdfArguments(['app.exe','--arbitrary-option',a,'--foo='+b,'https://example.com/a.pdf','file:///C:/a.pdf']),[]);
});
test('caminhos UNC/longos são reconhecidos; relativos, ADS e dispositivos são recusados',()=>{
    for(const p of ['\\\\servidor\\share\\São José.pdf','\\\\?\\C:\\gráfica\\a.pdf','\\\\?\\UNC\\servidor\\share\\a.pdf'])assert.equal(isPDFPath(p),true,p);
    for(const p of ['a.pdf','C:a.pdf','\\a.pdf','\\\\.\\pipe\\a.pdf','\\\\?\\GLOBALROOT\\a.pdf','C:\\a:stream.pdf','C:\\a.pdf.exe','C:\\bad\u0000.pdf'])assert.equal(isPDFPath(p),false,p);
});
test('fila antes de a interface estar pronta conserva todos os arquivos e deduplica caminhos pendentes',async()=>{
    let reads=0,notifications=0;
    const q=new ExternalPDFQueue(()=>notifications++,async paths=>{reads++;return paths;});
    q.add(['C:\\um.pdf','C:\\dois.pdf']);q.add(['c:\\UM.pdf','C:\\três.pdf']);
    assert.equal(q.list().length,3);assert.equal(reads,0);assert.equal(notifications,2);
    const ids=q.list().map(f=>f.id);assert.deepEqual(await q.read(ids),['C:\\um.pdf','C:\\dois.pdf','C:\\três.pdf']);
    assert.equal(q.active,true);await assert.rejects(q.read(ids),/Aguarde/);q.finish(ids);assert.equal(q.list().length,0);
    q.add(['C:\\um.pdf']);assert.equal(q.list().length,1,'Reabrir depois da conclusão é intencionalmente permitido');
});
test('IDs opacos não aceitam caminhos do renderer; falha libera leitura e permite repetir sem perder a fila',async()=>{
    let fail=true;const q=new ExternalPDFQueue(()=>{},async()=>{if(fail)throw new Error('sem acesso');return [];});q.add(['C:\\a.pdf']);
    const ids=q.list().map(f=>f.id);await assert.rejects(q.read(['C:\\a.pdf']),/inválida/);
    await assert.rejects(q.read(ids),/sem acesso/);assert.equal(q.active,false);assert.equal(q.list().length,1);
    fail=false;await q.read(ids);q.release(ids);assert.equal(q.active,false);await q.read(ids);q.finish(ids);
});
test('limite da fila é atômico e não descarta os pedidos anteriores',()=>{
    const q=new ExternalPDFQueue();q.add(['C:\\a.pdf']);assert.throws(()=>q.add(Array.from({length:128},(_,i)=>`C:\\b${i}.pdf`)),/128/);assert.equal(q.list().length,1);
});
test('leitura real preserva acentos/espaços e bytes; inexistente, texto e arquivo acima do limite são recusados',async()=>{
    const folder=await mkdtemp(path.join(tmpdir(),'4printshops-open-node-'));
    const good=path.join(folder,'Proposta São José com espaços.pdf'),bad=path.join(folder,'inválido.pdf'),big=path.join(folder,'grande.pdf');
    try{
        const pdf=await PDFDocument.create();pdf.addPage();const bytes=Buffer.from(await pdf.save());await writeFile(good,bytes);await writeFile(bad,'isto não é PDF');
        assert.deepEqual((await readExternalPDFs([good]))[0],{name:path.basename(good),bytes});
        await assert.rejects(readExternalPDFs([path.join(folder,'ausente.pdf')]),/não encontrado/);
        await assert.rejects(readExternalPDFs([bad]),/cabeçalho/);
        const handle=await open(big,'w');await handle.truncate(IMPORT_LIMIT+1);await handle.close();
        await assert.rejects(readExternalPDFs([big]),/512 MiB/);
    }finally{for(const file of [good,bad,big])await unlink(file).catch(()=>{});}
});

test('arquivo real em caminho acima de 260 caracteres permanece legível sem truncar nome',async()=>{
    const folder=await mkdtemp(path.join(tmpdir(),'4printshops-open-long-'));
    const nested=path.join(folder,'p'.repeat(100),'q'.repeat(100));await mkdir(nested,{recursive:true});
    const file=path.join(nested,'Proposta São José '+ 'a'.repeat(90)+'.pdf');
    try{const pdf=await PDFDocument.create();pdf.addPage();const bytes=Buffer.from(await pdf.save());await writeFile(file,bytes);assert.ok(file.length>260);assert.deepEqual((await readExternalPDFs([file]))[0].bytes,bytes);}
    finally{await unlink(file).catch(()=>{});}
});
