import test from 'node:test';
import assert from 'node:assert/strict';
import {PDFDocument, PDFName} from '../public/vendor/pdf-lib/pdf-lib.js';
import {PDFCombinacao, analisarImagem, importarImagem, matrizImagem} from './helpers/pdf-combine-modules.mjs';
import {PDFWorkspace} from './helpers/pdf-resize-modules.mjs';
import {fixtures, jpg, png} from './helpers/pdf-combine-fixtures.mjs';
import {criarPDFAssinado} from './helpers/pdf-signature-fixture.mjs';

test('combina dois PDFs de várias páginas, JPEG e PNG na ordem escolhida sem rasterizar PDFs',async()=>{
    const session=new PDFCombinacao(), files=await fixtures(); await session.add(files);
    assert.equal(session.pages.length,7);assert.equal(session.sources.length,5);
    session.move(6,0);session.move(4,2);
    const order=session.pages.map(p=>[p.source.name,p.index]);
    const pdf=await PDFDocument.load(await session.build());assert.equal(pdf.getPageCount(),7);
    assert.deepEqual(order[0],['Transparência.png',0]);
    assert.equal(pdf.getPage(1).node.Resources().lookup(PDFName.of('Font')).keys().length,1);
    assert.equal(pdf.getPage(0).node.Resources().lookup(PDFName.of('XObject')).keys().length,1);
});
test('EXIF 1 a 8 usa matrizes sem distorção e não depende do DPI ausente/incorreto',async()=>{
    for(let orientation=1;orientation<=8;orientation++){
        const bytes=jpg(orientation,{dpi:1}), info=analisarImagem(bytes,'jpg');assert.equal(info.orientation,orientation);
        const {pdf}=await importarImagem(bytes,'jpg');const page=pdf.getPage(0);
        assert.equal(page.getWidth(),orientation>=5?86.4:57.6);assert.equal(page.getHeight(),orientation>=5?57.6:86.4);
        assert.equal(matrizImagem(57.6,86.4,orientation).length,6);
        assert.equal(analisarImagem(jpg(orientation,{dpi:600}),'jpg').width,info.width);
    }
});
test('PNG conserva máscara de transparência e JPEG mantém os dados comprimidos originais',async()=>{
    const {pdf}=await importarImagem(png(),'png');await pdf.save();
    const images=pdf.context.enumerateIndirectObjects().map(([,o])=>o).filter(o=>o.dict?.get(PDFName.of('Subtype'))?.toString()==='/Image');
    assert.ok(images.some(o=>o.dict.has(PDFName.of('SMask'))));
    const original=jpg(6), imported=await importarImagem(original,'jpg');await imported.pdf.save();
    const image=imported.pdf.context.enumerateIndirectObjects().map(([,o])=>o).find(o=>o.dict?.get(PDFName.of('Filter'))?.toString()==='/DCTDecode');
    assert.deepEqual(image.getContents(),new Uint8Array(original));
});
test('perfis ICC, CMYK e arquivos corrompidos não entram silenciosamente no lote',async()=>{
    assert.throws(()=>analisarImagem(jpg(1,{icc:true}),'jpg'),/ICC/);assert.throws(()=>analisarImagem(png({profile:true}),'png'),/ICC/);
    const cmyk=Buffer.from(jpg());for(let p=0;p<cmyk.length-12;p++)if(cmyk[p]===255&&[192,193,194].includes(cmyk[p+1])){cmyk[p+9]=4;break;}
    assert.throws(()=>analisarImagem(cmyk,'jpg'),/CMYK/);
    const session=new PDFCombinacao();await session.add([(await fixtures())[0]]);const ids=session.pages.map(p=>p.id);
    for(const file of [new File(['bad'],'corrompido.jpg'),new File(['bad'],'incompatível.txt')])await assert.rejects(session.add([file]));
    assert.deepEqual(session.pages.map(p=>p.id),ids);
    await assert.rejects(session.add([{name:'grande.jpg',size:513*1024*1024}]),/512/);
});
test('remoção por página/arquivo e movimento inválido preservam o rascunho',async()=>{
    const session=new PDFCombinacao();await session.add(await fixtures());
    const ids=session.pages.map(p=>p.id);assert.throws(()=>session.move(0,999),/inválida/);assert.deepEqual(session.pages.map(p=>p.id),ids);
    session.remove(session.pages[0].id);assert.equal(session.pages.length,6);
    session.remove(session.pages[1].id,true);assert.equal(session.pages.length,4);
    assert.equal((await PDFDocument.load(await session.build())).getPageCount(),4);
});
test('novo combinado marca pendência; acrescentar preserva giro/exclusão e exportação quita a versão atual',async()=>{
    const session=new PDFCombinacao();const files=await fixtures();await session.add(files.slice(0,3));
    const workspace=new PDFWorkspace();await workspace.abrir(files[0]);workspace.girarPagina(1,90);workspace.excluirPaginas([2]);
    const result=new File([await session.build()],'combinado.pdf');await workspace.adicionarArquivos([result],'combinado.pdf',{fontes:session.origins()});
    assert.equal(workspace.getTotalPaginas(),6);assert.equal(workspace.pdf.getPage(0).getRotation().angle,90);assert.equal(workspace.alterado,true);
    const version=workspace.capturarVersao();await workspace.obterBytes();assert.equal(workspace.marcarExportado(version),true);
    await workspace.adicionarArquivos([result],'novo.pdf',{substituir:true,fontes:session.origins()});assert.equal(workspace.getTotalPaginas(),5);assert.equal(workspace.nomeArquivo,'novo.pdf');assert.equal(workspace.alterado,true);
});
for(const modo of ['ajustar','preencher'])test(`A4 ${modo} reutiliza transformação vetorial e preserva PNG/JPEG`,async()=>{
    const session=new PDFCombinacao();await session.add(await fixtures());
    const workspace=new PDFWorkspace();await workspace.adicionarArquivos([new File([await session.build()],'lote.pdf')]);
    const prep=await workspace.prepararRedimensionamento({modo,largura:210,altura:297,paginas:[1,2,3,4,5,6,7]});workspace.aplicarRedimensionamento(prep);
    const result=await PDFDocument.load(await workspace.obterBytes());assert.equal(result.getPageCount(),7);
    for(const page of result.getPages()){assert.ok(Math.abs(page.getWidth()-595.2756)<.01);assert.ok(Math.abs(page.getHeight()-841.8898)<.01);}
    assert.ok(result.getPage(0).node.Resources().has(PDFName.of('Font')));
});
test('origens assinadas mantêm diagnóstico e identidade visual após exclusão/reordenação do rascunho',async()=>{
    const session=new PDFCombinacao();await session.add([new File([await criarPDFAssinado({total:3})],'assinada.pdf')]);session.move(0,2);
    const workspace=new PDFWorkspace();await workspace.adicionarArquivos([new File([await session.build()],'lote.pdf')],'lote.pdf',{fontes:session.origins()});
    assert.equal(workspace.obterResumoAssinaturas().presentes,1);assert.ok(workspace.obterResumoAssinaturas().paginasVisuais.length);
});

test('350 páginas e PNG de 3 megapixels combinam dentro dos limites, com exclusão e movimento distante',async()=>{
    const pdf=await PDFDocument.create();for(let i=0;i<350;i++)pdf.addPage([400+i,600]);
    const session=new PDFCombinacao();await session.add([new File([await pdf.save()],'Grande 350 páginas.pdf'),new File([png({width:2048,height:1536})],'Imagem grande.png')]);
    session.move(4,49);session.remove(session.pages[0].id);
    const result=await PDFDocument.load(await session.build());assert.equal(result.getPageCount(),350);assert.equal(result.getPage(48).getWidth(),404);
    assert.equal(session.pages.at(-1).source.type,'PNG');
});

test('perfis globais PDF, camadas e formulários bloqueiam antes de modificar o rascunho',async()=>{
    const session=new PDFCombinacao();await session.add([(await fixtures())[0]]);
    for(const key of ['OutputIntents','OCProperties','form']){
        const pdf=await PDFDocument.create(),page=pdf.addPage();
        if(key==='form')pdf.getForm().createTextField('campo').addToPage(page);
        else pdf.catalog.set(PDFName.of(key),pdf.context.obj([]));
        await assert.rejects(session.add([new File([await pdf.save()],key+'.pdf')]),/perfil|camadas|Formulários/);
        assert.equal(session.pages.length,2);
    }
});
