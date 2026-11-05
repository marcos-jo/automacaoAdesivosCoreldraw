import test from 'node:test';import assert from 'node:assert/strict';
import {planejarMontagem,mmParaPt} from '../public/js/tools/pdf-imposition-layout.js';
import {analisarMontagem,sugerirAjusteMontagem,areaImprimivelConhecida} from '../public/js/tools/pdf-imposition-bounds.js';
import {gerarMontagem} from './helpers/pdf-imposition-modules.mjs';
import {PDFDocument,rgb,PDFName,decodePDFRawStream} from '../public/vendor/pdf-lib/pdf-lib.js';
const A4={largura:210,altura:297},zero={esquerda:0,direita:0,superior:0,inferior:0};
const base={modo:'nup',paginas:[1,2],largura:420,altura:297,colunas:2,linhas:1,escalaModo:'real',modoMarcas:'sem',margens:zero,centralizar:true};
const plan=(c=base,p=[A4,A4])=>planejarMontagem(c,p,{permitirExcesso:true});
const near=(a,b)=>assert.ok(Math.abs(a-b)<1e-6,`${a} != ${b}`);
test('2 A4 em A3 paisagem sem margem: encaixe exato, sem corte ou sobreposição',()=>{
    const p=plan(),g=analisarMontagem(p);assert.equal(g.requerConfirmacao,false);near(g.sheets[0].aproveitamento,100);near(p.folhas[0][1].x,210);near(p.folhas[0][0].escala,1);assert.ok(g.palco.x<0&&g.palco.width>420);
});
test('2 A4 com margens 10: excesso de margem separado do corte físico e ajuste máximo',()=>{
    const c={...base,margens:{esquerda:10,direita:10,superior:10,inferior:10}},p=plan(c),g=analisarMontagem(p);
    assert.deepEqual(g.sheets[0].foraFolha,[]);assert.equal(g.sheets[0].foraMargens.length,4);assert.equal(g.requerConfirmacao,true);near(p.folhas[0][0].largura,210);
    const fit=sugerirAjusteMontagem(p,c);near(fit.escalaMin,277/297*100);const after=plan(fit.config);assert.equal(analisarMontagem(after).requerConfirmacao,false);near(after.folhas[0][0].altura,277);
    assert.equal(analisarMontagem(plan({...fit.config,fatorMontagem:fit.fatorMontagem+.00001})).requerConfirmacao,true);
});
test('marcas individuais, externas e divisões: excesso visível e encaixe reserva o traço inteiro',()=>{
    for(const modoMarcas of ['individuais','grupo','bordas']){const c={...base,modoMarcas},p=plan(c),g=analisarMontagem(p);assert.ok(g.sheets[0].foraFolha.length);assert.ok(g.palco.width>420);const fit=sugerirAjusteMontagem(p,c),after=plan(fit.config);assert.equal(analisarMontagem(after).requerConfirmacao,false);assert.ok(fit.escalaMin<100);}
});
test('Step & Repeat 120%, giro, espaçamento zero e excesso não mudam a escala solicitada',()=>{
    for(const rotacao of [0,90,270])for(const modoMarcas of ['sem','individuais','grupo','bordas']){const c={...base,modo:'repeat',paginas:[1],quantidade:2,escalaModo:'personalizada',escala:120,rotacao,modoMarcas};const p=plan(c,[A4]);near(p.folhas[0][0].escala,1.2);assert.ok(analisarMontagem(p).requerConfirmacao);const fit=sugerirAjusteMontagem(p,c);assert.equal(analisarMontagem(plan(fit.config,[A4])).requerConfirmacao,false);near(p.folhas[0][0].largura/p.folhas[0][0].altura,rotacao?297/210:210/297);}
});
test('caminho de inspeção preserva exatamente os planos válidos e a regressão A3 → A6',()=>{
    let checked=0;for(const modo of ['repeat','nup'])for(const modoMarcas of ['sem','individuais','grupo','bordas'])for(const rotacao of [0,90,180,270])for(const ordem of ['horizontal','vertical'])for(const escalaModo of ['real','ajustar','personalizada']){
        const c={...base,largura:450,altura:450,modo,modoMarcas,rotacao,ordem,escalaModo,escala:87.5,quantidade:2},sources=modo==='repeat'?[{largura:90,altura:50}]:[{largura:90,altura:50},{largura:80,altura:40}];
        assert.deepEqual(plan(c,sources),planejarMontagem(c,sources));checked++;
    }assert.equal(checked,192);
    const c={...base,modo:'repeat',paginas:[1],automatico:true,quantidade:4,largura:297,altura:420,pecaLargura:105,pecaAltura:148.4848484848485},p=plan(c,[{largura:297,altura:420}]);near(p.folhas[0][0].largura,105);near(p.folhas[0][0].altura,420*105/297);assert.deepEqual(p,planejarMontagem(c,[{largura:297,altura:420}]));
});
test('grade automática grande mostra pelo menos uma peça; limites e sangria continuam protegidos',()=>{
    const c={...base,modo:'repeat',automatico:true,quantidade:1,largura:100,altura:100};assert.equal(plan(c,[A4]).capacidade,1);
    assert.throws(()=>plan({...base,largura:NaN}),/valor fora/);assert.throws(()=>plan({...base,modo:'repeat',pecaLargura:90,pecaAltura:90},[A4]),/deformaria/);
    assert.throws(()=>plan({...base,sangria:3}),/não existe/);const s={...A4,sangriaDisponivel:3},p=plan({...base,sangria:3,modoMarcas:'grupo'},[s,s]);assert.throws(()=>sugerirAjusteMontagem(p,{...base,sangria:3,modoMarcas:'grupo'}),/sangria/);
    assert.throws(()=>sugerirAjusteMontagem(plan({...base,espacoX:1000}),{...base,espacoX:1000}),/consomem/);
});
test('área imprimível só é sobreposta para dimensões correspondentes; coordenadas top-left convertidas',()=>{
    const p=plan(),source={printer:'Fila sintética',geometry:{width:420,height:297,left:4,top:5,printableWidth:411,printableHeight:285}};
    const g=analisarMontagem(p,source);assert.equal(g.printable.y,7);assert.equal(g.printable.printer,source.printer);assert.equal(g.sheets[0].foraImprimivel.length,4);
    assert.equal(areaImprimivelConhecida(p,{...source.geometry,width:297,height:420}),null);assert.equal(areaImprimivelConhecida(p,{...source.geometry,left:NaN}),null);assert.equal(analisarMontagem(p).printable,null);
});
test('PDF de produção fica físico; prévia expande caixas sem mudar conteúdo, vetores ou matrizes',async()=>{
    const doc=await PDFDocument.create();for(let i=0;i<2;i++){const p=doc.addPage([mmParaPt(210),mmParaPt(297)]);p.drawRectangle({x:0,y:0,width:p.getWidth(),height:p.getHeight(),color:rgb(.9,.2,.4)});p.drawText('Texto vetorial A4',{x:30,y:60});}const source=await doc.save(),original=source.slice();
    const c={...base,escalaModo:'personalizada',escala:120,modoMarcas:'bordas'};
    await assert.rejects(gerarMontagem(source,c),/não cabe/);
    const result=await gerarMontagem(source,c,{previsualizarExcesso:true}),output=await PDFDocument.load(result.bytes),board=await PDFDocument.load(result.bytesPrevia);
    near(output.getPage(0).getWidth(),mmParaPt(420));near(output.getPage(0).getHeight(),mmParaPt(297));assert.ok(board.getPage(0).getWidth()>output.getPage(0).getWidth());assert.ok(board.getPage(0).getMediaBox().x<0);assert.deepEqual(source,original);
    const content=pdf=>pdf.getPage(0).node.Contents().asArray().map(ref=>Buffer.from(decodePDFRawStream(pdf.context.lookup(ref)).decode()).toString()).join('\n');assert.equal(content(output),content(board));assert.match(content(output),/Do/);
    assert.ok(output.context.enumerateIndirectObjects().some(([,o])=>o.dict?.get(PDFName.of('Subtype'))?.toString()==='/Form'));
});
