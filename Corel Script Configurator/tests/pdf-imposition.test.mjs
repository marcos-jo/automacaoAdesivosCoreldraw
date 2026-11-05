import test from 'node:test';
import assert from 'node:assert/strict';
import { planejarMontagem, mmParaPt } from '../public/js/tools/pdf-imposition-layout.js';
const card = { id: 1, largura: 90, altura: 50 };
const base = { modo: 'repeat', largura: 297, altura: 420, automatico: true, quantidade: 45, marcas: true };
test('cartão 90×50 em A3: reserva de marcas, folhas e medidas precisas', () => {
    const p = planejarMontagem(base, [card]);
    assert.equal(p.colunas, 2); assert.equal(p.linhas, 6); assert.equal(p.capacidade, 12); assert.equal(p.totalFolhas, 4);
    assert.equal(p.folhas.flat().length, 45); assert.equal(mmParaPt(25.4), 72);
    for (const piece of p.folhas.flat()) { assert.equal(piece.largura, 90); assert.equal(piece.altura, 50); assert.equal(piece.cortes.length, 8);
        for (const mark of piece.cortes) { assert.ok(Math.abs(Math.hypot(mark.x2-mark.x1, mark.y2-mark.y1)-3)<1e-9); for (const [x,y] of [[mark.x1,mark.y1],[mark.x2,mark.y2]]) { assert.ok(x>=0&&x<=297&&y>=0&&y<=420); } }
    }
});
test('N-up não repete páginas, mantém vagas, ordem horizontal e vertical', () => {
    const config = { ...base, modo: 'nup', colunas: 2, linhas: 2, marcas: false };
    const fontes = [card, null, { ...card, id: 3 }, { ...card, id: 4 }, { ...card, id: 5 }];
    const p = planejarMontagem(config, fontes); assert.equal(p.totalFolhas,2); assert.deepEqual(p.folhas.flat().map(p=>p.origem),[1,3,4,5]); assert.equal(p.folhas[0][1].indice,2);
    const vertical = planejarMontagem({...config,ordem:'vertical'}, [card,{...card,id:2},card,card]);
    assert.equal(vertical.folhas[0][0].x,vertical.folhas[0][1].x); assert.ok(vertical.folhas[0][0].y>vertical.folhas[0][1].y);
});
test('A6, mistos A4/A3, margens assimétricas e escala proporcional', () => {
    for (const grade of [[1,2],[2,2],[2,3],[2,4]]) {
        const p=planejarMontagem({...base,modo:'nup',colunas:grade[0],linhas:grade[1],margens:{esquerda:13,direita:7,superior:11,inferior:5}},[{largura:210,altura:297},{largura:297,altura:420},{largura:105,altura:148}]);
        assert.equal(p.util.largura,277); assert.equal(p.util.altura,404);
        for(const piece of p.folhas.flat())assert.ok(piece.x>=13&&piece.y>=5&&piece.x+piece.largura<=290+1e-7&&piece.y+piece.altura<=409+1e-7);
    }
});
test('100%, 120%, escala decimal e rotação mantêm proporção', () => {
    for(const percent of [100,120,'87,5']) {
        const p=planejarMontagem({...base,marcas:false,escalaModo:'personalizada',escala:percent,rotacao:90},[card]);
        const first=p.folhas[0][0];assert.equal(first.largura,50*Number(String(percent).replace(',','.'))/100);assert.equal(first.altura,90*Number(String(percent).replace(',','.'))/100);
    }
});
test('sangria ausente, marcas invasivas, deformação e grade impossível são recusadas', () => {
    assert.throws(()=>planejarMontagem({...base,sangria:3},[card]),/não existe/);
    assert.throws(()=>planejarMontagem({...base,sangria:3,afastamento:2},[{...card,sangriaDisponivel:3}]),/Afaste/);
    assert.doesNotThrow(()=>planejarMontagem({...base,sangria:3},[{...card,sangriaDisponivel:3}]));
    assert.throws(()=>planejarMontagem({...base,pecaLargura:80,pecaAltura:50},[card]),/deformaria/);
    assert.throws(()=>planejarMontagem({...base,automatico:false,colunas:10,linhas:10},[card]),/excede/);
    assert.throws(()=>planejarMontagem({...base,modo:'nup',colunas:2,linhas:2,escalaModo:'real'},[{largura:297,altura:420}]),/não cabe/);
});
test('erros não alteram entrada e limites impedem lotes sem controle', () => {
    const original=JSON.stringify(base);assert.throws(()=>planejarMontagem({...base,quantidade:10001},[card]));assert.equal(JSON.stringify(base),original);
    for(const value of [NaN,Infinity,-1,'abc'])assert.throws(()=>planejarMontagem({...base,largura:value},[card]));
    assert.throws(()=>planejarMontagem({...base,modo:'repeat'},[card,card]),/exatamente/);
});
test('sangria física é conferida depois da escala, sem inventar pixels na borda', () => {
    assert.throws(()=>planejarMontagem({...base,escalaModo:'personalizada',escala:50,sangria:3},[{...card,sangriaDisponivel:3}]),/escala reduz/);
    assert.doesNotThrow(()=>planejarMontagem({...base,escalaModo:'personalizada',escala:50,sangria:3},[{...card,sangriaDisponivel:6}]));
});

import {readFile} from 'node:fs/promises';
import {PDFDocument, PDFName, degrees, rgb, decodePDFRawStream} from '../public/vendor/pdf-lib/pdf-lib.js';
import {gerarMontagem} from './helpers/pdf-imposition-modules.mjs';
async function fixture(){const pdf=await PDFDocument.create();for(const [i,[w,h]] of [[90,50],[210,297],[297,420],[105,148]].entries()){const p=pdf.addPage([mmParaPt(w),mmParaPt(h)]);p.drawText('Pagina '+(i+1),{x:10,y:20,size:12});p.drawRectangle({x:20,y:30,width:40,height:20,color:rgb(.6,0,.8),opacity:.5});}return pdf;}
test('motor vetorial gera folhas, mantém fontes/grupo de transparência e original intacto',async()=>{
    const pdf=await fixture();pdf.getPage(0).node.set(PDFName.of('Group'),pdf.context.obj({S:'Transparency',CS:'DeviceRGB',I:true}));const bytes=await pdf.save(),original=bytes.slice();
    const result=await gerarMontagem(bytes,{...base,paginas:[1]});assert.deepEqual(bytes,original);
    const output=await PDFDocument.load(result.bytes);assert.equal(output.getPageCount(),4);assert.ok(Math.abs(output.getPage(0).getWidth()-mmParaPt(297))<1e-7);
    const stream=output.getPage(0).node.Contents().asArray().map(ref=>Buffer.from(decodePDFRawStream(output.context.lookup(ref)).decode()).toString()).join('\n');
    const lines=[...stream.matchAll(/(-?[\d.]+) (-?[\d.]+) m\s+(-?[\d.]+) (-?[\d.]+) l/g)];assert.equal(lines.length,12*8);
    const first=result.plan.folhas[0][0].cortes[0];for(const [index,key] of ['x1','y1','x2','y2'].entries())assert.ok(Math.abs(Number(lines[0][index+1])-mmParaPt(first[key]))<1e-6);
    const forms=output.context.enumerateIndirectObjects().map(([,o])=>o).filter(o=>o.dict?.get(PDFName.of('Subtype'))?.toString()==='/Form');assert.equal(forms.length,1);assert.ok(forms[0].dict.has(PDFName.of('Group')));assert.ok(forms[0].dict.lookup(PDFName.of('Resources')).has(PDFName.of('Font')));
    assert.equal(forms[0].dict.lookup(PDFName.of('Group')).get(PDFName.of('S')).toString(),'/Transparency');
});
test('motor N-up misto conserva sequência, rotações e múltiplas folhas',async()=>{
    const pdf=await fixture();pdf.getPage(1).setRotation(degrees(90));const bytes=await pdf.save();
    const result=await gerarMontagem(bytes,{...base,modo:'nup',colunas:2,linhas:2,paginas:[2,3,null,4,1]});
    assert.equal(result.plan.totalFolhas,2);assert.deepEqual(result.plan.folhas.flat().map(p=>p.origem),[2,3,4,1]);assert.ok(result.plan.folhas[0][0].largura>result.plan.folhas[0][0].altura);
    const out=await PDFDocument.load(result.bytes);assert.equal(out.getPageCount(),2);assert.ok(result.avisos.some(s=>s.includes('TrimBox')));
});
test('TrimBox/BleedBox explícitas: corte e sangria reais, sem fallback inválido',async()=>{
    const pdf=await fixture(),p=pdf.getPage(0);p.setTrimBox(mmParaPt(3),mmParaPt(3),mmParaPt(84),mmParaPt(44));p.setBleedBox(0,0,mmParaPt(90),mmParaPt(50));
    const result=await gerarMontagem(await pdf.save(),{...base,paginas:[1],sangria:3});assert.ok(Math.abs(result.plan.folhas[0][0].largura-84)<1e-7);
    p.setCropBox(20,20,10,10);await assert.rejects(gerarMontagem(await pdf.save(),{...base,paginas:[1]}),/caixas/);
});
test('aparências/cores complexas são bloqueadas e cancelamento não publica resultado',async()=>{
    for(const key of ['AcroForm','OCProperties','OutputIntents','StructTreeRoot']){const pdf=await fixture();pdf.catalog.set(PDFName.of(key),pdf.context.obj({}));await assert.rejects(gerarMontagem(await pdf.save(),{...base,paginas:[1]}));}
    const pdf=await fixture();pdf.getPage(0).node.set(PDFName.of('Annots'),pdf.context.obj([pdf.context.obj({Subtype:'Text'})]));await assert.rejects(gerarMontagem(await pdf.save(),{...base,paginas:[1]}),error=>error.diagnostico?.bloqueios.some(b=>b.codigo==='anotacao-montagem'));
    const abort=new AbortController();abort.abort();await assert.rejects(gerarMontagem(await pdf.save(),{...base,paginas:[1]},{signal:abort.signal}),{name:'AbortError'});
});

test('rotação automática N-up escolhe ajuste maior sem deformar',()=>{
    const config={...base,modo:'nup',colunas:1,linhas:2,marcas:false};
    const plain=planejarMontagem(config,[{largura:210,altura:297}]);
    const auto=planejarMontagem({...config,rotacaoAutomatica:true},[{largura:210,altura:297}]);
    assert.ok(auto.folhas[0][0].escala>plain.folhas[0][0].escala);assert.equal(auto.folhas[0][0].rotacao,90);
});
test('montagem usa workspace com união/giro/exclusão/ordem/redimensionamento pendentes, sem quitá-lo',async()=>{
    const {PDFWorkspace}=await import('./helpers/pdf-resize-modules.mjs');const doc=await fixture(),bytes=await doc.save();
    const arquivo={name:'sintetico.pdf',arrayBuffer:async()=>bytes.slice().buffer};const ws=new PDFWorkspace();await ws.abrir(arquivo);await ws.adicionarArquivos([arquivo]);ws.excluirPaginas([8]);ws.girarPagina(1,90);await ws.moverPagina(1,3);
    const prep=await ws.prepararRedimensionamento({paginas:[1],largura:100,altura:150,modo:'ajustar'});ws.aplicarRedimensionamento(prep);
    const version=ws.capturarVersao(),snapshot=await ws.obterBytes();const result=await gerarMontagem(snapshot,{...base,modo:'nup',colunas:2,linhas:2,paginas:[1,2,3,4,5,6,7]});
    assert.equal(result.plan.totalFolhas,2);assert.equal(ws.pdf,version.pdf);assert.equal(ws.revisao,version.revisao);assert.equal(ws.alterado,true);assert.deepEqual(await ws.obterBytes(),snapshot);
});

test('individual mantém exatamente posições/escala do caso aprovado; grupo encosta sem marcas internas',()=>{
    const config={...base,automatico:false,colunas:2,linhas:2,quantidade:4,espacoX:0,espacoY:0};
    const legacy=planejarMontagem(config,[card]);const individual=planejarMontagem({...config,modoMarcas:'individuais'},[card]);assert.deepEqual(individual.folhas,legacy.folhas);assert.equal(individual.folhas[0].reduce((n,p)=>n+p.cortes.length,0),32);
    const group=planejarMontagem({...config,modoMarcas:'grupo'},[card]);const [a,b,c]=group.folhas[0];assert.ok(Math.abs(a.x+a.largura-b.x)<1e-9);assert.ok(Math.abs(c.y+c.altura-a.y)<1e-9);assert.ok(group.folhas[0].every(p=>p.cortes.length===0));assert.equal(group.cortesGrupo[0].length,8);
    const left=Math.min(...group.folhas[0].map(p=>p.x)),right=Math.max(...group.folhas[0].map(p=>p.x+p.largura)),bottom=Math.min(...group.folhas[0].map(p=>p.y)),top=Math.max(...group.folhas[0].map(p=>p.y+p.altura));
    for(const m of group.cortesGrupo[0]){assert.ok(m.x1<=left||m.x1>=right||m.y1<=bottom||m.y1>=top);for(const [x,y] of [[m.x1,m.y1],[m.x2,m.y2]])assert.ok(x>=10&&x<=287&&y>=10&&y<=410);}
});
test('grupo em N-up e folha parcial/vazia: somente contorno das peças presentes',()=>{
    const config={...base,modo:'nup',colunas:2,linhas:2,modoMarcas:'grupo',espacoX:0,espacoY:0};
    const p=planejarMontagem(config,[card,card,card,card,card]);assert.equal(p.totalFolhas,2);assert.deepEqual(p.cortesGrupo.map(m=>m.length),[8,8]);assert.ok(p.folhas.flat().every(p=>!p.cortes.length));
    const empty=planejarMontagem(config,[card,null,null,null,null,null,null,null]);assert.deepEqual(empty.cortesGrupo.map(m=>m.length),[8,0]);
    assert.equal(planejarMontagem({...config,modoMarcas:'sem'},[card]).cortesGrupo[0].length,0);
    assert.throws(()=>planejarMontagem({...config,modoMarcas:'outro'},[card]),/Modo de marcas/);
});
test('PDF grupo contém exatamente oito segmentos externos, sem multiplicação por peça',async()=>{
    const pdf=await fixture();const result=await gerarMontagem(await pdf.save(),{...base,quantidade:4,modoMarcas:'grupo',paginas:[1]});const output=await PDFDocument.load(result.bytes);
    const stream=output.getPage(0).node.Contents().asArray().map(ref=>Buffer.from(decodePDFRawStream(output.context.lookup(ref)).decode()).toString()).join('\n');const lines=[...stream.matchAll(/(-?[\d.]+) (-?[\d.]+) m\s+(-?[\d.]+) (-?[\d.]+) l/g)];assert.equal(lines.length,8);
    for(const [i,line] of result.plan.cortesGrupo[0].entries())for(const [j,key] of ['x1','y1','x2','y2'].entries())assert.ok(Math.abs(Number(lines[i][j+1])-mmParaPt(line[key]))<1e-6);
});
test('regressão física A3 → A6 → Step & Repeat preserva dimensões e escala',async()=>{
    const {PDFWorkspace}=await import('./helpers/pdf-resize-modules.mjs');const pdf=await PDFDocument.create();pdf.addPage([mmParaPt(297),mmParaPt(420)]).drawText('A3 para A6',{x:20,y:40});const bytes=await pdf.save();const ws=new PDFWorkspace();await ws.abrir({name:'sintetico-a3.pdf',arrayBuffer:async()=>bytes.slice().buffer});const prep=await ws.prepararRedimensionamento({paginas:[1],largura:105,altura:148,modo:'ajustar'});ws.aplicarRedimensionamento(prep);const version=ws.capturarVersao();
    for(const mode of ['individuais','grupo','sem']){const result=await gerarMontagem(await ws.obterBytes(),{...base,quantidade:5,paginas:[1],modoMarcas:mode});for(const p of result.plan.folhas.flat()){assert.ok(Math.abs(p.largura-105)<1e-7);assert.ok(Math.abs(p.altura-148)<1e-7);assert.ok(Math.abs(p.escala-1)<1e-7);}assert.equal(ws.revisao,version.revisao);assert.equal(ws.alterado,true);}
});

test('bordas 2×1: cantos e divisão comum no topo/embaixo, sem duplicação',()=>{
    const config={...base,automatico:false,colunas:2,linhas:1,quantidade:2,modoMarcas:'bordas',espacoX:0,espacoY:0};const p=planejarMontagem(config,[card]);const [a,b]=p.folhas[0];assert.ok(Math.abs(a.x+a.largura-b.x)<1e-8);const marks=p.cortesGrupo[0];assert.equal(marks.length,10);assert.ok(p.folhas[0].every(piece=>piece.cortes.length===0));
    const division=a.x+a.largura,vertical=marks.filter(m=>Math.abs(m.x1-division)<1e-8&&m.x1===m.x2);assert.equal(vertical.length,2);assert.ok(vertical.some(m=>m.y1>a.y+a.altura));assert.ok(vertical.some(m=>m.y1<a.y));
    assert.equal(new Set(marks.map(m=>JSON.stringify(m))).size,10);
});
test('bordas em várias linhas/colunas projetam todos os cortes somente fora do bloco',()=>{
    for(const [cols,rows] of [[1,3],[2,2],[3,2],[4,3]]){const p=planejarMontagem({...base,largura:600,altura:600,automatico:false,colunas:cols,linhas:rows,quantidade:cols*rows,modoMarcas:'bordas'},[card]);const pieces=p.folhas[0],marks=p.cortesGrupo[0];assert.equal(marks.length,2*(cols+1)+2*(rows+1));
        const left=Math.min(...pieces.map(p=>p.x)),right=Math.max(...pieces.map(p=>p.x+p.largura)),bottom=Math.min(...pieces.map(p=>p.y)),top=Math.max(...pieces.map(p=>p.y+p.altura));
        for(const m of marks){assert.ok(m.y1<=bottom&&m.y2<=bottom||m.y1>=top&&m.y2>=top||m.x1<=left&&m.x2<=left||m.x1>=right&&m.x2>=right);assert.ok(Math.abs(Math.hypot(m.x2-m.x1,m.y2-m.y1)-3)<1e-8);for(const [x,y] of [[m.x1,m.y1],[m.x2,m.y2]])assert.ok(x>=10&&x<=590&&y>=10&&y<=590);}
    }
});
test('bordas preserva posições do modo grupo; intervalos positivos marcam os dois cortes',()=>{
    for(const mode of ['repeat','nup']){const config={...base,modo:mode,automatico:false,colunas:2,linhas:2,quantidade:4,espacoX:2,espacoY:3};const sources=mode==='repeat'?[card]:[card,{largura:105,altura:148},{largura:210,altura:297},card];const corners=planejarMontagem({...config,modoMarcas:'grupo'},sources),divisions=planejarMontagem({...config,modoMarcas:'bordas'},sources);assert.deepEqual(divisions.folhas,corners.folhas);assert.equal(corners.cortesGrupo[0].length,8);assert.ok(divisions.cortesGrupo[0].length>8);}
    const p=planejarMontagem({...base,automatico:false,colunas:2,linhas:1,quantidade:2,espacoX:2,modoMarcas:'bordas'},[card]);assert.equal(p.cortesGrupo[0].length,12);
});
test('bordas N-up uniforme, vagas e folha parcial não criam marcas de peças inexistentes',()=>{
    const config={...base,modo:'nup',colunas:2,linhas:2,largura:212.1,altura:132.1,modoMarcas:'bordas'};const p=planejarMontagem(config,[card,card,card,card,card]);assert.equal(p.cortesGrupo[0].length,12);assert.equal(p.cortesGrupo[1].length,8);
    const blank=planejarMontagem(config,[card,null,null,null,null,null,null,null]);assert.deepEqual(blank.cortesGrupo.map(m=>m.length),[8,0]);
});
test('PDF 2×1 contém exatamente dez marcas vetoriais nas coordenadas de contorno e divisão',async()=>{
    const pdf=await fixture();const result=await gerarMontagem(await pdf.save(),{...base,paginas:[1],automatico:false,colunas:2,linhas:1,quantidade:2,modoMarcas:'bordas'});const output=await PDFDocument.load(result.bytes);const stream=output.getPage(0).node.Contents().asArray().map(ref=>Buffer.from(decodePDFRawStream(output.context.lookup(ref)).decode()).toString()).join('\n');const lines=[...stream.matchAll(/(-?[\d.]+) (-?[\d.]+) m\s+(-?[\d.]+) (-?[\d.]+) l/g)];assert.equal(lines.length,10);for(const [i,line] of result.plan.cortesGrupo[0].entries())for(const [j,key] of ['x1','y1','x2','y2'].entries())assert.ok(Math.abs(Number(lines[i][j+1])-mmParaPt(line[key]))<1e-6);
});
