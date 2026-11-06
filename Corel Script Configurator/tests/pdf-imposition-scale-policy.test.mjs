import test from 'node:test';import assert from 'node:assert/strict';
import {planejarMontagem} from '../public/js/tools/pdf-imposition-layout.js';
import {analisarMontagem,sugerirAjusteMontagem} from '../public/js/tools/pdf-imposition-bounds.js';
import {resumirMontagem} from '../public/js/tools/pdf-imposition-summary.js';
import {PDFDocument} from '../public/vendor/pdf-lib/pdf-lib.js';
import {gerarMontagem} from './helpers/pdf-imposition-modules.mjs';
import {conferirTransformacoes} from './helpers/pdf-imposition-proof.mjs';
const zero={esquerda:0,direita:0,superior:0,inferior:0},a6=n=>Array.from({length:n},(_,i)=>({id:i+1,largura:105,altura:148}));
const base={modo:'nup',largura:297,altura:420,colunas:2,linhas:2,margens:zero,modoMarcas:'sem'};
const near=(a,b)=>assert.ok(Math.abs(a-b)<1e-7,`${a} != ${b}`);
for(const distribution of ['celulas','compacto']){
 test(`três A6 em A3: padrão e tamanho real preservam 105×148 em ${distribution}`,()=>{
  for(const orientacao of [[297,420],[420,297]])for(const escalaModo of [undefined,'real','reduzir']){
   const p=planejarMontagem({...base,nupDistribuicao:distribution,largura:orientacao[0],altura:orientacao[1],escalaModo},a6(3),{permitirExcesso:true});
   assert.equal(p.totalFolhas,1);assert.equal(p.folhas[0].length,3);assert.deepEqual(p.folhas[0].map(p=>p.origem),[1,2,3]);for(const piece of p.folhas[0]){near(piece.escala,1);near(piece.largura,105);near(piece.altura,148);}assert.equal(analisarMontagem(p).requerConfirmacao,false);
  }
 });
 test(`quatro A6 em A4: ${distribution}, retrato/paisagem com giro autorizado`,()=>{
  for(const [largura,altura]of [[210,297],[297,210]])for(const escalaModo of ['real','reduzir']){
   const p=planejarMontagem({...base,nupDistribuicao:distribution,largura,altura,escalaModo,rotacaoAutomatica:true},a6(4));assert.equal(p.totalFolhas,1);assert.equal(p.folhas[0].length,4);
   for(const piece of p.folhas[0]){near(piece.escala,1);assert.deepEqual([piece.largura,piece.altura].sort((a,b)=>a-b),[105,148]);if(largura>altura)assert.equal(piece.rotacao,90);}
  }
 });
 test(`políticas explícitas em ${distribution}: reduzir não amplia; preencher permite ampliar`,()=>{
  const small=planejarMontagem({...base,nupDistribuicao:distribution,escalaModo:'reduzir'},a6(3));assert.ok(small.folhas.flat().every(p=>p.escala===1));
  assert.ok(sugerirAjusteMontagem(small,{...base,nupDistribuicao:distribution,escalaModo:'reduzir'}).escalaMax<=100);
  assert.ok(planejarMontagem({...base,nupDistribuicao:distribution,escalaModo:'reduzir',fatorMontagem:3},a6(3)).folhas.flat().every(p=>p.escala<=1));
  const fill=planejarMontagem({...base,nupDistribuicao:distribution,escalaModo:'preencher'},a6(3));assert.ok(fill.folhas.flat().every(p=>p.escala>1));
  const c={...base,nupDistribuicao:distribution,escalaModo:'preencher'},summary=resumirMontagem({config:c,plan:fill,avisos:[]});assert.equal(summary.ampliacao,true);assert.match(summary.confirmacao,/AMPLIAÇÃO INTENCIONAL/);assert.match(summary.basico,/Preencher células/);
  const tiny={...base,nupDistribuicao:distribution,largura:100,altura:140,escalaModo:'reduzir'};
  const reduced=planejarMontagem(tiny,a6(3));assert.ok(reduced.folhas.flat().every(p=>p.escala<1));assert.equal(analisarMontagem(reduced).requerConfirmacao,false);
  const real=planejarMontagem({...tiny,escalaModo:'real'},a6(3),{permitirExcesso:true});assert.ok(real.folhas.flat().every(p=>p.escala===1));assert.equal(analisarMontagem(real).requerConfirmacao,true);
  assert.throws(()=>planejarMontagem({...tiny,escalaModo:'real'},a6(3)),/excede|não cabe/);
 });
 test(`marcas e margens não reduzem tamanho real em ${distribution}`,()=>{
  for(const modoMarcas of ['sem','grupo','bordas']){
   const c={...base,nupDistribuicao:distribution,largura:210,altura:297,escalaModo:'real',modoMarcas},p=planejarMontagem(c,a6(4),{permitirExcesso:true});assert.ok(p.folhas.flat().every(p=>p.escala===1));assert.equal(analisarMontagem(p).requerConfirmacao,modoMarcas!=='sem');
   const large=planejarMontagem({...c,largura:320,altura:450},a6(4),{permitirExcesso:true});assert.ok(large.folhas.flat().every(p=>p.escala===1));assert.equal(analisarMontagem(large).requerConfirmacao,false);
  }
  const margin=planejarMontagem({...base,nupDistribuicao:distribution,largura:210,altura:297,margens:{esquerda:10,direita:10,superior:10,inferior:10},escalaModo:'real'},a6(4),{permitirExcesso:true});assert.ok(margin.folhas.flat().every(p=>p.escala===1));assert.equal(analisarMontagem(margin).requerConfirmacao,true);
 });
 test(`mistos em ${distribution}: dimensões de origem preservadas sem distorção`,()=>{
  const sources=[{id:1,largura:90,altura:50},{id:2,largura:105,altura:148},{id:3,largura:120,altura:80}],p=planejarMontagem({...base,nupDistribuicao:distribution,escalaModo:'real'},sources);
  for(const piece of p.folhas.flat()){const s=sources[piece.origem-1];near(piece.escala,1);near(piece.largura,s.largura);near(piece.altura,s.altura);}
 });
}
test('caso relatado: 135,6667% pertence ao preenchimento explícito; tamanho real continua A6',()=>{
 const c={...base,modoMarcas:'bordas'},fill=planejarMontagem({...c,escalaModo:'preencher'},a6(3)),real=planejarMontagem({...c,escalaModo:'real'},a6(3));
 near(fill.folhas[0][0].escala*100,135.66666666666666);near(fill.folhas[0][0].largura,142.45);near(fill.folhas[0][0].altura,200.78666666666666);near(real.folhas[0][0].escala,1);near(real.folhas[0][0].largura,105);near(real.folhas[0][0].altura,148);
});
test('PDFs serializados: matrizes e fontes físicas nas três políticas e distribuições',async()=>{
 const source=await PDFDocument.create();for(let i=1;i<=3;i++)source.addPage([105*72/25.4,148*72/25.4]).drawText('PAGE '+i,{x:20,y:60});const bytes=await source.save(),original=bytes.slice();
 for(const nupDistribuicao of ['celulas','compacto'])for(const escalaModo of ['real','reduzir','preencher'])for(const [largura,altura]of [[297,420],[100,140]]){
  const config={...base,nupDistribuicao,escalaModo,largura,altura,paginas:[1,2,3]},r=await gerarMontagem(bytes,config,{previsualizarExcesso:true});assert.equal(await conferirTransformacoes(r.bytes,r.plan),3);assert.deepEqual(bytes,original);
  const pdf=await PDFDocument.load(r.bytes);near(pdf.getPage(0).getWidth(),largura*72/25.4);near(pdf.getPage(0).getHeight(),altura*72/25.4);assert.ok(r.bytesPrevia.length);
 }
});
