import test from 'node:test';import assert from 'node:assert/strict';
import {planejarMontagem} from '../public/js/tools/pdf-imposition-layout.js';
import {analisarMontagem,sugerirAjusteMontagem} from '../public/js/tools/pdf-imposition-bounds.js';
import {resumirMontagem} from '../public/js/tools/pdf-imposition-summary.js';
import {gerarMontagem} from './helpers/pdf-imposition-modules.mjs';
import {conferirTransformacoes} from './helpers/pdf-imposition-proof.mjs';
import {PDFDocument,degrees} from '../public/vendor/pdf-lib/pdf-lib.js';
const near=(a,b)=>assert.ok(Math.abs(a-b)<1e-7,`${a} != ${b}`),zero={esquerda:0,direita:0,superior:0,inferior:0};
const base={modo:'nup',nupDistribuicao:'compacto',largura:210,altura:297,colunas:2,linhas:2,escalaModo:'ajustar',modoMarcas:'sem',margens:zero};
const a6=Array.from({length:4},(_,i)=>({id:i+1,largura:105,altura:148})),plan=(c=base,sources=a6)=>planejarMontagem(c,sources,{permitirExcesso:true});
function semSobreposicao(p){for(const pieces of p.folhas)for(let i=0;i<pieces.length;i++)for(let j=i+1;j<pieces.length;j++){const a=pieces[i],b=pieces[j],overlapX=Math.min(a.x+a.largura,b.x+b.largura)-Math.max(a.x,b.x),overlapY=Math.min(a.y+a.altura,b.y+b.altura)-Math.max(a.y,b.y);assert.ok(overlapX<1e-7||overlapY<1e-7,'Peças sobrepostas');}}
test('quatro A6 em A4 retrato: tamanho real, contato nas duas direções e 1 mm livre',()=>{
    for(const escalaModo of ['real','ajustar'])for(const rotacaoAutomatica of [false,true]){const p=plan({...base,escalaModo,rotacaoAutomatica}),[a,b,c]=p.folhas[0];near(a.escala,1);near(a.largura,105);near(a.altura,148);near(b.x,a.x+a.largura);near(c.y+c.altura,a.y);near(a.x,0);near(c.y,.5);assert.ok(p.folhas[0].every(p=>p.rotacao===0));assert.equal(analisarMontagem(p).requerConfirmacao,false);semSobreposicao(p);}
});
test('A4 paisagem: giro autorizado de 90° mantém 100%, desligado não inventa rotação',()=>{
    const p=plan({...base,largura:297,altura:210,rotacaoAutomatica:true});for(const piece of p.folhas[0]){near(piece.escala,1);near(piece.largura,148);near(piece.altura,105);assert.equal(piece.rotacao,90);}semSobreposicao(p);assert.equal(analisarMontagem(p).requerConfirmacao,false);
    const plain=plan({...base,largura:297,altura:210});assert.ok(plain.folhas[0].every(p=>p.rotacao===0));near(plain.folhas[0][0].escala,210/296);
    const real=plan({...base,largura:297,altura:210,escalaModo:'real'});near(real.folhas[0][0].escala,1);assert.equal(analisarMontagem(real).requerConfirmacao,true);
});
test('células reproduzem 66,8581% e folga; compacto mantém 100% sem reservar espaço para marcas',()=>{
    const c={...base,largura:297,altura:210,modoMarcas:'bordas'},cells=plan({...c,nupDistribuicao:'celulas'}),p=cells.folhas[0];near(p[0].escala,98.95/148);near(p[1].x-p[0].x-p[0].largura,72.2489864864865);
    const compact=plan({...c,rotacaoAutomatica:true});near(compact.folhas[0][0].escala,1);assert.ok(analisarMontagem(compact).sheets[0].foraFolha.length);assert.equal(compact.cortesGrupo[0].length,12);assert.throws(()=>planejarMontagem({...c,rotacaoAutomatica:true},a6),/Marcas externas excedem/);
});
test('marcas externas e divisões ficam fora do bloco, e ajuste explícito pode reservar seu espaço',()=>{
    for(const modoMarcas of ['grupo','bordas']){const c={...base,modoMarcas},p=plan(c);near(p.folhas[0][0].escala,1);const all=p.folhas[0],left=Math.min(...all.map(p=>p.x)),right=Math.max(...all.map(p=>p.x+p.largura)),bottom=Math.min(...all.map(p=>p.y)),top=Math.max(...all.map(p=>p.y+p.altura));
        for(const l of p.cortesGrupo[0])assert.ok(Math.max(l.x1,l.x2)<left||Math.min(l.x1,l.x2)>right||Math.max(l.y1,l.y2)<bottom||Math.min(l.y1,l.y2)>top);
        let fitted=sugerirAjusteMontagem(p,c),q=plan(fitted.config);near(q.folhas[0][0].escala,197.9/210);assert.equal(analisarMontagem(q).requerConfirmacao,false);for(let i=0;i<3;i++){fitted=sugerirAjusteMontagem(q,fitted.config);q=plan(fitted.config);near(q.folhas[0][0].escala,197.9/210);}
    }
});
test('mistos: contato por faixa sem células uniformes, ordem horizontal/vertical e intervalos reais',()=>{
    const sources=[{id:1,largura:80,altura:50},{id:2,largura:40,altura:90},{id:3,largura:60,altura:30},{id:4,largura:100,altura:45}];
    for(const ordem of ['horizontal','vertical'])for(const gap of [0,7]){const p=plan({...base,escalaModo:'real',ordem,espacoX:gap,espacoY:gap},sources),[a,b,c]=p.folhas[0];assert.deepEqual(p.folhas[0].map(p=>p.origem),[1,2,3,4]);near(a.largura,80);near(b.altura,90);if(ordem==='horizontal'){near(b.x-a.x-a.largura,gap);near(a.y+a.altura,b.y+b.altura);near(b.y-(c.y+c.altura),gap);}else{near(a.y-(b.y+b.altura),gap);near(a.x,b.x);near(c.x-(a.x+a.largura),gap);}semSobreposicao(p);}
});
test('sangrias e marcas individuais não atravessam peças; vagas só no tradicional',()=>{
    assert.throws(()=>plan({...base,modoMarcas:'individuais'}),error=>error.code==='COMPACT_MARKS');assert.doesNotThrow(()=>plan({...base,modoMarcas:'individuais',espacoX:7,espacoY:7,escalaModo:'real'}));
    const source=a6.map(p=>({...p,sangriaDisponivel:6}));assert.throws(()=>plan({...base,sangria:3},source),/não sobrepor sangrias/);
    const q=plan({...base,largura:320,altura:450,sangria:3,espacoX:6,espacoY:6,escalaModo:'real',modoMarcas:'bordas',afastamento:4},source);semSobreposicao(q);
    assert.throws(()=>plan(base,[a6[0],null]),/vagas 0/);assert.doesNotThrow(()=>plan({...base,nupDistribuicao:'celulas'},[a6[0],null]));
});
test('350 páginas e 350 posições: busca limitada e sem sobreposição ou permutação',()=>{
    const sources=Array.from({length:350},(_,i)=>({id:i+1,largura:20+(i%5),altura:30+(i%7)})),p=plan({...base,rotacaoAutomatica:true},sources);assert.equal(p.totalFolhas,88);assert.deepEqual(p.folhas.flat().map(p=>p.origem),sources.map(p=>p.id));semSobreposicao(p);
    const q=plan({...base,colunas:25,linhas:14,rotacaoAutomatica:true},sources);assert.equal(q.totalFolhas,1);semSobreposicao(q);assert.ok(q.folhas[0].every(p=>p.escala>0&&[0,90].includes(p.rotacao)));
});
test('PDFs gerados: quatro A6 e dez páginas mistas conservam matrizes, escala, giro e original',async()=>{
    const source=await PDFDocument.create();for(let i=0;i<10;i++){const w=i<4?105:80+i,h=i<4?148:40+i*2,p=source.addPage([w*72/25.4,h*72/25.4]);p.drawText('Vetor '+(i+1),{x:10,y:20});if(i===8)p.setRotation(degrees(90));}const bytes=await source.save(),original=bytes.slice();let checked=0;
    for(const [w,h]of [[210,297],[297,210]])for(const modoMarcas of ['sem','grupo','bordas']){const c={...base,largura:w,altura:h,paginas:[1,2,3,4],rotacaoAutomatica:true,modoMarcas},r=await gerarMontagem(bytes,c,{previsualizarExcesso:true});r.config=c;checked+=await conferirTransformacoes(r.bytes,r.plan);assert.match(resumirMontagem(r).basico,/escala efetiva: 100%/);assert.match(resumirMontagem(r).basico,w>h?/giro na montagem: 90°/:/giro na montagem: 0°/);}
    for(const ordem of ['horizontal','vertical'])for(const gap of [0,5]){const r=await gerarMontagem(bytes,{...base,ordem,espacoX:gap,espacoY:gap,paginas:Array.from({length:10},(_,i)=>i+1),rotacaoAutomatica:true,modoMarcas:'bordas'},{previsualizarExcesso:true});assert.equal(r.plan.totalFolhas,3);checked+=await conferirTransformacoes(r.bytes,r.plan);semSobreposicao(r.plan);assert.deepEqual(r.plan.folhas.flat().map(p=>p.origem),[1,2,3,4,5,6,7,8,9,10]);}
    assert.equal(checked,64);assert.deepEqual(bytes,original);
});
