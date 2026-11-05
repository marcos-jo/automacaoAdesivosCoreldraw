import test from 'node:test';import assert from 'node:assert/strict';
import {ConfiguracoesMontagem} from '../public/js/tools/pdf-imposition-settings.js';
import {resumirMontagem} from '../public/js/tools/pdf-imposition-summary.js';
import {gerarMontagem} from './helpers/pdf-imposition-modules.mjs';
import {conferirTransformacoes} from './helpers/pdf-imposition-proof.mjs';
import {PDFDocument,degrees} from '../public/vendor/pdf-lib/pdf-lib.js';
const near=(a,b)=>assert.ok(Math.abs(a-b)<1e-7,`${a} != ${b}`),margins=n=>({esquerda:n,direita:n,superior:n,inferior:n});
const base={modo:'repeat',paginas:[1],largura:297,altura:420,automatico:true,quantidade:1,escalaModo:'real',modoMarcas:'individuais',margens:margins(0)};
test('modos guardam escala, grade, origem, margens e ajuste próprios sem compartilhar referências',()=>{
    const original={grade:'auto',paginas:'1',escalaModo:'real',escala:'100',esquerda:'0'},settings=new ConfiguracoesMontagem(original);
    let nup=settings.trocar('nup',{...original,escalaModo:'personalizada',escala:'120'},1);assert.equal(nup.values.escalaModo,'ajustar');assert.equal(nup.values.grade,'4');assert.equal(nup.fator,1);
    const fitted={...nup.values,paginas:'1-4',grade:'custom',esquerda:'10'},grid={grade:'4',colunas:'2',linhas:'2'};
    for(let i=0;i<20;i++){const repeat=settings.trocar('repeat',fitted,.25,grid);assert.equal(repeat.values.grade,'auto');assert.equal(repeat.values.escala,'120');assert.equal(repeat.fator,1);nup=settings.trocar('nup',repeat.values,repeat.fator);assert.equal(nup.values.paginas,'1-4');assert.equal(nup.fator,.25);assert.deepEqual(nup.base,grid);}
    nup.values.escala='999';assert.notEqual(settings.estados.nup.values.escala,'999');settings.limparAjustes();assert.equal(settings.estados.nup.fator,1);assert.equal(settings.estados.nup.values.grade,'4');assert.throws(()=>settings.trocar('invalido',{}));
});
test('caso relatado: redução N-up é esperada, 120 inativo não aparece como escala aplicada',async()=>{
    const pdf=await PDFDocument.create();pdf.addPage([395.82*72/25.4,494.59*72/25.4]).drawText('Escala sintetica',{x:30,y:60});const source=await pdf.save();
    const r=await gerarMontagem(source,{...base,margens:margins(10),escalaModo:'ajustar',escala:120},{previsualizarExcesso:true});r.config={...base,escalaModo:'ajustar',escala:120};near(r.plan.folhas[0][0].escala,1);near(r.plan.folhas[0][0].largura,395.82);
    const c={...base,modo:'nup',automatico:false,colunas:2,linhas:2,margens:margins(10),escalaModo:'ajustar',escala:120},n=await gerarMontagem(source,c,{previsualizarExcesso:true});n.config=c;
    near(n.plan.folhas[0][0].largura,126.4);near(n.plan.folhas[0][0].altura,157.94092264160477);const summary=resumirMontagem(n);assert.equal(summary.manual,'não aplicada');assert.match(summary.automatica,/31,9337%/);assert.doesNotMatch(summary.basico,/120/);assert.equal(summary.rows[0].origem,'395,82 × 494,59 mm');assert.equal(await conferirTransformacoes(n.bytes,n.plan),1);assert.equal(await conferirTransformacoes(r.bytes,r.plan),1);
});
test('100%, 120%, ajuste e fator explícito: resumo e matrizes PDF coincidem em páginas diferentes/rotacionadas',async()=>{
    const pdf=await PDFDocument.create();for(const [i,[w,h]]of [[210,297],[297,420],[100,150],[200,100]].entries()){const p=pdf.addPage([w*72/25.4,h*72/25.4]);p.drawText('Pagina '+i,{x:20,y:30});if(i===2)p.setRotation(degrees(90));}const bytes=await pdf.save(),original=bytes.slice();let checked=0;
    for(const margin of [0,10])for(const modo of ['repeat','nup'])for(const escalaModo of ['real','personalizada','ajustar'])for(const fatorMontagem of [1,.8]){
        const c={...base,modo,paginas:modo==='repeat'?[1]:[1,2,3,4],automatico:modo==='repeat',colunas:2,linhas:2,escalaModo,escala:120,fatorMontagem,margens:margins(margin)},r=await gerarMontagem(bytes,c,{previsualizarExcesso:true});r.config=c;checked+=await conferirTransformacoes(r.bytes,r.plan);
        const summary=resumirMontagem(r);for(const row of summary.rows){const p=r.plan.folhas.flat().find(p=>p.origem===row.pagina);near(row.escala,p.escala*100);}assert.equal(summary.rows.length,modo==='repeat'?1:4);assert.deepEqual(bytes,original);
    }assert.equal(checked,60);
});
test('origem física e recorte não são confundidos; medidas desejadas substituem percentual com explicação',async()=>{
    const pdf=await PDFDocument.create(),p=pdf.addPage([210*72/25.4,297*72/25.4]);p.setTrimBox(5*72/25.4,5*72/25.4,200*72/25.4,287*72/25.4);p.drawText('Corte',{x:30,y:40});const c={...base,pecaLargura:100,escalaModo:'personalizada',escala:120},r=await gerarMontagem(await pdf.save(),c,{previsualizarExcesso:true});r.config=c;const summary=resumirMontagem(r);assert.equal(summary.rows[0].origem,'210 × 297 mm');assert.equal(summary.rows[0].corte,'200 × 287 mm');assert.equal(summary.rows[0].final,'100 × 143,5 mm');assert.equal(summary.rows[0].escala,50);assert.match(summary.manual,/tamanho final informado/);assert.equal(await conferirTransformacoes(r.bytes,r.plan),1);
});
