import test from 'node:test';
import assert from 'node:assert/strict';
import {paginarLivreto,planejarLivreto,girarRetangulo} from '../public/js/tools/pdf-booklet-layout.js';
const fontes=(n,w=148,h=210)=>Array.from({length:n},(_,i)=>({id:i+1,largura:w,altura:h,sangriaDisponivel:3}));
const near=(a,b)=>assert.ok(Math.abs(a-b)<1e-7,`${a} != ${b}`);
for(const count of [4,8,12,16,5,7,10,1,2000])test(`livreto ${count}: paginação, brancos finais e simulação de dobra`,()=>{
    const p=paginarLivreto(count);assert.equal(p.total%4,0);assert.equal(p.folhas.length,p.total/4);
    assert.deepEqual(p.leitura.map(s=>s.logica),Array.from({length:p.total},(_,i)=>i+1));
    assert.deepEqual(p.leitura.filter(s=>!s.origem).map(s=>s.logica),p.brancas);
    assert.deepEqual(p.aberturas.flat().filter(Boolean),p.leitura);
    for(const f of p.folhas){assert.equal(f.frente[0].logica+f.frente[1].logica,p.total+1);assert.equal(f.verso[0].logica+f.verso[1].logica,p.total+1);assert.equal(f.frente[1].logica+1,f.verso[0].logica);}
});
test('oito páginas: ordem física de referência',()=>{assert.deepEqual(paginarLivreto(8).folhas.map(f=>[f.frente.map(s=>s.origem),f.verso.map(s=>s.origem)]),[[[8,1],[2,7]],[[6,3],[4,5]]]);});
for(const orientacao of ['paisagem','retrato'])for(const verso of ['normal','girar180'])test(`virada física de referência associa frente/verso da mesma página: ${orientacao}/${verso}`,()=>{
    const p=planejarLivreto({orientacao,verso},fontes(16));
    const flipX=(orientacao==='paisagem')===(verso==='normal');
    for(let i=0;i<p.faces.length;i+=2){
        const a=p.faces[i].slots[1],b=p.faces[i+1].slots[0],r=a.moldura,s=b.moldura;
        near(flipX?p.folha.largura-r.x-r.width:r.x,s.x);near(flipX?r.y:p.folha.altura-r.y-r.height,s.y);near(r.width,s.width);near(r.height,s.height);
        assert.equal(a.logica+1,b.logica);
    }
});
for(const orientacao of ['paisagem','retrato'])for(const verso of ['normal','girar180'])test(`faces alinhadas e leitura inversa: ${orientacao}/${verso}`,()=>{
    const p=planejarLivreto({orientacao,verso},fontes(8,140,200));assert.deepEqual(p.bloqueios,[]);
    for(const [index,face]of p.faces.entries())for(const [slotIndex,slot]of face.slots.entries()){
        const piece=p.folhas[index][slotIndex],canonical=girarRetangulo({x:piece.x,y:piece.y,width:piece.largura,height:piece.altura},p.folha.largura,p.folha.altura,(4-face.giro/90)%4);
        near(canonical.x,(slotIndex?148.5:.5)+4);near(canonical.y,5);
        assert.equal(piece.rotacao,face.giro);assert.equal(piece.origem,slot.origem);
    }
    assert.deepEqual(p.leitura.folhas.flat().map(s=>s.origem),[1,2,3,4,5,6,7,8]);
    for(const s of p.leitura.folhas.flat()){near(s.largura,140);near(s.altura,200);assert.equal(s.rotacao,0);}
});
test('A5/A4 nominal e A4/A3 a 100%, sem reduzir silenciosamente',()=>{
    const a=planejarLivreto({},fontes(4));assert.deepEqual(a.bloqueios,[]);assert.equal(a.aposDobra.largura,148.5);assert.equal(a.final.largura,148);assert.equal(a.folhas[0][0].escala,1);
    const b=planejarLivreto({final:'A4',papel:'A3'},fontes(4,210,297));assert.deepEqual(b.bloqueios,[]);near(b.folhas[0][0].largura,210);
    assert.ok(planejarLivreto({final:'A4',papel:'A4'},fontes(4,210,297)).bloqueios.length);
    assert.ok(planejarLivreto({},fontes(4,210,297)).bloqueios.length);
    const c=planejarLivreto({escalaModo:'ajustar'},fontes(4,210,297));assert.deepEqual(c.bloqueios,[]);near(c.folhas[0][0].escala,148/210);
});
test('margens, registro, giro, sangria e marcas externos não alteram o formato final',()=>{
    const p=planejarLivreto({papel:'A3',margensConteudo:{esquerda:3,direita:3,superior:3,inferior:3},registroX:1,registroY:-1,sangria:1,marcasCorte:true,marcasDobra:true},fontes(4,140,200));
    assert.deepEqual(p.bloqueios,[]);assert.equal(p.cortesGrupo[0].length,10);assert.equal(p.final.largura,148);
    assert.equal(p.folhas[0][0].clip.x+p.folhas[0][0].clip.width,210);
    assert.equal(p.folhas[0][1].clip.x,210);
    assert.ok(planejarLivreto({registroX:1},fontes(4)).bloqueios.length);
    const rotated=planejarLivreto({rotacao:90},fontes(4,210,148));assert.deepEqual(rotated.bloqueios,[]);near(rotated.folhas[0][0].largura,148);
});
test('entrada inválida não aceita esquema parcial',()=>{
    for(const n of [0,-1,2.5,2001])assert.throws(()=>paginarLivreto(n));
    assert.throws(()=>planejarLivreto({escalaModo:'personalizada',escala:0},fontes(4)));
    assert.throws(()=>planejarLivreto({sangria:4},fontes(4)));
});
