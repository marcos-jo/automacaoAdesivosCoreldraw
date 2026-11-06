import test from 'node:test';
import assert from 'node:assert/strict';
import {selecionarPaginasMontagem as select} from '../public/js/tools/pdf-imposition-selection.js';
import {ConfiguracoesMontagem} from '../public/js/tools/pdf-imposition-settings.js';
import {PDFDocument} from '../public/vendor/pdf-lib/pdf-lib.js';
import {gerarMontagem} from './helpers/pdf-imposition-modules.mjs';
import {conferirTransformacoes} from './helpers/pdf-imposition-proof.mjs';
const all=total=>Array.from({length:total},(_,i)=>i+1);
for(const total of [1,4,5,8,10]){
    test(`seleção ${total}: todas por padrão, intervalo inválido não é truncado`,()=>{
        assert.deepEqual(select({modo:'nup',total}),all(total));
        assert.deepEqual(select({modo:'nup',total,texto:'999'}),all(total),'campo inativo não substitui Todas');
        for(const texto of [String(total+1),`1-${total+1}`,'0','-1','abc','2-1','',' ','1,,1','1.5','1-0','1e2'])assert.throws(()=>select({modo:'nup',selecao:'especificas',texto,total}),e=>e.code==='MONTAGEM_SELECAO',texto);
        assert.deepEqual(select({modo:'nup',selecao:'especificas',texto:String(total),total}),[total]);
    });
    test(`PDF ${total}: células/compacto, 2/4/6/8-up, ordem, última folha e matrizes`,async()=>{
        const pdf=await PDFDocument.create();for(let i=1;i<=total;i++)pdf.addPage([105*72/25.4,148*72/25.4]).drawText('Pagina '+i,{x:20,y:50});
        const bytes=await pdf.save(),original=bytes.slice();
        for(const nupDistribuicao of ['celulas','compacto'])for(const capacity of [2,4,6,8]){
            const paginas=select({modo:'nup',total}),r=await gerarMontagem(bytes,{modo:'nup',nupDistribuicao,paginas,largura:297,altura:420,colunas:2,linhas:capacity/2,escalaModo:'ajustar',modoMarcas:'sem',margens:{esquerda:0,direita:0,superior:0,inferior:0}},{previsualizarExcesso:true});
            assert.equal(r.plan.totalFolhas,Math.ceil(total/capacity));assert.deepEqual(r.plan.folhas.flat().map(p=>p.origem),all(total));assert.equal(r.plan.folhas.at(-1).length,total%capacity||capacity);
            assert.equal((await PDFDocument.load(r.bytes)).getPageCount(),Math.ceil(total/capacity));assert.equal(await conferirTransformacoes(r.bytes,r.plan),total);assert.deepEqual(bytes,original);
        }
    });
}
test('seleção específica preserva ordem explícita; reabertura com menos páginas não remapeia',()=>{
    const config={modo:'nup',selecao:'especificas',texto:'5,1,3-4',total:5};assert.deepEqual(select(config),[5,1,3,4]);assert.throws(()=>select({...config,total:4}),/Página 5 inexistente/);
    assert.deepEqual(select({modo:'nup',total:8}),all(8));assert.deepEqual(select({modo:'nup',total:1}),[1]);
    assert.deepEqual(select({...config,texto:'1,1'}),[1,1],'repetição somente se explicitamente pedida');
});
test('repetição mantém arte única e seleção independente; quantidade de peças não vira seleção',()=>{
    assert.deepEqual(select({modo:'repeat',texto:'4',total:5}),[4]);assert.throws(()=>select({modo:'repeat',texto:'1-2',total:5}),/exatamente uma arte/);
    const s=new ConfiguracoesMontagem({paginas:'3',quantidade:'24',grade:'auto'}),nup=s.trocar('nup',{paginas:'3',quantidade:'24',grade:'auto'});assert.equal(nup.values.nupSelecao,'todas');
    const repeat=s.trocar('repeat',{...nup.values,paginas:'1-4',nupSelecao:'especificas'});assert.equal(repeat.values.paginas,'3');assert.equal(repeat.values.quantidade,'24');
    assert.equal(s.trocar('nup',repeat.values).values.nupSelecao,'especificas');
});
test('orçamento de seleção é validado antes de alocar ou expandir entradas enormes',()=>{
    for(const total of [0,-1,NaN,1.5])assert.throws(()=>select({modo:'nup',total}),/determinar/);
    assert.throws(()=>select({modo:'nup',total:10001}),/10.000/);
    assert.throws(()=>select({modo:'nup',selecao:'especificas',total:20000,texto:'1-20000'}),/10.000/);
    assert.throws(()=>select({modo:'nup',selecao:'especificas',total:10,texto:'99999999999999999999'}),e=>e.code==='MONTAGEM_SELECAO');
});
