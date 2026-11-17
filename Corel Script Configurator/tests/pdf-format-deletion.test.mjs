import test from 'node:test';
import assert from 'node:assert/strict';
import {PDFFormatos} from '../public/js/tools/pdf-formats.js';

function storage(raw = null) {
    let writes = 0;
    return {ler: () => raw, gravar: value => {raw = value; writes++;}, get raw(){return raw;}, get writes(){return writes;}};
}
test('A: exclusões de todos os tamanhos iniciais persistem; catálogo vazio não é reinicializado', async () => {
    const s=storage(), c=new PDFFormatos(s); await c.pronta;
    assert.equal(s.writes,0);
    const ids=c.listar().map(f=>f.id);assert.equal(ids.length,9);assert.ok(!ids.includes('SRA3'));
    for(const id of ids) {
        await c.excluir(id); const novo=new PDFFormatos(s); await novo.pronta;
        assert.ok(!novo.listar().some(f=>f.id===id));assert.equal(s.writes,ids.indexOf(id)+1);
    }
    assert.deepEqual(new PDFFormatos(s).listar(),[]);
});
test('A: legado SRA3 não reaparece; metadados, IDs e medidas dos demais ficam intactos', async () => {
    const existentes=[{id:'custom-sra3',nome:'SRA3',largura:320,altura:450,origem:{manter:true}},
        ...[['F4',328,478],['F6',318,318],['F8',238,328]].map(([nome,largura,altura])=>({id:'custom-'+nome,nome,largura,altura,extra:'original'}))];
    const s=storage(JSON.stringify({versao:1,formatos:existentes,extensao:'preservar'})),c=new PDFFormatos(s);await c.pronta;
    assert.equal(c.listar().filter(f=>f.nome==='SRA3').length,1);assert.equal(s.writes,0);
    await c.excluir('custom-sra3');const raw=s.raw;
    const novo=new PDFFormatos(s);await novo.pronta;
    assert.ok(!novo.listar().some(f=>f.nome==='SRA3'));assert.equal(s.raw,raw);
    assert.deepEqual(JSON.parse(raw).formatos.slice(0,3),existentes.slice(1));assert.equal(JSON.parse(raw).extensao,'preservar');
    assert.equal(novo.listar().length,3);
    assert.ok(!novo.listar().some(f=>f.nome==='A3'));
});
test('A: ausência legada de SRA3 e nomes iniciais personalizados não criam duplicatas', async () => {
    const s=storage(JSON.stringify({versao:1,formatos:[{id:'custom-a3',nome:'A3',largura:300,altura:400}]})),c=new PDFFormatos(s);await c.pronta;
    assert.equal(c.listar().filter(f=>f.nome==='A3').length,1);assert.ok(!c.listar().some(f=>f.nome==='SRA3'));
    await c.excluir('custom-a3');assert.ok(!new PDFFormatos(s).listar().some(f=>f.nome==='A3'));
});
test('A: falha de persistência não publica exclusão nem converte o documento legado', async () => {
    const s=storage(),c=new PDFFormatos(s);await c.pronta;const raw=s.raw,before=c.listar();let notify=0;c.assinar(()=>notify++);
    s.gravar=()=>{throw Error('disco indisponível');};await assert.rejects(c.excluir('A4'),/indisponível/);
    assert.equal(s.raw,raw);assert.deepEqual(c.listar(),before);assert.equal(notify,0);
});
test('B: atualização enfileirada não permite sobrescrever edição ou exclusão com uma seleção obsoleta', async () => {
    const s=storage(),c=new PDFFormatos(s);await c.pronta;const antigo=c.listar().find(f=>f.id==='A4');
    const primeiro=c.salvar({nome:'A4 novo',largura:211,altura:298},'A4');
    const segundo=c.salvar({nome:'A4 antigo',largura:210,altura:297},'A4',antigo);
    const terceiro=c.excluir('A4',antigo);
    await primeiro;await assert.rejects(segundo,/outra ferramenta/);await assert.rejects(terceiro,/outra ferramenta/);
    assert.equal(c.listar().find(f=>f.id==='A4').nome,'A4 novo');assert.equal(s.writes,1);
});
