import test from 'node:test';
import assert from 'node:assert/strict';
import {createHash} from 'node:crypto';
import {PDFFormatos,FORMATOS_INICIAIS} from '../public/js/tools/pdf-formats.js';
import {PDFArmazenamentoFormatos} from '../public/js/tools/pdf-format-storage.js';
import {lerBackupFormatos,criarBackupFormatos} from '../public/js/tools/pdf-format-backup.js';

function fixture(raw=null){const values=new Map([['4printshops.pdf.formatos.v1',raw],['preferencias','intactas']]);return {values,ler:()=>values.get('4printshops.pdf.formatos.v1'),gravar:v=>values.set('4printshops.pdf.formatos.v1',v)};}
test('A: primeiro uso oferece exatamente os nove tamanhos; editar/excluir/recriar não duplica iniciais',async()=>{
    const s=fixture(),c=new PDFFormatos(s);await c.pronta;
    assert.deepEqual(c.listar().map(({nome,largura,altura})=>[nome,largura,altura]),[['A3',297,420],['A4',210,297],['A5',148,210],['A6',105,148],['A7',74,105],['F4',328,478],['F8',238,328],['F6',318,318],['Faixa',320,700]]);
    assert.equal(s.ler(),null);assert.ok(c.listar().every(f=>!f.padrao));
    await c.salvar({nome:'A3 editado',largura:300,altura:430},'A3');await c.excluir('A3');
    const novo=await c.salvar({nome:'A3',largura:297,altura:420});
    assert.notEqual(novo.id,'A3');assert.equal(new PDFFormatos(s).listar().filter(f=>f.nome==='A3').length,1);
});
test('A: qualquer documento existente é catálogo inicializado, inclusive legado vazio sem marcador',async()=>{
    for(const formatos of [[],[{id:'custom-legado',nome:'Outro',largura:99,altura:199}]]){
        const raw=JSON.stringify({versao:1,formatos}),s=fixture(raw),c=new PDFFormatos(s);await c.pronta;
        assert.equal(c.listar().length,formatos.length);assert.equal(s.ler(),raw);
    }
});
test('A: backup verificável, redefinição e restauração exata preservam IDs/metadata e somente a chave do catálogo',async()=>{
    const original=JSON.stringify({versao:1,extensao:{origem:'sintetico'},formatos:[{id:'custom-antigo',nome:' F4 teste ',largura:'328,0',altura:'478',extra:{cor:'azul'}}]},null,2);
    const s=fixture(original),c=new PDFFormatos(s);await c.pronta;const backup=await c.exportarBackup();
    assert.equal(s.ler(),original);assert.equal(backup.sha256,createHash('sha256').update(backup.conteudo).digest('hex'));
    assert.equal((await lerBackupFormatos(JSON.stringify(backup))).original,original);
    await c.redefinir(JSON.stringify(backup));assert.equal(c.listar().length,9);assert.equal(s.values.get('preferencias'),'intactas');
    await c.restaurarBackup(JSON.stringify(backup));assert.equal(s.ler(),original);assert.equal(c.listar()[0].id,'custom-antigo');assert.deepEqual(c.listar()[0].extra,{cor:'azul'});assert.equal(s.values.get('preferencias'),'intactas');
});
test('A: backup divergente, obsoleto e catálogo inválido recusam alteração',async()=>{
    const s=fixture(),c=new PDFFormatos(s);await c.pronta;const backup=await c.exportarBackup();
    const tampered={...backup,conteudo:backup.conteudo+' '};await assert.rejects(c.redefinir(tampered),/SHA-256/);assert.equal(s.ler(),null);
    await c.excluir('A4');const raw=s.ler();await assert.rejects(c.redefinir(backup),/não corresponde/);assert.equal(s.ler(),raw);
    const invalid=await criarBackupFormatos('{"versao":2,"formatos":[]}',{versao:1,formatos:[]});await assert.rejects(c.restaurarBackup(invalid),/inválido/);assert.equal(s.ler(),raw);
    const broken={...backup,sha256:'0'.repeat(64)};await assert.rejects(c.restaurarBackup(broken),/SHA-256/);assert.equal(s.ler(),raw);
});
test('A: restauração de primeiro uso e catálogo vazio não reinicializa sugestões por acidente',async()=>{
    const s=fixture(),c=new PDFFormatos(s);await c.pronta;const first=await c.exportarBackup();
    for(const f of c.listar())await c.excluir(f.id);const empty=await c.exportarBackup();
    await c.redefinir(empty);assert.equal(c.listar().length,9);await c.restaurarBackup(empty);assert.deepEqual(new PDFFormatos(s).listar(),[]);
    await c.restaurarBackup(first);assert.deepEqual(c.listar().map(f=>f.id),FORMATOS_INICIAIS.map(f=>f.id));
});
test('A: redefinição/recuperação não sobrescrevem mudança externa nem publicam falha de escrita',async()=>{
    const s=fixture(),a=new PDFFormatos(s),b=new PDFFormatos(s);await a.pronta;const backup=await a.exportarBackup();
    await b.salvar({nome:'Concorrente',largura:101,altura:201});const raw=s.ler();await assert.rejects(a.redefinir(backup),/outra janela/);assert.equal(s.ler(),raw);
    const c=new PDFFormatos(s);await c.pronta;const atual=await c.exportarBackup(),before=c.listar();s.gravar=()=>{throw Error('disco indisponivel');};
    await assert.rejects(c.redefinir(atual),/indisponivel/);await assert.rejects(c.restaurarBackup(backup),/indisponivel/);assert.deepEqual(c.listar(),before);assert.equal(s.ler(),raw);
});
test('A: redefinição não é aprovada por fallback somente em memória ou leitura indisponível',async()=>{
    let raw='{"versao":1,"formatos":[]}',failRead=false,failWrite=false;
    const local={getItem:()=>{if(failRead)throw Error('leitura');return raw;},setItem:(_k,v)=>{if(failWrite)throw Error('quota');raw=v;}};
    const c=new PDFFormatos(new PDFArmazenamentoFormatos(()=>local,'reset-sintetico'));await c.pronta;const backup=await c.exportarBackup();
    failRead=true;await assert.rejects(c.redefinir(backup),/indisponível/);assert.equal(raw,'{"versao":1,"formatos":[]}');
    failRead=false;const d=new PDFFormatos(new PDFArmazenamentoFormatos(()=>local,'reset-quota-sintetico'));await d.pronta;const outro=await d.exportarBackup();
    failWrite=true;await assert.rejects(d.redefinir(outro),/não confirmada/);assert.equal(raw,'{"versao":1,"formatos":[]}');assert.deepEqual(d.listar(),[]);
});
