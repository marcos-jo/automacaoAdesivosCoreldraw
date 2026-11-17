import test from 'node:test';
import assert from 'node:assert/strict';
import {PDFFormatos} from '../public/js/tools/pdf-formats.js';
import {orientarFolha} from '../public/js/tools/pdf-sheet-orientation.js';

const existentes=[{id:'custom-f4',nome:'F4',largura:328,altura:478,origem:'legado'},{id:'custom-f6',nome:'F6',largura:318,altura:318},{id:'custom-f8',nome:'F8',largura:238,altura:328}];
function fixture(dados={versao:1,formatos:existentes,extensao:{manter:true}}){
 let raw=JSON.stringify(dados),writes=0;
 return {get raw(){return raw;},get writes(){return writes;},externo(v){raw=v;},ler:()=>raw,gravar:v=>{raw=v;writes++;}};
}
test('B: v1/F4/F6/F8, IDs, nomes, medidas e campos desconhecidos preservados sem gravação na carga',async()=>{
 const storage=fixture(),original=storage.raw,c=new PDFFormatos(storage);await c.pronta;
 assert.equal(storage.raw,original);assert.equal(storage.writes,0);
 for(const f of c.listar().filter(f=>!f.padrao)){orientarFolha(f.largura,f.altura,'paisagem');orientarFolha(f.largura,f.altura,'retrato');}
 const novo=await c.salvar({nome:'Convite',largura:'140,5',altura:180});
 const doc=JSON.parse(storage.raw);assert.deepEqual(doc.formatos.slice(0,3),existentes);assert.deepEqual(doc.extensao,{manter:true});assert.equal(doc.versao,1);
 const reaberto=new PDFFormatos(storage);await reaberto.pronta;assert.deepEqual(reaberto.listar().find(f=>f.id===novo.id),{...novo,padrao:false});
 await reaberto.salvar({nome:'F4 novo nome',largura:328,altura:478},'custom-f4');assert.equal(JSON.parse(storage.raw).formatos[0].origem,'legado');
});
test('B: gravações assíncronas serializadas, duplicatas concorrentes recusadas e assinatura removível',async()=>{
 const base=fixture(),storage={ler:async()=>base.raw,gravar:async v=>{await new Promise(r=>setTimeout(r,5));base.gravar(v);}},c=new PDFFormatos(storage);let avisos=0;const remover=c.assinar(()=>avisos++);
 await Promise.all([c.salvar({nome:'A',largura:100,altura:200}),c.salvar({nome:'B',largura:101,altura:201})]);
 assert.equal(JSON.parse(base.raw).formatos.length,5);assert.equal(avisos,2);remover();
 const resultados=await Promise.allSettled([c.salvar({nome:'Único',largura:90,altura:50}),c.salvar({nome:' único ',largura:90,altura:50})]);
 assert.equal(resultados.filter(r=>r.status==='fulfilled').length,1);assert.equal(JSON.parse(base.raw).formatos.length,6);assert.equal(avisos,2);
});
test('B: nomes e medidas legados são preservados em disco e comparados canonicamente',async()=>{
 const legado={id:'custom-legado',nome:'  Papel   especial  ',largura:'123,45',altura:'234.56'};
 const s=fixture({versao:1,formatos:[legado]}),c=new PDFFormatos(s);await c.pronta;
 assert.equal(c.listar().find(f=>f.id===legado.id).largura,123.45);assert.equal(c.listar().find(f=>f.id===legado.id).nome,legado.nome);
 await assert.rejects(c.salvar({nome:'papel especial',largura:123.45,altura:234.56}),/existe/);
 await c.salvar({nome:'Outro',largura:90,altura:50});assert.deepEqual(JSON.parse(s.raw).formatos[0],legado);
});
test('B: catálogo inválido, versão desconhecida e falha de leitura não podem sobrescrever o original',async()=>{
 for(const raw of ['{',JSON.stringify({versao:2,formatos:existentes}),JSON.stringify({versao:1,formatos:[...existentes,existentes[0]]})]){
  const s=fixture();s.externo(raw);const c=new PDFFormatos(s);await c.pronta;await assert.rejects(c.salvar({nome:'X',largura:1,altura:2}),/bloqueada/);await assert.rejects(c.excluir('custom-f4'),/bloqueada/);assert.equal(s.raw,raw);assert.equal(s.writes,0);
 }
 let writes=0;const c=new PDFFormatos({ler(){throw Error('Indisponível');},gravar(){writes++;}});await assert.rejects(c.salvar({nome:'X',largura:1,altura:2}),/bloqueada/);assert.equal(writes,0);
});
test('B: falha de gravação mantém lista/armazenamento e conflito entre instâncias não perde alterações',async()=>{
 const s=fixture(),a=new PDFFormatos(s),b=new PDFFormatos(s);await Promise.all([a.pronta,b.pronta]);const raw=s.raw,save=s.gravar;
 s.gravar=()=>{throw Error('Falha de gravação');};await assert.rejects(a.salvar({nome:'X',largura:1,altura:2}),/Falha/);assert.equal(s.raw,raw);assert.equal(a.personalizados.length,3);
 s.gravar=save;await a.salvar({nome:'X',largura:1,altura:2});const depois=s.raw;await assert.rejects(b.salvar({nome:'Y',largura:3,altura:4}),/outra janela/);assert.equal(s.raw,depois);assert.equal(JSON.parse(s.raw).formatos.length,4);
});
