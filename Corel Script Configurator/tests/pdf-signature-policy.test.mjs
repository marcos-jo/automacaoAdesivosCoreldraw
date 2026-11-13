import test from 'node:test';
import assert from 'node:assert/strict';
import { PDFDocument, PDFName, PDFHexString, PDFRef } from '../public/vendor/pdf-lib/pdf-lib.js';
import { PDFWorkspace, assinaturasUrl, motorUrl, copiaImpressaoUrl, impressaoUrl } from './helpers/pdf-resize-modules.mjs';
import { readFile } from 'node:fs/promises';
import { PDFCombinacao } from './helpers/pdf-combine-modules.mjs';
import { diagnosticarMontagem, gerarMontagem, prepararCopiaMontagem } from './helpers/pdf-imposition-modules.mjs';
import { criarPDFAssinado } from './helpers/pdf-signature-fixture.mjs';
import { PDFDocumentoVisual } from '../public/js/tools/pdf-visual-document.js';
const { diagnosticarAssinaturas, exigirOperacaoSemAssinatura } = await import(assinaturasUrl);
const { redimensionarDocumento } = await import(motorUrl);
const { reconstruirCopiaImpressao } = await import(copiaImpressaoUrl);
const { prepararParaImpressao } = await import(impressaoUrl);
let editor=await readFile(new URL('../public/js/tools/pdf-editor.js',import.meta.url),'utf8');
editor=editor.replaceAll('"./pdf-xref-policy.js"',JSON.stringify(new URL('../public/js/tools/pdf-xref-policy.js',import.meta.url).href));
editor=editor.replaceAll(JSON.stringify('/vendor/pdf-lib/pdf-lib.js'),JSON.stringify(new URL('../public/vendor/pdf-lib/pdf-lib.js',import.meta.url).href)).replaceAll(JSON.stringify('./pdf-signatures.js'),JSON.stringify(assinaturasUrl));
const legado=await import('data:text/javascript;base64,'+Buffer.from(editor).toString('base64'));
const N=PDFName.of, codigo='PDF_ASSINATURA_PROTEGIDA';
const arquivo=bytes=>({name:'sintetico.pdf',size:bytes.length,arrayBuffer:async()=>bytes.slice().buffer});
async function caso(tipo){
 if(['aplicada','multipla','vazio','comum'].includes(tipo))return criarPDFAssinado({total:3,paginas:tipo==='multipla'?[1,3]:tipo==='comum'?[]:[1],vazia:tipo==='vazio'});
 const pdf=await PDFDocument.create();pdf.addPage([200,300]);pdf.addPage([200,300]);
 if(tipo==='inconclusivo')pdf.catalog.set(N('AcroForm'),pdf.context.obj({Fields:[],SigFlags:3}));
 else if(tipo==='referencia-ausente')pdf.catalog.set(N('Perms'),pdf.context.obj({DocMDP:PDFRef.of(999999)}));
 else {
  const value=pdf.context.obj({Type:'Sig',ByteRange:[0,10,20,10],Contents:PDFHexString.of('01020304')});
  if(tipo==='orfao')pdf.context.register(value);
  else pdf.catalog.set(N('Perms'),pdf.context.register(pdf.context.obj({DocMDP:pdf.context.register(value)})));
 }
 return pdf.save();
}
for(const tipo of ['aplicada','multipla','perms-indireto','inconclusivo','referencia-ausente','orfao'])test(`PDF-008 ${tipo}: todos os caminhos internos recusam antes de modificar/serializar`,async()=>{
 const bytes=await caso(tipo),original=bytes.slice(),ws=new PDFWorkspace();await ws.abrir(arquivo(bytes));
 const estado=ws.capturarEstado();let saves=0;ws.pdf.save=()=>{++saves;throw Error('save não deve iniciar');};
 assert.deepEqual(await ws.obterBytesVisualizacao(),bytes);
 const visual=Object.create(PDFDocumentoVisual.prototype);visual.workspace=ws;assert.deepEqual(await visual.getData(),bytes);
 const sync=[()=>ws.girarPagina(1,90),()=>ws.excluirPaginas([1]),()=>ws.aplicarRecorte({}),()=>ws.aplicarRedimensionamento({}),()=>ws.aplicarCopiaImpressao({})];
 for(const op of sync)assert.throws(op,{codigo});
 const asyncs=[()=>ws.reordenarPaginas([2,1]),()=>ws.moverPagina(1,2),()=>ws.duplicarPaginas([1]),()=>ws.obterBytes(),()=>ws.prepararRecorte(1,{x:10,y:10,width:100,height:100}),()=>ws.prepararRedimensionamento({}),()=>ws.prepararCopiaImpressao({}),()=>ws.adicionarArquivos([arquivo(bytes)])];
 for(const op of asyncs)await assert.rejects(op(),{codigo});
 assert.equal(saves,0);assert.deepEqual(ws.capturarEstado(),estado);assert.deepEqual(bytes,original);
 const clean=new PDFWorkspace();await clean.abrir(arquivo(await caso('comum')));const anterior=clean.capturarEstado();await assert.rejects(clean.adicionarArquivos([arquivo(bytes)]),{codigo});assert.deepEqual(clean.capturarEstado(),anterior);
 const draft=new PDFCombinacao();await assert.rejects(draft.add([arquivo(bytes)]),e=>e.codigo===codigo||e.cause?.codigo===codigo);assert.equal(draft.pages.length,0);
 // Simula chamador interno/worker que tenta pular add().
 draft.sources=[{pdf:await PDFDocument.load(bytes)}];await assert.rejects(draft.build(),{codigo});
 const pdf=await PDFDocument.load(bytes);assert.ok(diagnosticarMontagem(pdf).bloqueios.some(b=>b.codigo==='assinatura-protegida'));
 for(const modo of ['repeat','nup','livreto'])await assert.rejects(gerarMontagem(bytes,{modo}),e=>e.diagnostico?.bloqueios.some(b=>b.codigo==='assinatura-protegida'));
 await assert.rejects(prepararCopiaMontagem(bytes,{confirmado:true,validar:()=>{throw Error('validação não deve iniciar');}}),e=>!!e.diagnostico);
 await assert.rejects(redimensionarDocumento(pdf,{}),{codigo});
 await assert.rejects(reconstruirCopiaImpressao(pdf,{}),{codigo});
 assert.throws(()=>prepararParaImpressao(pdf,{bloqueios:[],removiveis:[]}),{codigo});
 const documento={getData:async()=>bytes.slice()};
 await assert.rejects(legado.salvarPdfGirado(documento,new Map(),'saida.pdf'),{codigo});
 await assert.rejects(legado.salvarPaginasSelecionadas(documento,[1],'extrair',new Map(),'saida.pdf'),{codigo});
});
test('PDF-008: campo vazio não é assinatura aplicada; PDF comum mantém edição/união/exportação',async()=>{
 const vazio=await PDFDocument.load(await caso('vazio'));assert.equal(diagnosticarAssinaturas(vazio).estado,'campos_vazios');assert.doesNotThrow(()=>exigirOperacaoSemAssinatura(vazio));
 const ws=new PDFWorkspace();await ws.abrir(arquivo(await caso('vazio')));ws.girarPagina(1,90);assert.equal(ws.pdf.getPage(0).getRotation().angle,90);
 await assert.rejects(ws.prepararRecorte(1,{x:10,y:10,width:100,height:100}),/AcroForm|assinatura/); // Proteção conservadora de formulário permanece.
 const common=new PDFWorkspace(),bytes=await caso('comum');await common.abrir(arquivo(bytes));common.girarPagina(1,90);await common.moverPagina(1,2);await common.adicionarArquivos([arquivo(bytes)]);assert.equal((await PDFDocument.load(await common.obterBytes())).getPageCount(),6);
 const draft=new PDFCombinacao();await draft.add([arquivo(bytes)]);assert.equal((await PDFDocument.load(await draft.build())).getPageCount(),3);
});

test('PDF-008: revalida a cópia preparada antes de adotá-la no workspace',async()=>{
 const ws=new PDFWorkspace();await ws.abrir(arquivo(await caso('comum')));
 const prep=await ws.prepararRecorte(1,{x:10,y:10,width:100,height:100}),estado=ws.capturarEstado();
 const protegido=await PDFDocument.load(await caso('perms-indireto'));
 prep.pdf=protegido;
 assert.throws(()=>ws.aplicarRecorte(prep),{codigo});assert.deepEqual(ws.capturarEstado(),estado);
});
