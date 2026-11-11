import test from 'node:test';
import assert from 'node:assert/strict';
import {Elemento} from './helpers/pdf-dom.mjs';
import {PDFCombinacao} from './helpers/pdf-combine-modules.mjs';
import {PDFDocument} from '../public/vendor/pdf-lib/pdf-lib.js';
import {PreparacaoConferencia} from '../public/js/tools/pdf-import-review.js';
import {instalarImportacaoDesktop} from '../public/js/app/desktop-client.js';
import {executarImportacao} from '../public/js/tools/pdf-import-executor.js';
import {pdfProtegido} from './helpers/pdf-import-protection-fixtures.mjs';
const deferred=()=>{let resolve,reject;const promise=new Promise((r,j)=>{resolve=r;reject=j;});return {promise,resolve,reject};};
const tick=()=>new Promise(r=>setImmediate(r));
async function pdf(pages=1){const p=await PDFDocument.create();for(let i=0;i<pages;i++)p.addPage([300,400]);return new File([await p.save()],'sintetico.pdf',{type:'application/pdf'});}
function setup(options={}){
 globalThis.document={createElement:tag=>new Elemento(tag)};
 const list=new Elemento('ol'),count=new Elemento('p'),previous=new PDFCombinacao(),previews=[];
 class Preview{
  constructor(host){this.host=host;this.gate=options.gate;previews.push(this);}
  async preparar(bytes){if(this.gate)await this.gate;return {bytes};}
  async adotar(){return true;}
  async mostrar(){const c=new Elemento('canvas');c.width=100;this.host.replaceChildren(c);return true;}
  async destruir(){this.destroyed=true;}
 }
 const review=new PreparacaoConferencia({list,count,PDFCombinacao,PDFPreview:Preview,wordReports:new WeakMap(),prazo:options.prazo});
 return {review,list,count,previous,previews};
}
test('conferência antecipada: placeholders, contagem real e miniaturas antes da conclusão atômica',async()=>{
 const f=setup(),file=await pdf(100),b=f.review.begin([{name:'a.pdf',type:'PDF'},{name:'b.docx',type:'DOCX'}],f.previous,()=>{});
 assert.equal(f.list.children.length,2);assert.match(f.count.textContent,/0 de 2/);f.review.stage(1,'word');assert.equal(b.items[1].state.textContent,'Convertendo Word');
 await assert.rejects(f.review.complete(),/não está pronta/);await f.review.accept(0,file);assert.equal(b.items[0].thumb.children[0].tag,'img');assert.match(f.count.textContent,/1 de 2/);
 assert.equal(f.previous.pages.length,0);await assert.rejects(f.review.complete(),/não está pronta/);await f.review.accept(1,await pdf());
 const result=await f.review.complete();assert.equal(result.pages.length,101);assert.equal(f.previous.pages.length,0);assert.equal(f.review.batch,null);assert.ok(f.previews.every(p=>p.destroyed));
});
test('falha no meio conserva origens válidas e exige remover o item antes da conferência',async()=>{
 const f=setup();await f.previous.add([await pdf()]);const original=[...f.previous.pages];
 const b=f.review.begin([{name:'ok.pdf',type:'PDF'},{name:'falha.pdf',type:'PDF'},{name:'final.pdf',type:'PDF'}],f.previous,()=>{});
 await f.review.accept(0,await pdf());assert.equal(await f.review.accept(1,new File(['not PDF'],'falha.pdf')),false);
 assert.match(b.items[1].state.textContent,/PDF inválido/);assert.equal(b.items[1].remove.hidden,true);
 await f.review.accept(2,await pdf(2));assert.equal(await f.review.complete(),null);assert.deepEqual(f.previous.pages,original);
 b.items[1].remove.click();const result=await f.review.complete();assert.equal(result.pages.length,4);assert.deepEqual(f.previous.pages,original);
});
test('cancelamento durante miniatura não publica resultado tardio e reabertura é independente',async()=>{
 const gate=deferred(),f=setup({gate:gate.promise}),file=await pdf();f.review.begin([{name:file.name,type:'PDF'}],f.previous,()=>{});
 const old=f.review.accept(0,file);await tick();f.review.cancel();const next=f.review.begin([{name:'novo.pdf',type:'PDF'}],f.previous,()=>{});gate.resolve();await old;
 assert.equal(f.review.batch,next);assert.equal(next.items[0].ready,false);assert.equal(next.items[0].thumb.children.length,0);await f.review.accept(0,file);assert.equal((await f.review.complete()).pages.length,1);
});
test('prazo excedido preserva o lote para análise e não deixa miniatura tardia entrar',async()=>{
 const gate=deferred(),f=setup({gate:gate.promise,prazo:15}),file=await pdf();
 const b=f.review.begin([{name:file.name,type:'PDF'}],f.previous,()=>{});
 assert.equal(await f.review.accept(0,file),false);assert.match(b.items[0].state.textContent,/excedeu o prazo/);
 assert.equal(await f.review.complete(),null);gate.resolve();await tick();
 b.items[0].remove.click();assert.equal(await f.review.complete(),null);
 assert.equal(b.draft.pages.length,0);assert.equal(b.items[0].thumb.children.length,0);
 f.review.cancel();assert.ok(f.previews.every(p=>p.destroyed));
});
test('cancelar o executor termina o worker e rejeita a tarefa em andamento',async()=>{
 const original=globalThis.Worker;let worker;
 globalThis.Worker=class {constructor(){worker=this;}postMessage(){}terminate(){this.encerrado=true;}};
 try{
  const controller=new AbortController();const result=executarImportacao('preparar',{}, {signal:controller.signal});
  controller.abort();await assert.rejects(result,{name:'AbortError'});assert.equal(worker.encerrado,true);
 }finally{globalThis.Worker=original;}
});
test('a versão empacotada do pdf-lib sinaliza criptografia para verificação no PDF.js',async()=>{
 const draft=new PDFCombinacao();
 for(const password of ['', 'sintetica']){
  await assert.rejects(draft.add([new File([pdfProtegido(password)],'protegido.pdf')]),{codigo:'PDF_CRIPTOGRAFADO'});
 }
 assert.equal(draft.pages.length,0);
});
test('cancelar antes da leitura terminar preserva os arquivos e o documento anterior',async()=>{
 const f=setup(),file=await pdf(),gate=deferred();let cancelled=0;const b=f.review.begin([{name:file.name,type:'PDF'}],f.previous,()=>{cancelled++;});
 const delayed={name:file.name,size:file.size,arrayBuffer:async()=>{await gate.promise;return file.arrayBuffer();}};
 const old=f.review.accept(0,delayed);f.review.cancel();gate.resolve();await old;assert.equal(cancelled,1);assert.equal(b.items[0].ready,false);assert.equal(f.previous.pages.length,0);
});
function desktopFixture({releaseOnFailure=false}={}){
 globalThis.document={createElement:tag=>new Elemento(tag)};const root=new Elemento('div'),status=new Elemento('p'),input=new Elemento('input');status.id='pdf-status';Object.assign(input,{tagName:'INPUT',type:'file',accept:'.pdf,.docx,.doc',multiple:true});root.append(status,input);
 let observer,request,cancelled=0,finished=0;const response=deferred(),accepted=deferred(),calls=[];
 globalThis.window={desktop4PrintShops:{observarImportacao:fn=>{observer=fn;return()=>{observer=null;};},importarArquivos:async(_multiple,id)=>{request=id;return response.promise;},cancelarImportacao:async()=>{cancelled++;}}};
 const stop=instalarImportacaoDesktop(root,{iniciar:()=>{input.disabled=true;return true;},finalizar:()=>{input.disabled=false;finished++;},conferencia:{iniciar:(files,cancel)=>{calls.push(['selection',files]);globalThis.__cancelReview=cancel;},etapa:(...args)=>calls.push(['stage',...args]),arquivo:async(...args)=>{calls.push(['entry',args[0]]);await accepted.promise;},concluir:()=>calls.push(['complete']),cancelar:()=>calls.push(['cancel']),falhar:()=>{calls.push(['failure']);if(releaseOnFailure)accepted.resolve();}}});
 input.click();const publish=value=>observer?.({jobId:request,...value});
 return{input,root,calls,response,accepted,publish,stop,get finished(){return finished;},get cancelled(){return cancelled;},get observing(){return !!observer;}};
}
test('ponte espera validação/miniatura antes de liberar; eventos de outro lote são ignorados',async()=>{
 const f=desktopFixture();f.publish({jobId:'outro-lote',kind:'selection',files:[]});assert.equal(f.calls.length,0);
 f.publish({kind:'selection',files:[{index:0,name:'a.pdf',type:'PDF'}]});f.publish({kind:'entry',index:0,entry:{name:'a.pdf',bytes:new Uint8Array([1])}});f.response.resolve({progressive:true,count:1});await tick();assert.equal(f.input.disabled,true);assert.equal(f.finished,0);assert.ok(!f.calls.some(c=>c[0]==='complete'));
 f.accepted.resolve();await tick();assert.equal(f.finished,1);assert.equal(f.input.disabled,false);assert.ok(f.calls.some(c=>c[0]==='complete'));f.stop();assert.equal(f.observing,false);
});
test('ponte cancelada ignora evento/resultado atrasado e remove observador ao encerrar',async()=>{
 const f=desktopFixture();f.publish({kind:'selection',files:[{index:0,name:'a.docx',type:'DOCX'}]});globalThis.__cancelReview();f.publish({kind:'entry',index:0,entry:{name:'a.pdf',bytes:new Uint8Array([1])}});f.response.resolve({progressive:true,count:1});await tick();assert.equal(f.cancelled,1);assert.ok(!f.calls.some(c=>c[0]==='entry'||c[0]==='complete'));assert.equal(f.finished,1);f.stop();assert.equal(f.observing,false);
});
test('resposta IPC anterior aos eventos não conclui nem libera lote sem todos os arquivos',async()=>{
 const f=desktopFixture();f.response.resolve({progressive:true,count:1});await tick();assert.equal(f.finished,0);
 f.publish({kind:'selection',files:[{index:0,name:'atrasado.pdf',type:'PDF'}]});await tick();assert.equal(f.finished,0);
 f.publish({kind:'entry',index:0,entry:{name:'atrasado.pdf',bytes:new Uint8Array([1])}});await tick();assert.equal(f.finished,0);
 f.accepted.resolve();await tick();assert.equal(f.finished,1);assert.ok(f.calls.some(c=>c[0]==='complete'));f.stop();
});
test('falha do backend aborta a prévia pendente antes de aguardar a fila, sem adotar parte do lote',async()=>{
 const f=desktopFixture({releaseOnFailure:true});f.publish({kind:'selection',files:[{index:0,name:'ok.pdf',type:'PDF'},{index:1,name:'erro.doc',type:'DOC'}]});f.publish({kind:'entry',index:0,entry:{name:'ok.pdf',bytes:new Uint8Array([1])}});await tick();
 f.publish({kind:'stage',index:1,state:'failed'});f.response.reject(Error('Falha sintética'));await tick();assert.ok(f.calls.some(c=>c[0]==='failure'));assert.ok(!f.calls.some(c=>c[0]==='complete'));assert.equal(f.finished,1);f.stop();
});
