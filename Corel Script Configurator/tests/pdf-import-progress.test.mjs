import test from 'node:test';
import assert from 'node:assert/strict';
import {Elemento} from './helpers/pdf-dom.mjs';
import {PDFCombinacao} from './helpers/pdf-combine-modules.mjs';
import {PDFDocument} from '../public/vendor/pdf-lib/pdf-lib.js';
import {PreparacaoConferencia} from '../public/js/tools/pdf-import-review.js';
import {instalarImportacaoDesktop} from '../public/js/app/desktop-client.js';
const deferred=()=>{let resolve,reject;const promise=new Promise((r,j)=>{resolve=r;reject=j;});return {promise,resolve,reject};};
const tick=()=>new Promise(r=>setImmediate(r));
async function pdf(pages=1){const p=await PDFDocument.create();for(let i=0;i<pages;i++)p.addPage([300,400]);return new File([await p.save()],'sintetico.pdf',{type:'application/pdf'});}
function setup(){
 globalThis.document={createElement:tag=>new Elemento(tag)};
 const list=new Elemento('ol'),count=new Elemento('p'),host=new Elemento('div'),previous=new PDFCombinacao();
 class Preview{async preparar(bytes){if(this.gate)await this.gate;return {bytes};}async adotar(){return true;}async mostrar(){const c=new Elemento('canvas');c.width=100;host.replaceChildren(c);return true;}invalidar(){}async destruir(){this.destroyed=true;}async descartar(){}}
 const review=new PreparacaoConferencia({list,count,host,PDFCombinacao,PDFPreview:Preview,wordReports:new WeakMap()});
 return {review,list,count,previous};
}
test('conferência antecipada: placeholders, contagem real e miniaturas antes da conclusão atômica',async()=>{
 const f=setup(),file=await pdf(100),b=f.review.begin([{name:'a.pdf',type:'PDF'},{name:'b.docx',type:'DOCX'}],f.previous,()=>{});
 assert.equal(f.list.children.length,2);assert.match(f.count.textContent,/0 de 2/);f.review.stage(1,'word');assert.equal(b.items[1].state.textContent,'Convertendo Word');
 await assert.rejects(f.review.complete(),/não está pronta/);await f.review.accept(0,file);assert.equal(b.items[0].thumb.children[0].tag,'img');assert.match(f.count.textContent,/1 de 2/);
 assert.equal(f.previous.pages.length,0);await assert.rejects(f.review.complete(),/não está pronta/);await f.review.accept(1,await pdf());
 const result=await f.review.complete();assert.equal(result.pages.length,101);assert.equal(f.previous.pages.length,0);assert.equal(f.review.batch,null);assert.equal(b.preview.destroyed,true);
});
test('falha no meio conserva rascunho anterior e nunca libera seleção parcial',async()=>{
 const f=setup();await f.previous.add([await pdf()]);const original=[...f.previous.pages];
 const b=f.review.begin([{name:'ok.pdf',type:'PDF'},{name:'falha.pdf',type:'PDF'}],f.previous,()=>{});
 await f.review.accept(0,await pdf());await assert.rejects(f.review.accept(1,new File(['not PDF'],'falha.pdf')));f.review.fail(1);
 assert.equal(b.items[1].state.textContent,'Falha');await assert.rejects(f.review.complete());assert.deepEqual(f.previous.pages,original);f.review.cancel();assert.deepEqual(f.previous.pages,original);
});
test('cancelamento durante miniatura não publica resultado tardio e reabertura é independente',async()=>{
 const f=setup(),file=await pdf(),b=f.review.begin([{name:file.name,type:'PDF'}],f.previous,()=>{}),gate=deferred();b.preview.gate=gate.promise;
 const old=f.review.accept(0,file);await tick();f.review.cancel();const next=f.review.begin([{name:'novo.pdf',type:'PDF'}],f.previous,()=>{});gate.resolve();await old;
 assert.equal(f.review.batch,next);assert.equal(next.items[0].ready,false);assert.equal(next.items[0].thumb.children.length,0);await f.review.accept(0,file);assert.equal((await f.review.complete()).pages.length,1);
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
