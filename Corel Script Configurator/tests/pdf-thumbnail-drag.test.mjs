import test from 'node:test';
import assert from 'node:assert/strict';
import {Elemento} from './helpers/pdf-dom.mjs';
import {PDFArrasteMiniaturas} from '../public/js/tools/pdf-thumbnail-drag.js';

function fixture(total){
 const frames=new Map();let serial=0;
 globalThis.document={createElement:tag=>new Elemento(tag)};
 globalThis.requestAnimationFrame=fn=>{frames.set(++serial,fn);return serial;};
 globalThis.cancelAnimationFrame=id=>frames.delete(id);
 const container=new Elemento('div'),list=new Elemento('div'),source=new Elemento('button');
 container.clientHeight=600;container.scrollHeight=total*166;container.scrollTop=0;
 container.getBoundingClientRect=()=>({top:100,bottom:700,left:0,right:200});
 source.dataset.pagina='1';list.append(source);container.append(list);
 const moves=[],drag=new PDFArrasteMiniaturas(container,{alturaItem:166,obterLista:()=>list,obterTotal:()=>total,bloqueado:()=>false,aoMover:(...args)=>moves.push(args),aoRolar:()=>{}});
 const transfer={setData(){}};
 const over=y=>container.emitir('dragover',container,{dataTransfer:transfer,clientY:y,clientX:100});
 const tick=()=>{const first=frames.entries().next().value;if(!first)return;frames.delete(first[0]);first[1]();};
 const flush=()=>{let budget=3000;while(frames.size&&budget-->0)tick();assert.ok(budget>0,'animação termina no limite');};
 source.emitir('dragstart',source,{dataTransfer:transfer});
 return {container,source,drag,frames,moves,over,tick,flush,drop:y=>container.emitir('drop',container,{dataTransfer:transfer,clientY:y}),cleanup(){drag.destruir();delete globalThis.document;delete globalThis.requestAnimationFrame;delete globalThis.cancelAnimationFrame;}};
}
for(const total of [8,100,350])test(`autoscroll ${total}: extremos, indicador e destino correto sem ultrapassar a área rolável`,()=>{
 const f=fixture(total);
 try{
  f.over(696);f.flush();assert.equal(f.container.scrollTop,total*166-600);assert.equal(f.drag.posicao,total);assert.equal(f.drag.marcador.style.top,`${total*166-4}px`);assert.equal(f.frames.size,0);
  f.over(104);f.flush();assert.equal(f.container.scrollTop,0);assert.equal(f.drag.posicao,0);assert.equal(f.frames.size,0);
  f.over(696);f.flush();f.drop(696);assert.deepEqual(f.moves,[[1,total]]);assert.equal(f.drag.marcador,null);assert.equal(f.frames.size,0);
 }finally{f.cleanup();}
});
test('sair da faixa, sair do painel e cancelar não deixam quadros pendentes nem aplicam movimento',()=>{
 const f=fixture(350);
 try{
  f.over(696);f.tick();assert.ok(f.frames.size);f.over(400);f.tick();assert.equal(f.frames.size,0);
  f.over(696);f.tick();f.container.emitir('dragleave',f.container,{clientX:300,clientY:750,relatedTarget:null});assert.equal(f.frames.size,0);assert.equal(f.drag.marcador.hidden,true);
  f.over(104);f.source.emitir('dragend');assert.equal(f.frames.size,0);assert.equal(f.drag.origem,null);assert.equal(f.drag.marcador,null);assert.deepEqual(f.moves,[]);
  f.drag.destruir();for(const name of ['dragstart','dragover','drop','dragend','dragleave'])assert.equal(f.container.listeners.get(name).length,0);
 }finally{f.cleanup();}
});
