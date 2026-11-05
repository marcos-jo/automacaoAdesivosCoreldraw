import test from 'node:test';
import assert from 'node:assert/strict';
import {mkdtemp,readdir,readFile,writeFile,rm} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import path from 'node:path';
import {ImageExportStore,validateImagePlan,validateImageBytes} from '../desktop/image-exports.mjs';
import {imageSize,imagePages,imageNames,imageDensity,imageDestination} from '../public/js/tools/pdf-image-export.js';
import {instalarAtalhos} from '../public/js/app/shortcuts.js';
import {png,jpg} from './helpers/pdf-combine-fixtures.mjs';

test('DPI gera pixels efetivos em A4/A3, orientação e limites de memória',()=>{
 for(const dpi of [150,300,600]){const s=imageSize(210*72/25.4,297*72/25.4,dpi);assert.equal(s.width,Math.ceil(210*dpi/25.4-1e-7));assert.equal(s.height,Math.ceil(297*dpi/25.4-1e-7));assert.equal(s.estimatedMemory,s.pixels*16);}
 assert.equal(imageSize(297*72/25.4,420*72/25.4,300).width,3508);assert.throws(()=>imageSize(842,1191,600),/36/);assert.throws(()=>imageSize(Infinity,10,300));assert.throws(()=>imageSize(1,1,72));
});
test('páginas atual/todas/seleção/intervalo e nomes ordenados não aceitam caminhos',()=>{
 assert.deepEqual(imagePages('all','',3,2),[1,2,3]);assert.deepEqual(imagePages('current','',3,2),[2]);assert.deepEqual(imagePages('selected','',3,1,[3,1]),[1,3]);assert.deepEqual(imagePages('range','3,1-2,2',3,1),[1,2,3]);
 for(const range of ['0','4','2-1','1-x',''])assert.throws(()=>imagePages('range',range,3,1));assert.throws(()=>imagePages('selected','',3,1,[]));
 assert.deepEqual(imageNames('São José',[1,12,350],'png'),['São José_p0001.png','São José_p0012.png','São José_p0350.png']);for(const name of ['../escape','CON','a/b','nome.'])assert.throws(()=>imageNames(name,[1],'png'));
});
for(const format of ['png','jpeg'])test('gravação '+format+' publica somente lote completo e conserva arquivos existentes',async()=>{
 const root=await mkdtemp(path.join(tmpdir(),'4ps-images-'));const data=format==='png'?png():jpg();
 try{await writeFile(path.join(root,'original.png'),'preservado');const store=new ImageExportStore({}, {showOpenDialog:async()=>({filePaths:[root]})});const plan={base:'Balcão',format,names:imageNames('Balcão',[1,2],format)};const job=await store.request({action:'begin',plan});
 await store.request({action:'write',token:job.token,index:0},data);assert.deepEqual(await readdir(job.location),[plan.names[0]+'.part']);await store.request({action:'write',token:job.token,index:1},data);const result=await store.request({action:'finish',token:job.token});assert.equal(result.count,2);assert.deepEqual(await readFile(path.join(job.location,plan.names[0])),Buffer.from(data));assert.equal(await readFile(path.join(root,'original.png'),'utf8'),'preservado');assert.equal(store.active,false);
 }finally{await rm(root,{recursive:true,force:true});}
});
test('cancelamento e falhas removem exclusivamente arquivos criados pelo lote',async()=>{
 const root=await mkdtemp(path.join(tmpdir(),'4ps-images-'));
 try{await writeFile(path.join(root,'cliente.png'),'intacto');const store=new ImageExportStore({}, {showOpenDialog:async()=>({filePaths:[root]})}),plan={base:'teste',format:'png',names:['teste_p0001.png','teste_p0002.png']};
 for(const action of ['cancel','bad']){const job=await store.request({action:'begin',plan});await store.request({action:'write',token:job.token,index:0},png());if(action==='cancel')await store.request({action:'cancel',token:job.token});else await assert.rejects(store.request({action:'write',token:job.token,index:1},new Uint8Array([1])));assert.deepEqual(await readdir(root),['cliente.png']);assert.equal(store.active,false);}
 const canceled=new ImageExportStore({}, {showOpenDialog:async()=>({canceled:true})});assert.deepEqual(await canceled.request({action:'begin',plan}),{canceled:true});
 }finally{await rm(root,{recursive:true,force:true});}
});
test('plano, token, ordem e extensão são validados no processo Desktop',async()=>{
 for(const plan of [{base:'../a',format:'png',names:['a.png']},{base:'a',format:'png',names:['a.png','a.png']},{base:'a',format:'exe',names:['a.exe']},{base:'a',format:'png',names:['../a.png']}])assert.throws(()=>validateImagePlan(plan));assert.throws(()=>validateImageBytes(jpg(),'png'));
 const root=await mkdtemp(path.join(tmpdir(),'4ps-images-'));try{const store=new ImageExportStore({}, {showOpenDialog:async()=>({filePaths:[root]})}),plan={base:'a',format:'png',names:['a.png']};const job=await store.request({action:'begin',plan});await assert.rejects(store.request({action:'write',token:'falso',index:0},png()));assert.equal(store.active,true);await assert.rejects(store.request({action:'write',token:job.token,index:1},png()));assert.equal(store.active,false);}finally{await rm(root,{recursive:true,force:true});}
});
test('atalhos únicos bloqueiam padrão Chromium, repetição, campos, modais e ações indisponíveis',async()=>{
 let handler;const calls=[],target={addEventListener:(_type,h)=>{handler=h;},removeEventListener:(_type,h)=>assert.equal(h,handler)};globalThis.document={querySelector:()=>null};let available=true;
 try{const remove=instalarAtalhos({execute:action=>calls.push(action),available:()=>available},target);const key=(key,extra={})=>{const event={key,ctrlKey:true,preventDefault(){this.prevented=true;},target:{closest:()=>false},...extra};handler(event);return event;};
 for(const keyName of ['o','s','p','e'])assert.equal(key(keyName).prevented,true);assert.deepEqual(calls,['open','save','print','export']);key('p',{repeat:true});key('s',{target:{closest:()=>true}});document.querySelector=()=>({});key('p');document.querySelector=()=>null;available=false;key('e');assert.equal(calls.length,4);remove();
 }finally{delete globalThis.document;}
});

test('densidade PNG/JPEG é registrada sem recodificar a imagem',()=>{
 const original=png(),output=imageDensity(original,'png',300);assert.equal(output.length,original.length+21);assert.equal(new DataView(output.buffer).getUint32(41),Math.round(300/0.0254));assert.deepEqual(Buffer.from(output.subarray(54)),Buffer.from(original.subarray(33)));
 const jpeg=jpg(),encoded=imageDensity(jpeg,'jpeg',600),pos=Buffer.from(encoded).indexOf(Buffer.from('JFIF\0'));assert.ok(pos>=0);assert.equal(encoded[pos+7],1);assert.equal(new DataView(encoded.buffer).getUint16(pos+8),600);
});
test('um arquivo criado por outro processo na pasta do lote não é substituído nem removido',async()=>{
 const root=await mkdtemp(path.join(tmpdir(),'4ps-images-'));try{
  const store=new ImageExportStore({}, {showOpenDialog:async()=>({filePaths:[root]})});const job=await store.request({action:'begin',plan:{base:'a',format:'png',names:['a.png']}});await store.request({action:'write',token:job.token,index:0},png());await writeFile(path.join(job.location,'a.png'),'outro operador');await assert.rejects(store.request({action:'finish',token:job.token}));assert.equal(await readFile(path.join(job.location,'a.png'),'utf8'),'outro operador');assert.deepEqual(await readdir(job.location),['a.png']);assert.equal(store.active,false);
 }finally{await rm(root,{recursive:true,force:true});}
});

test('cancelar uma gravação em andamento aborta os bytes e limpa os arquivos do lote',async()=>{
 const root=await mkdtemp(path.join(tmpdir(),'4ps-images-'));try{
  const store=new ImageExportStore({}, {showOpenDialog:async()=>({filePaths:[root]})});const job=await store.request({action:'begin',plan:{base:'a',format:'png',names:['a.png']}});
  const writing=store.request({action:'write',token:job.token,index:0},png());const cancel=store.cancelAll();await Promise.allSettled([writing,cancel]);assert.equal(store.active,false);assert.deepEqual(await readdir(root),[]);
 }finally{await rm(root,{recursive:true,force:true});}
});

test('destino do navegador recusa conflito de pasta e remove somente arquivos criados',async()=>{
 const original=globalThis.showDirectoryPicker;
 try{
  let creates=0;globalThis.showDirectoryPicker=async()=>({getDirectoryHandle:async(_name,options)=>{if(options?.create)creates++;return {};}});
  const plan={base:'teste',format:'png',names:['teste_p0001.png']};await assert.rejects(imageDestination(plan),/já existe/);assert.equal(creates,0);
  const data=new Map(),removed=[];const missing=()=>new DOMException('Ausente','NotFoundError');
  const folder={async getFileHandle(name,options){if(!options?.create&&!data.has(name))throw missing();if(!data.has(name))data.set(name,null);return {async createWritable(){return {async write(value){data.set(name,value);},async close(){},async abort(){}};}};},async removeEntry(name){removed.push(name);data.delete(name);}};
  globalThis.showDirectoryPicker=async()=>({async getDirectoryHandle(_name,options){if(!options?.create)throw missing();return folder;},async removeEntry(){}});
  const destination=await imageDestination(plan);await destination.write(0,png());assert.equal(data.size,1);await destination.cancel();assert.equal(data.size,0);assert.deepEqual(removed,plan.names);
  const conflicting=await imageDestination(plan);data.set(plan.names[0],'outro operador');await assert.rejects(conflicting.write(0,png()),/já existe/);await conflicting.cancel();assert.equal(data.get(plan.names[0]),'outro operador');
 }finally{if(original===undefined)delete globalThis.showDirectoryPicker;else globalThis.showDirectoryPicker=original;}
});