import assert from 'node:assert/strict';
import {writeFile} from 'node:fs/promises';
import path from 'node:path';
const sleep=ms=>new Promise(r=>setTimeout(r,ms));
// CDP intercepta o gesto nativo. Nenhum scrollTop é atribuído durante o arraste.
export async function testarArrasteMiniaturas({ev,cdp,wait,folder,screenshot,socket}){
 let data;const reports=[];
 const intercepted=e=>{const m=JSON.parse(e.data);if(m.method==='Input.dragIntercepted')data=m.params.data;};socket.addEventListener('message',intercepted);
 const ready=()=>wait(()=>ev('!!window.thumbs?.pdf&&!document.querySelector("#pdf-arquivo").disabled&&!thumbs.ocupado'),'editor pronto');
 const state=()=>ev(`({scroll:thumbs.container.scrollTop,frame:thumbs.arraste.frame,origin:thumbs.arraste.origem,position:thumbs.arraste.posicao,buttons:thumbs.itens.size,canvases:thumbs.container.querySelectorAll('canvas').length,statusPointer:document.querySelector('#pdf-status')?getComputedStyle(document.querySelector('#pdf-status')).pointerEvents:null})`);
 const point=where=>ev(`(()=>{const r=thumbs.container.getBoundingClientRect();return {x:r.x+r.width/2,y:${where==='bottom'?'Math.min(r.bottom,innerHeight)-4':where==='top'?'Math.max(r.top,0)+4':'r.y+r.height/2'}};})()`);
 const hover=async p=>{await cdp('Input.dispatchMouseEvent',{type:'mouseMoved',...p,button:'left',buttons:1});await cdp('Input.dispatchDragEvent',{type:'dragEnter',...p,data});await cdp('Input.dispatchDragEvent',{type:'dragOver',...p,data});};
 const start=async number=>{
  data=null;const p=await ev(`(()=>{const r=thumbs.container.querySelector('button[data-pagina="${number}"]').getBoundingClientRect();return {x:r.x+r.width/2,y:r.y+r.height/2};})()`);
  await cdp('Input.dispatchMouseEvent',{type:'mouseMoved',...p});await cdp('Input.dispatchMouseEvent',{type:'mousePressed',...p,button:'left',clickCount:1});
  await cdp('Input.dispatchMouseEvent',{type:'mouseMoved',x:p.x+8,y:p.y+12,button:'left',buttons:1});await sleep(30);
  await cdp('Input.dispatchMouseEvent',{type:'mouseMoved',x:p.x+10,y:p.y-40,button:'left',buttons:1});await wait(()=>!!data,'arraste nativo');assert.equal((await state()).origin,number);
 };
 // O CDP exige destino neutro antes de cancelar seu arraste interceptado (como no roteiro geral).
 const end=async(type,p)=>{if(type==='dragCancel'){p=await point('center');await hover(p);}if(type==='drop')await hover(p);await cdp('Input.dispatchDragEvent',{type,...p,data});await cdp('Input.dispatchMouseEvent',{type:'mouseReleased',...p,button:'left',clickCount:1});await wait(()=>ev('!thumbs.arraste.origem&&!thumbs.arraste.frame'),'fim do gesto');};
 const stable=async()=>{const a=await state();await sleep(350);const b=await state();assert.equal(b.scroll,a.scroll,'rolagem parada');assert.equal(b.frame,null,'sem animação remanescente');};
 const order=()=>ev('dragWorkspace.pdf.getPages().map(p=>p.ref.toString())');
 const resetScroll=()=>ev('thumbs.container.scrollTop=0;thumbs.container.dispatchEvent(new Event("scroll"))');
 const open=async count=>{
  await ev(`(async()=>{const{PDFDocument}=await import('/vendor/pdf-lib/pdf-lib.js'),p=await PDFDocument.create();for(let i=1;i<=${count};i++)p.addPage([400+i/1000,500]).drawText('PAGE '+i,{x:30,y:70});const bytes=await p.save();window.dragOriginal=Array.from(bytes);window.dragInput=new File([bytes],'Autoscroll ${count}.pdf');const t=new DataTransfer();t.items.add(dragInput);const e=document.querySelector('#pdf-arquivo');e.files=t.files;e.dispatchEvent(new Event('change',{bubbles:true}));})()`);
  await wait(()=>ev(`document.querySelector('#pdf-confirmar-descarte').open||(!document.querySelector('#pdf-arquivo').disabled&&window.thumbs?.pdf?.numPages===${count})`),'importação');
  await ev(`if(document.querySelector('#pdf-confirmar-descarte').open)document.querySelector('#pdf-descartar-sair').click()`);await ready();await wait(()=>ev(`thumbs.pdf.numPages===${count}`),'contagem');
 };
 try{
  await ev(`(async()=>{const{PDFMiniaturas}=await import('/js/tools/pdf-thumbnails.js');const old=PDFMiniaturas.prototype.definirDocumento;PDFMiniaturas.prototype.definirDocumento=function(...args){window.thumbs=this;return old.apply(this,args);};const{PDFWorkspace}=await import('/js/tools/pdf-workspace.js');const oldOpen=PDFWorkspace.prototype.abrir;PDFWorkspace.prototype.abrir=function(...args){window.dragWorkspace=this;return oldOpen.apply(this,args);};window.dragErrors=[];window.addEventListener('unhandledrejection',e=>dragErrors.push(String(e.reason)));})()`);
  await ev('document.querySelector("[data-tool=pdfs]").click()');await wait(()=>ev('!!document.querySelector("#pdf-arquivo")'),'editor');await cdp('Input.setInterceptDrags',{enabled:true});
  for(const [count,width,height]of (process.env.THUMB_DRAG_LIFECYCLE_ONLY==='1'?[]:[[8,1024,600],[100,1440,1000],[350,1366,768]])){
   await cdp('Emulation.setDeviceMetricsOverride',{width,height,deviceScaleFactor:1,mobile:false});await open(count);const original=await order();
   await start(1);const bottom=await point('bottom');await hover(bottom);
   assert.equal(await ev(`thumbs.container.contains(document.elementFromPoint(${bottom.x},${bottom.y}))`),true,'aviso não intercepta a borda');
   await wait(()=>ev('thumbs.container.scrollTop>=thumbs.container.scrollHeight-thumbs.container.clientHeight-1'),'última miniatura por autoscroll');
   const lower=await state();assert.equal(lower.position,count);assert.ok(lower.buttons<=16);assert.ok(lower.canvases<=16);await screenshot(`inferior-${count}`);
   await end('drop',bottom);await ready();assert.deepEqual(await order(),[...original.slice(1),original[0]]);
   const widths=await ev(`(async()=>{const{PDFDocument}=await import('/vendor/pdf-lib/pdf-lib.js');const p=await PDFDocument.load(await dragWorkspace.obterBytes());return p.getPages().map(p=>p.getWidth());})()`);
   assert.deepEqual(widths.map(n=>Math.round((n-400)*1000)),[...Array.from({length:count-1},(_,i)=>i+2),1],'ordem do PDF serializado');
   await ev('thumbs.container.scrollTop=thumbs.container.scrollHeight;thumbs.container.dispatchEvent(new Event("scroll"))');await wait(()=>ev(`!!thumbs.itens.get(${count})`),'última origem');
   await start(count);const top=await point('top');await hover(top);await wait(()=>ev('thumbs.container.scrollTop===0'),'primeira miniatura por autoscroll');assert.equal((await state()).position,0);
   await end('drop',top);await ready();assert.deepEqual(await order(),original);await stable();
   assert.equal(await ev(`(async()=>{const b=new Uint8Array(await dragInput.arrayBuffer());return b.every((v,i)=>v===dragOriginal[i]);})()`),true);
   reports.push({count,width,height,firstToLast:true,lastToFirst:true,serializedOrder:true,originalIntact:true,virtualWindow:lower.buttons});console.log('Extremos e ordem aprovados: '+count+' páginas.');
  }
  await open(100);const original=await order(),outside={x:700,y:250};
  await start(1);await hover(await point('bottom'));await wait(()=>ev('thumbs.container.scrollTop>700'),'rolagem para cancelar');
  await hover(await point('center'));await stable();await hover(await point('bottom'));const resumed=await state();await wait(()=>ev(`thumbs.container.scrollTop>${resumed.scroll}+100`),'retomada');
  await end('dragCancel',await point('bottom'));await stable();assert.deepEqual(await order(),original);assert.notEqual((await state()).statusPointer,'none');
  assert.equal(await ev('document.querySelector(".pdf-miniatura-insercao")===null'),true);
  await resetScroll();await start(1);await hover(await point('bottom'));await wait(()=>ev('thumbs.container.scrollTop>200'),'rolagem antes de sair');await hover(outside);await stable();await end('drop',outside);await stable();assert.deepEqual(await order(),original);
  await resetScroll();await start(1);await hover(await point('bottom'));await wait(()=>ev('!!thumbs.arraste.frame'),'animação ativa');
  await ev('document.querySelector("[data-tool=qrcodes]").click()');await wait(()=>ev('thumbs.suspensa&&!thumbs.arraste.origem'),'aba suspensa');await end('dragCancel',outside);await stable();
  await ev('document.querySelector("[data-tool=pdfs]").click()');await ready();assert.deepEqual(await order(),original);
  await resetScroll();await start(1);await hover(await point('bottom'));await wait(()=>ev('!!thumbs.arraste.frame'),'troca com animação ativa');await open(8);await end('dragCancel',outside);await stable();
  await start(1);await hover(await point('bottom'));await wait(()=>ev('!!thumbs.arraste.frame'),'fechamento com animação ativa');
  await ev('window.closedThumbs=thumbs;document.querySelector("[data-close-tool=pdfs]").click()');await wait(()=>ev('closedThumbs.encerrada&&!closedThumbs.arraste.origem&&!closedThumbs.arraste.frame'),'editor encerrado');await end('dragCancel',outside);await stable();assert.equal(await ev('closedThumbs.container.children.length'),0);assert.deepEqual(await ev('dragErrors'),[]);
  const listeners=await cdp('Runtime.evaluate',{expression:'Object.keys(getEventListeners(closedThumbs.container))',includeCommandLineAPI:true,returnByValue:true});assert.deepEqual(listeners.result.value,[],'listeners do painel removidos');
  await writeFile(path.join(folder,'results.json'),JSON.stringify({passed:true,reports,cancelDuringScroll:true,outsideDropNoMove:true,pauseReplaceClose:true,noPrinting:true},null,2));
 }finally{socket.removeEventListener('message',intercepted);}
}
