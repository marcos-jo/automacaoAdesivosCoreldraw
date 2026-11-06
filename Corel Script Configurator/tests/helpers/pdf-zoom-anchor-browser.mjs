import assert from 'node:assert/strict';import{writeFile}from'node:fs/promises';import path from'node:path';
export async function diagnosticarZoom({ev,cdp,wait,folder,screenshot}){
    await cdp('Emulation.setDeviceMetricsOverride',{width:Number(process.env.UX27_WIDTH||1440),height:Number(process.env.UX27_HEIGHT||1000),deviceScaleFactor:1,mobile:false});
    await ev(`(async()=>{
        const{PDFControlesVisualizacao}=await import('/js/tools/pdf-view-controls.js');window.zoomTrace=[];
        window.zoomSnapshot=()=>{const c=document.querySelector('#pdf-visualizacao'),p=c?.querySelector('canvas');if(!c||!p)return null;const box=e=>{const r=e.getBoundingClientRect();return {x:r.x,y:r.y,width:r.width,height:r.height};};return {container:box(c),canvas:box(p),client:[c.clientWidth,c.clientHeight],scroll:[c.scrollLeft,c.scrollTop],extent:[c.scrollWidth,c.scrollHeight],zoom:window.zoomControl?.zoom,generation:window.zoomControl?.geracao,anchor:window.zoomControl?.ancora};};
        window.zoomOriginalAmpliar=PDFControlesVisualizacao.prototype.ampliar;window.zoomOriginalSolicitar=PDFControlesVisualizacao.prototype.solicitar;
        PDFControlesVisualizacao.prototype.ampliar=function(...args){window.zoomControl=this;zoomTrace.push({event:'zoom-before',state:zoomSnapshot()});return zoomOriginalAmpliar.apply(this,args);};
        PDFControlesVisualizacao.prototype.solicitar=function(...args){zoomTrace.push({event:'request',resize:args[0]||false,state:zoomSnapshot()});return zoomOriginalSolicitar.apply(this,args);};
    })()`);
    await ev('document.querySelector("[data-tool=pdfs]").click()');await wait(()=>ev('!!document.querySelector("#pdf-arquivo")'),'editor');
    await ev(`(async()=>{const{PDFDocument}=await import('/vendor/pdf-lib/pdf-lib.js'),p=await PDFDocument.create();for(const[w,h]of [[600,900],[900,600]])p.addPage([w,h]).drawText('Zoom geométrico',{x:40,y:100});const t=new DataTransfer();t.items.add(new File([await p.save()],'Zoom sintetico.pdf'));const e=document.querySelector('#pdf-arquivo');e.files=t.files;e.dispatchEvent(new Event('change',{bubbles:true}));})()`);
    await wait(()=>ev('document.querySelector("#pdf-pagina-atual").textContent==="1 / 2"&&!document.querySelector("#pdf-arquivo").disabled'),'origem');await ev('document.querySelector("#pdf-proxima").click()');await wait(()=>ev('document.querySelector("#pdf-pagina-atual").textContent==="2 / 2"'),'segunda');
    await ev('document.querySelector("#pdf-zoom-real").click()');await wait(()=>ev('document.querySelector("#pdf-visualizacao canvas").width===900'),'tamanho real');await new Promise(r=>setTimeout(r,150));
    await ev(`(()=>{const c=document.querySelector('#pdf-visualizacao');for(const name of ['scrollLeft','scrollTop']){const descriptor=Object.getOwnPropertyDescriptor(Element.prototype,name);Object.defineProperty(c,name,{configurable:true,get(){return descriptor.get.call(this);},set(v){zoomTrace.push({event:name,value:v,before:zoomSnapshot()});descriptor.set.call(this,v);zoomTrace.push({event:name+'-after',state:zoomSnapshot()});}});}c.addEventListener('scroll',()=>zoomTrace.push({event:'scroll',state:zoomSnapshot()}));zoomTrace.length=0;})()`);
    const before=await ev('zoomSnapshot()'),point={x:before.container.x+before.container.width/2,y:before.container.y+before.container.height/2},fraction={x:(point.x-before.canvas.x)/before.canvas.width,y:(point.y-before.canvas.y)/before.canvas.height};
    await cdp('Input.dispatchMouseEvent',{type:'mouseWheel',...point,deltaX:0,deltaY:-180});await wait(()=>ev('document.querySelector("#pdf-visualizacao canvas").width>1000'),'zoom');await new Promise(r=>setTimeout(r,600));
    const after=await ev('zoomSnapshot()'),trace=await ev('zoomTrace'),error={x:after.canvas.x+fraction.x*after.canvas.width-point.x,y:after.canvas.y+fraction.y*after.canvas.height-point.y};
    const required={x:after.canvas.x+after.scroll[0]+fraction.x*after.canvas.width-point.x,y:after.canvas.y+after.scroll[1]+fraction.y*after.canvas.height-point.y},limits={x:after.extent[0]-after.client[0],y:after.extent[1]-after.client[1]},expected={x:Math.max(0,Math.min(limits.x,required.x)),y:Math.max(0,Math.min(limits.y,required.y))};
    await screenshot('zoom-diagnostico');const result={before,point,fraction,after,error,required,limits,expected,trace};await writeFile(path.join(folder,'zoom-diagnostic.json'),JSON.stringify(result,null,2));
    console.log('Zoom: erro X='+error.x.toFixed(4)+'; Y='+error.y.toFixed(4)+' px.');
    if(process.env.UX27_ZOOM_DIAG!=='1')assert.ok(Math.max(Math.abs(after.scroll[0]-expected.x),Math.abs(after.scroll[1]-expected.y))<1.5,JSON.stringify({error,required,limits,actual:after.scroll}));
    await ev(`(async()=>{const{PDFControlesVisualizacao}=await import('/js/tools/pdf-view-controls.js');PDFControlesVisualizacao.prototype.ampliar=zoomOriginalAmpliar;PDFControlesVisualizacao.prototype.solicitar=zoomOriginalSolicitar;})()`);
    return result;
}

export async function testarMatrizZoom({ev,cdp,wait,folder,screenshot}){
    // Documento real do editor, misto e com uma rotação aplicada no workspace.
    await ev(`(async()=>{const{PDFDocument,degrees}=await import('/vendor/pdf-lib/pdf-lib.js'),p=await PDFDocument.create();for(const[w,h,a]of [[600,900,0],[900,600,0],[420,760,90],[1100,430,270]]){const page=p.addPage([w,h]);page.drawText('Zoom de pagina mista',{x:30,y:100});page.setRotation(degrees(a));}const t=new DataTransfer();t.items.add(new File([await p.save()],'Zoom misto.pdf'));const e=document.querySelector('#pdf-arquivo');e.files=t.files;e.dispatchEvent(new Event('change',{bubbles:true}));})()`);
    await wait(()=>ev('document.querySelector("#pdf-pagina-atual").textContent==="1 / 4"&&!document.querySelector("#pdf-arquivo").disabled'),'misto');
    const results=[];
    const capture=()=>ev(`(()=>{const c=document.querySelector('#pdf-visualizacao'),r=c.querySelector('canvas').getBoundingClientRect(),b=c.getBoundingClientRect();return {box:{x:b.x,y:b.y,width:b.width,height:b.height},canvas:{x:r.x,y:r.y,width:r.width,height:r.height},scroll:[c.scrollLeft,c.scrollTop],limits:[c.scrollWidth-c.clientWidth,c.scrollHeight-c.clientHeight],size:[c.clientWidth,c.clientHeight]};})()`);
    for(const [width,height,density]of [[1366,768,1],[1440,1000,1],[1920,1080,1],[1366,768,1.25],[1024,600,1]]){
        await cdp('Emulation.setDeviceMetricsOverride',{width,height,deviceScaleFactor:density,mobile:false});await new Promise(r=>setTimeout(r,160));
        for(let page=1;page<=4;page++){
            await ev('document.querySelector("#pdf-zoom-real").click()');await new Promise(r=>setTimeout(r,150));
            const current=await ev('Number(document.querySelector("#pdf-pagina-atual").textContent.split(" / ")[0])');
            for(let n=current;n>page;n--){await ev('document.querySelector("#pdf-anterior").click()');await wait(()=>ev(`document.querySelector('#pdf-pagina-atual').textContent==='${n-1} / 4'`),'página anterior');}
            for(let n=current;n<page;n++){await ev('document.querySelector("#pdf-proxima").click()');await wait(()=>ev(`document.querySelector('#pdf-pagina-atual').textContent==='${n+1} / 4'`),'página');}
            assert.equal(await ev('document.querySelector("#pdf-pagina-atual").textContent'),`${page} / 4`);
            for(const action of ['wheel-in','wheel-out','plus','minus','rapid-wheel']){
                const before=await capture(),cx=before.canvas,box=before.box;
                const point=action.includes('wheel')?{x:(Math.max(box.x+20,cx.x)+Math.min(box.x+before.size[0]-20,cx.x+cx.width))/2,y:(Math.max(box.y+20,cx.y)+Math.min(box.y+before.size[1]-20,cx.y+cx.height))/2}:{x:box.x+before.size[0]/2,y:box.y+before.size[1]/2};
                const fraction={x:(point.x-cx.x)/cx.width,y:(point.y-cx.y)/cx.height};
                if(action.includes('wheel')){for(let i=0;i<(action==='rapid-wheel'?6:1);i++)await cdp('Input.dispatchMouseEvent',{type:'mouseWheel',...point,deltaX:0,deltaY:action==='wheel-out'?120:action==='rapid-wheel'?-25:-180});}
                else await ev(`document.querySelector('#pdf-zoom-${action==='plus'?'mais':'menos'}').click()`);
                await new Promise(r=>setTimeout(r,220));const after=await capture(),required=[after.scroll[0]+after.canvas.x+fraction.x*after.canvas.width-point.x,after.scroll[1]+after.canvas.y+fraction.y*after.canvas.height-point.y],expected=required.map((n,i)=>Math.max(0,Math.min(after.limits[i],n))),errors=expected.map((n,i)=>Math.abs(after.scroll[i]-n));
                const record={width,height,density,page,action,errors,required,expected,actual:after.scroll,limits:after.limits,canvas:after.canvas};results.push(record);assert.ok(Math.max(...errors)<1.5,JSON.stringify(record));
            }
            // A próxima página pode ter dimensões e giro completamente diferentes.
        }
    }
    await ev('document.querySelector("#pdf-girar-direita").click()');await wait(()=>ev('!document.querySelector("#pdf-arquivo").disabled'),'giro no workspace');
    for(const button of ['#pdf-painel-alternar','#pdf-miniaturas-alternar']){
        await ev(`document.querySelector(${JSON.stringify(button)}).click()`);await new Promise(r=>setTimeout(r,180));
        const b=await capture(),point={x:b.box.x+b.size[0]/2,y:b.box.y+b.size[1]/2},fx=(point.x-b.canvas.x)/b.canvas.width,fy=(point.y-b.canvas.y)/b.canvas.height;
        await cdp('Input.dispatchMouseEvent',{type:'mouseWheel',...point,deltaX:0,deltaY:-120});await new Promise(r=>setTimeout(r,220));const a=await capture(),required=[a.scroll[0]+a.canvas.x+fx*a.canvas.width-point.x,a.scroll[1]+a.canvas.y+fy*a.canvas.height-point.y],expected=required.map((n,i)=>Math.max(0,Math.min(a.limits[i],n))),errors=expected.map((n,i)=>Math.abs(a.scroll[i]-n));
        const record={action:'workspace-rotation-and-panel-toggle',button,errors,required,expected,actual:a.scroll,limits:a.limits};results.push(record);assert.ok(Math.max(...errors)<1.5,JSON.stringify(record));
    }
    await screenshot('zoom-matriz-final');await writeFile(path.join(folder,'zoom-matrix.json'),JSON.stringify({results,maxError:Math.max(...results.flatMap(r=>r.errors)),productionZoomChanged:false},null,2));
    return results;
}
