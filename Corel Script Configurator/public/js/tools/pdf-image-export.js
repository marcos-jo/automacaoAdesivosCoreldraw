export const IMAGE_DPI = [150,300,600];
export function imageSize(width,height,dpi){
    if(!IMAGE_DPI.includes(dpi))throw Error('Escolha 150, 300 ou 600 DPI.');
    const w=Math.ceil(width*dpi/72-1e-7),h=Math.ceil(height*dpi/72-1e-7),pixels=w*h;
    if(![w,h,pixels].every(Number.isSafeInteger)||w<1||h<1||w>12000||h>12000||pixels>36000000)throw Error('Página acima de 36 megapixels ou 12.000 pixels por lado. Escolha menos DPI.');
    return {width:w,height:h,pixels,estimatedMemory:pixels*16};
}
export function imagePages(mode,text,total,current,selected=[]){
    if(mode==='all')return Array.from({length:total},(_,i)=>i+1);
    if(mode==='current')return [current];
    if(mode==='selected'){
        if(!selected.length)throw Error('Selecione páginas nas miniaturas antes de exportar.');
        return [...selected].sort((a,b)=>a-b);
    }
    if(mode!=='range')throw Error('Modo de páginas inválido.');
    const pages=new Set();
    if(!text.trim())throw Error('Informe páginas ou intervalos, por exemplo 1,3-5.');
    for(const part of text.split(',')){
        const match=part.trim().match(/^(\d+)(?:\s*-\s*(\d+))?$/);
        if(!match)throw Error('Intervalo de páginas inválido.');
        const from=Number(match[1]),to=Number(match[2]||match[1]);
        if(from<1||to<from||to>total)throw Error('Página fora do documento.');
        for(let n=from;n<=to;n++)pages.add(n);
    }
    return [...pages].sort((a,b)=>a-b);
}
export function imageNames(base,pages,format){
    if(!base||base.length>90||/[\\/:*?"<>|\x00-\x1f]/.test(base)||/[. ]$/.test(base)||/^(con|prn|aux|nul|com[1-9]|lpt[1-9])(?:\.|$)/i.test(base))throw Error('Informe um nome válido de até 90 caracteres.');
    return pages.map(page=>`${base}_p${String(page).padStart(4,'0')}.${format==='png'?'png':'jpg'}`);
}
// Ajusta apenas a densidade declarada, sem recodificar os pixels comprimidos.
export function imageDensity(bytes,format,dpi){
    if(!IMAGE_DPI.includes(dpi))throw Error('DPI inválido.');
    if(format==='jpeg'){
        const result=new Uint8Array(bytes);
        for(let pos=2;pos+4<result.length;){
            if(result[pos]!==255||result[pos+1]===218)break;
            const length=(result[pos+2]<<8)|result[pos+3];if(length<2)break;
            if(result[pos+1]===224&&length>=16&&String.fromCharCode(...result.subarray(pos+4,pos+9))==='JFIF\0'){
                const view=new DataView(result.buffer);result[pos+11]=1;view.setUint16(pos+12,dpi);view.setUint16(pos+14,dpi);return result;
            }
            pos+=length+2;
        }
        const header=new Uint8Array([255,224,0,16,74,70,73,70,0,1,1,1,dpi>>8,dpi&255,dpi>>8,dpi&255,0,0]);
        const output=new Uint8Array(bytes.length+header.length);output.set(bytes.subarray(0,2));output.set(header,2);output.set(bytes.subarray(2),2+header.length);return output;
    }
    const chunk=new Uint8Array(21),view=new DataView(chunk.buffer);view.setUint32(0,9);chunk.set([112,72,89,115],4);view.setUint32(8,Math.round(dpi/0.0254));view.setUint32(12,Math.round(dpi/0.0254));chunk[16]=1;
    let crc=-1;for(const value of chunk.subarray(4,17)){crc^=value;for(let bit=0;bit<8;bit++)crc=(crc>>>1)^((crc&1)?0xedb88320:0);}view.setUint32(17,(~crc)>>>0);
    const chunks=[bytes.subarray(0,33),chunk];let size=33+chunk.length;
    for(let pos=33;pos+12<=bytes.length;){const length=new DataView(bytes.buffer,bytes.byteOffset+pos,4).getUint32(0)+12;if(pos+length>bytes.length)throw Error('PNG codificado inválido.');if(String.fromCharCode(...bytes.subarray(pos+4,pos+8))!=='pHYs'){chunks.push(bytes.subarray(pos,pos+length));size+=length;}pos+=length;}
    const output=new Uint8Array(size);let offset=0;for(const part of chunks){output.set(part,offset);offset+=part.length;}return output;
}const canceled=signal=>{if(signal?.aborted)throw new DOMException('Exportação cancelada','AbortError');};
export async function imageDestination(plan){
    const desktop=globalThis.window?.desktop4PrintShops;
    if(desktop?.exportarImagens){
        const result=await desktop.exportarImagens({action:'begin',plan});
        if(result.canceled)throw new DOMException('Exportação cancelada','AbortError');
        return {write:(index,bytes)=>desktop.exportarImagens({action:'write',token:result.token,index},bytes),finish:()=>desktop.exportarImagens({action:'finish',token:result.token}),cancel:()=>desktop.exportarImagens({action:'cancel',token:result.token})};
    }
    if(!globalThis.showDirectoryPicker)throw Error('Para escolher a pasta, utilize Chrome/Edge com acesso a diretórios ou a versão Desktop.');
    const root=await showDirectoryPicker({mode:'readwrite'});
    const name=plan.base+'_'+plan.format.toUpperCase()+'_'+crypto.randomUUID();
    try{await root.getDirectoryHandle(name);throw Error('A pasta de saída já existe. Nenhum arquivo foi alterado.');}
    catch(error){if(error.name!=='NotFoundError')throw error;}
    const folder=await root.getDirectoryHandle(name,{create:true});
    const created=[];
    return {
        async write(index,bytes){
            const filename=plan.names[index];
            try{await folder.getFileHandle(filename);throw Error('O arquivo de saída já existe. Exportação interrompida.');}
            catch(error){if(error.name!=='NotFoundError')throw error;}
            const file=await folder.getFileHandle(filename,{create:true});created.push(filename);
            const writer=await file.createWritable();try{await writer.write(bytes);await writer.close();}catch(error){await writer.abort().catch(()=>{});throw error;}
        },
        async finish(){return {saved:true,location:name,count:plan.names.length};},
        async cancel(){for(const filename of created)await folder.removeEntry(filename).catch(()=>{});await root.removeEntry(name);}
    };
}
export async function exportImages(bytes,config,{signal,progress=()=>{},destination}={}){
    const pdfjs=await import('/vendor/pdfjs/pdf.js');pdfjs.GlobalWorkerOptions.workerSrc='/vendor/pdfjs/pdf.worker.js';
    const task=pdfjs.getDocument({data:new Uint8Array(bytes)});
    let destruction; const destroy=()=>destruction??=(task.destroy());
    const stop=()=>{destroy().catch(()=>{});destination.cancel().catch(()=>{});};signal?.addEventListener('abort',stop,{once:true});
    let canvas;
    try{
        canceled(signal);const pdf=await task.promise;
        if(!config.pages.length||config.pages.length>1000||new Set(config.pages).size!==config.pages.length||config.pages.some(n=>!Number.isInteger(n)||n<1||n>pdf.numPages))throw Error('Seleção inválida ou acima de 1000 páginas.');
        if(!['png','jpeg'].includes(config.format)||!Number.isFinite(config.quality)||config.quality<.1||config.quality>1)throw Error('Formato ou qualidade inválidos.');
        let total=0;
        for(const number of config.pages){canceled(signal);const page=await pdf.getPage(number),base=page.getViewport({scale:1});total+=imageSize(base.width,base.height,config.dpi).pixels;page.cleanup();}
        if(total>1500000000)throw Error('O lote ultrapassa 1,5 bilhão de pixels. Exporte menos páginas ou reduza os DPI.');
        for(let i=0;i<config.pages.length;i++){
            canceled(signal);const page=await pdf.getPage(config.pages[i]),view=page.getViewport({scale:config.dpi/72}),size=imageSize(page.getViewport({scale:1}).width,page.getViewport({scale:1}).height,config.dpi);
            progress(`Página ${config.pages[i]} · ${i+1}/${config.pages.length} · ${size.width} × ${size.height} px · pico estimado ${Math.ceil(size.estimatedMemory/1048576)} MiB`);
            canvas=document.createElement('canvas');canvas.width=size.width;canvas.height=size.height;
            const context=canvas.getContext('2d',{alpha:false});if(!context)throw Error('Memória de imagem indisponível.');
            context.fillStyle='white';context.fillRect(0,0,canvas.width,canvas.height);
            const render=page.render({canvas,canvasContext:context,viewport:view,intent:'print',background:'rgb(255,255,255)'});
            let timer;const abort=()=>render.cancel();signal?.addEventListener('abort',abort,{once:true});
            try{await Promise.race([render.promise,new Promise((_,reject)=>{timer=setTimeout(()=>{render.cancel();reject(Error('Tempo limite ao renderizar a página.'));},60000);})]);}
            finally{clearTimeout(timer);signal?.removeEventListener('abort',abort);}
            canceled(signal);
            const blob=await new Promise((resolve,reject)=>canvas.toBlob(value=>value?resolve(value):reject(Error('Não foi possível codificar a imagem.')),'image/'+config.format,config.quality));
            canceled(signal);await destination.write(i,imageDensity(new Uint8Array(await blob.arrayBuffer()),config.format,config.dpi));
            canvas.width=canvas.height=0;canvas=null;page.cleanup();
            await new Promise(resolve=>setTimeout(resolve,0));
        }
        canceled(signal);return await destination.finish();
    }catch(error){await destination.cancel().catch(()=>{});if(signal?.aborted)throw new DOMException('Exportação cancelada','AbortError');throw error;}
    finally{signal?.removeEventListener('abort',stop);if(canvas)canvas.width=canvas.height=0;await destroy();}
}
