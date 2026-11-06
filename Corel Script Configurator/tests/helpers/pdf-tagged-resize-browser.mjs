// Mesmo fluxo no Chrome/Electron; somente sintéticos permanentes. Original opcional fica local.
export async function exercitarRedimensionamentoMarcado(fixtures) {
    const check=(value,message)=>{if(!value)throw Error(message);},sleep=ms=>new Promise(r=>setTimeout(r,ms));
    const wait=async(fn,label)=>{for(let n=0;n<600;n++){if(await fn())return;await sleep(100);}throw Error('Espera: '+label);};
    const q=s=>document.querySelector(s),field=id=>q('#pdf-redim-'+id),bytes=s=>Uint8Array.from(atob(s),c=>c.charCodeAt(0));
    const {PDFWorkspace}=await import('/js/tools/pdf-workspace.js'),{PDFDocument,PDFName}=await import('/vendor/pdf-lib/pdf-lib.js');
    const {exportImages,imageSize}=await import('/js/tools/pdf-image-export.js'),pdfjs=await import('/vendor/pdfjs/pdf.js');
    pdfjs.GlobalWorkerOptions.workerSrc='/vendor/pdfjs/pdf.worker.js';
    const originalOpen=PDFWorkspace.prototype.abrir,originalPrepare=PDFWorkspace.prototype.prepararRedimensionamento;let workspace,previewBytes;
    PDFWorkspace.prototype.abrir=async function(...args){const r=await originalOpen.apply(this,args);workspace=this;return r;};
    PDFWorkspace.prototype.prepararRedimensionamento=async function(...args){const r=await originalPrepare.apply(this,args);previewBytes=await r.pdf.save();return r;};
    const reports=[],confirmOriginal=window.confirm;
    // Chrome headless não opera o seletor nativo; a gravação IPC real é testada no Electron.
    if(!window.desktop4PrintShops)window.showDirectoryPicker=async()=>{const files=new Map(),folders=new Map();return{async getDirectoryHandle(name,{create}={}){if(folders.has(name))return folders.get(name);if(!create)throw new DOMException('Não existe','NotFoundError');const folder={async getFileHandle(filename,{create}={}){if(!files.has(filename)&&!create)throw new DOMException('Não existe','NotFoundError');return{async createWritable(){return{async write(data){files.set(filename,data);},async close(){},async abort(){}};}};},async removeEntry(name){files.delete(name);}};folders.set(name,folder);return folder;},async removeEntry(name){folders.delete(name);}};};
    async function render(data,width=720){const task=pdfjs.getDocument({data:new Uint8Array(data)});try{const pdf=await task.promise,page=await pdf.getPage(1),base=page.getViewport({scale:1}),view=page.getViewport({scale:width/base.width}),canvas=document.createElement('canvas');canvas.width=Math.ceil(view.width);canvas.height=Math.ceil(view.height);const ctx=canvas.getContext('2d',{alpha:false});await page.render({canvas,canvasContext:ctx,viewport:view,background:'rgb(255,255,255)'}).promise;const result={text:(await page.getTextContent()).items.map(i=>i.str).join('\n'),pixels:ctx.getImageData(0,0,canvas.width,canvas.height).data,width:canvas.width,height:canvas.height};canvas.width=canvas.height=0;return result;}finally{await task.destroy();}}
    try {
        q('[data-tool=pdfs]').click();await wait(()=>q('#pdf-arquivo'),'PDFs');
        for(const fixture of fixtures){
            window.confirm=()=>true;const hadPending=workspace?.alterado,source=bytes(fixture.data),dt=new DataTransfer();dt.items.add(new File([source],'Synthetic-local.pdf'));q('#pdf-arquivo').files=dt.files;q('#pdf-arquivo').dispatchEvent(new Event('change',{bubbles:true}));
            if(hadPending){await wait(()=>q('#pdf-confirmar-descarte').open,'proteção na troca');q('#pdf-descartar-sair').click();}
            await wait(()=>q('#pdf-visualizacao canvas')?.width>0&&!q('#pdf-arquivo').disabled,'importar');
            const before=await workspace.obterBytes(),rev=workspace.revisao,reference=await render(before);
            q('#pdf-ferramenta-redimensionar').click();for(const [id,value]of[['largura','200'],['altura','150']]){field(id).value=value;field(id).dispatchEvent(new Event('input',{bubbles:true}));}
            field('prever').click();await wait(()=>!field('prever').disabled&&field('resultado').textContent&&!field('resultado').textContent.includes('Preparando'),'prévia');
            if(fixture.blocked){
                check(field('aplicar').disabled,'estrutura não suportada deve bloquear');check(!field('imagem').hidden,'alternativa explícita disponível');
                window.confirm=()=>false;field('imagem').click();check(!q('#app-exportacao-dialogo').open,'recusa sem exportar');
                window.confirm=()=>true;field('imagem').click();check(q('#app-exportacao-dialogo').open&&q('#app-exportacao-formato').value==='png','atalho abre PNG');q('#app-exportacao-cancelar').click();
                check(workspace.revisao===rev,'falha não altera workspace');reports.push({name:fixture.name,blocked:true,fallbackConfirmed:true});continue;
            }
            check(!field('aplicar').disabled,'prévia válida: '+field('resultado').textContent);check(q('#pdf-salvar-copia').disabled,'não exportar prévia sem aplicar');check(workspace.revisao===rev,'prévia transacional');
            field('cancelar').click();await wait(()=>!q('#pdf-salvar-copia').disabled,'cancelar');check(workspace.revisao===rev,'cancelar conserva estado');
            field('prever').click();await wait(()=>!field('prever').disabled&&field('resultado').textContent.startsWith('Prévia temporária —'),'nova prévia');field('aplicar').click();await wait(()=>!q('#pdf-salvar-copia').disabled&&field('aplicar').disabled,'aplicar');
            check(workspace.alterado,'aplicar marca pendência');const output=await workspace.obterBytes(),parsed=await PDFDocument.load(output),p=parsed.getPage(0);
            check(output.length===previewBytes.length&&output.every((v,i)=>v===previewBytes[i]),'prévia e PDF exportável usam bytes idênticos');
            check(Math.abs(p.getWidth()*25.4/72-200)<1e-6&&Math.abs(p.getHeight()*25.4/72-150)<1e-6,'dimensões200×150');check(parsed.catalog.has(PDFName.of('StructTreeRoot')),'estrutura preservada');
            const actual=await render(output);check(reference.text===actual.text,'texto extraível idêntico');check(reference.width===actual.width&&reference.height===actual.height,'proporção idêntica');let different=0,max=0;for(let i=0;i<actual.pixels.length;i++){const delta=Math.abs(actual.pixels[i]-reference.pixels[i]);if(delta>2)different++;max=Math.max(max,delta);}check(different/actual.pixels.length<.005,'integridade visual normalizada');
            const images=[];for(const format of ['png','jpeg'])for(const dpi of [150,300,600]){
                let blobBytes;const destination={async write(_,data){blobBytes=data;},async finish(){return{count:1,location:'memória'};},async cancel(){}};
                await exportImages(output,{format,pages:[1],dpi,quality:.92},{destination});const bitmap=await createImageBitmap(new Blob([blobBytes],{type:'image/'+format})),size=imageSize(p.getWidth(),p.getHeight(),dpi);
                check(bitmap.width===size.width&&bitmap.height===size.height,'imagem usa workspace redimensionado');images.push({format,dpi,width:bitmap.width,height:bitmap.height,bytes:blobBytes.length});bitmap.close();
            }
            check(workspace.alterado,'imagens não quitam alterações');
            const toBlob=HTMLCanvasElement.prototype.toBlob;let failed=false,canceled=false;
            const destination={async write(){throw Error('Não deve gravar');},async finish(){throw Error('Não deve finalizar');},async cancel(){canceled=true;}};
            try{HTMLCanvasElement.prototype.toBlob=function(callback){callback(null);};await exportImages(output,{format:'png',pages:[1],dpi:150,quality:.92},{destination});}catch(error){failed=/codificar/.test(error.message);}finally{HTMLCanvasElement.prototype.toBlob=toBlob;}
            check(failed&&canceled,'falha de codificação tratada e destino cancelado');
            q('#pdf-salvar-copia').click();q('#app-exportacao-formato').value='png';q('#app-exportacao-formato').dispatchEvent(new Event('change'));q('#app-exportacao-dpi').value='150';q('#app-exportacao-confirmar').click();
            await wait(()=>!q('#app-exportacao-confirmar').disabled,'exportação UI');check(!q('#app-exportacao-dialogo').open,'PNG UI concluído');check(workspace.alterado,'exportação UI conserva pendências');
            reports.push({name:fixture.name,widthMm:200,heightMm:150,scale:200/254,textPreserved:true,components:actual.pixels.length,different,max,images,uiPNG:true,pendingPreserved:true,previewMatchesOutput:true,encodeErrorRecovered:true});
        }
        return {reports};
    }finally{window.confirm=confirmOriginal;PDFWorkspace.prototype.abrir=originalOpen;PDFWorkspace.prototype.prepararRedimensionamento=originalPrepare;}
}
