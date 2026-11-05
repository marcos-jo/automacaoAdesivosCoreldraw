// Executada em Chromium: compara regiões reais dos Forms, sem guias da interface.
export async function compararLivretoBrowser(source,result){
    const js=await import('/vendor/pdfjs/pdf.js'),tasks=[],canvases=[],scale=25.4/72*2;
    async function open(bytes){const t=js.getDocument({data:bytes.slice()});tasks.push(t);return t.promise;}
    async function render(doc,n,{paper=false}={}){
        const p=await doc.getPage(n);let v=p.getViewport({scale});
        if(paper){const top=v.convertToViewportPoint(0,result.plan.folha.altura*72/25.4);v=p.getViewport({scale,offsetX:-top[0],offsetY:-top[1]});}
        const c=document.createElement('canvas');canvases.push(c);c.width=paper?Math.round(result.plan.folha.largura*2):Math.round(v.width);c.height=paper?Math.round(result.plan.folha.altura*2):Math.round(v.height);
        await p.render({canvas:c,canvasContext:c.getContext('2d'),viewport:v,background:'white',intent:'print'}).promise;return c;
    }
    const proof={components:0,different:0,max:0,pairs:0,textMatches:0};
    function equal(a,b){if(a.length!==b.length)throw Error('Regiões com dimensões diferentes.');proof.components+=a.length;proof.pairs++;for(let i=0;i<a.length;i++){const d=Math.abs(a[i]-b[i]);if(d)proof.different++;proof.max=Math.max(proof.max,d);}}
    const pixels=(c,x,y,w,h)=>c.getContext('2d').getImageData(Math.round(x),Math.round(y),Math.round(w),Math.round(h)).data;
    try{
        const [src,out,read,preview]=await Promise.all([open(source),open(result.bytes),open(result.bytesLeitura),open(result.bytesPrevia)]),originals=new Map(),physical=new Map();
        for(let i=1;i<=src.numPages;i++)originals.set(i,await render(src,i));
        const w=result.plan.final.largura*2,h=result.plan.final.altura*2;
        for(const [i,face]of result.plan.faces.entries()){
            const c=await render(out,i+1),board=await render(preview,i+1,{paper:true});equal(pixels(c,0,0,c.width,c.height),pixels(board,0,0,c.width,c.height));
            const items=(await(await out.getPage(i+1)).getTextContent()).items.filter(t=>/^PAGINA \d+$/.test(t.str)).map(t=>Number(t.str.split(' ')[1]));
            if(JSON.stringify(items)!==JSON.stringify(face.slots.filter(s=>s.origem).map(s=>s.origem)))throw Error('Texto/paginação da face não corresponde ao plano.');proof.textMatches+=items.length;
            for(const slot of face.slots){const f=slot.moldura,part=pixels(c,f.x*2,c.height-(f.y+f.height)*2,w,h);if(slot.origem){equal(part,pixels(originals.get(slot.origem),0,0,w,h));physical.set(slot.logica,part);}else if(part.some(v=>v!==255))throw Error('Página branca recebeu conteúdo.');}
        }
        for(const [i,opening]of result.plan.paginacao.aberturas.entries()){
            const c=await render(read,i+1);for(const [side,slot]of opening.entries())if(slot?.origem)equal(pixels(c,side*w,0,w,h),physical.get(slot.logica));
        }
        return proof;
    }finally{for(const t of tasks)await t.destroy();for(const c of canvases)c.width=c.height=0;}
}

// A saída Windows pode conter glifos em contornos. Compara aparência/posição, não OCR.
export async function compararSaidaVirtualLivretoBrowser(source,printed){
    const js=await import('/vendor/pdfjs/pdf.js'),tasks=[],workers=[],scale=25.4/72*2;let before,after;
    async function open(bytes){const worker=new js.PDFWorker();workers.push(worker);const t=js.getDocument({data:bytes.slice(),worker});tasks.push(t);return t.promise;}
    async function pixels(doc,n){const p=await doc.getPage(n),c=document.createElement('canvas');c.width=420;c.height=594;try{await p.render({canvas:c,canvasContext:c.getContext('2d'),viewport:p.getViewport({scale}),background:'white',intent:'print'}).promise;return c.getContext('2d').getImageData(0,0,c.width,c.height).data;}finally{c.width=c.height=0;}}
    try{
        [before,after]=await Promise.all([open(source),open(printed)]);if(before.numPages!==after.numPages)throw Error('Quantidade de faces impressas diferente.');
        const pages=[];for(let n=1;n<=before.numPages;n++){const a=await pixels(before,n),b=await pixels(after,n);let above16=0,totalDifference=0;for(let i=0;i<a.length;i++){const d=Math.abs(a[i]-b[i]);if(d>16)above16++;totalDifference+=d;}pages.push({face:n,components:a.length,above16,fraction:above16/a.length,meanDifference:totalDifference/a.length});}return {pages,physicalPrint:false};
    }finally{await Promise.all(tasks.map(t=>t.destroy()));for(const w of workers)w.destroy();}
}
