// Executado exclusivamente no navegador do teste; a função é serializada pelo CDP.
export async function compararPecaMontadaBrowser(originalBytes,result){
    const pdfjs=await import('/vendor/pdfjs/pdf.js'),tasks=[],canvases=[];
    try{
        const load=async bytes=>{const t=pdfjs.getDocument({data:bytes.slice()});tasks.push(t);return t.promise;};
        const piece=result.plan.folhas[0][0],original=await load(originalBytes),montage=await load(result.bytes),a=await original.getPage(piece.origem),b=await montage.getPage(1),vb=b.getViewport({scale:2}),va=a.getViewport({scale:2*piece.escala,rotation:(a.rotate+piece.rotacao)%360});
        const left=piece.x*72/25.4*2,top=(result.plan.folha.altura-piece.y-piece.altura)*72/25.4*2;
        for(let i=0;i<2;i++){const canvas=document.createElement('canvas');canvas.width=Math.ceil(vb.width);canvas.height=Math.ceil(vb.height);canvases.push(canvas);}
        await a.render({canvas:canvases[0],canvasContext:canvases[0].getContext('2d'),viewport:va,transform:[1,0,0,1,left,top],background:'white',intent:'print'}).promise;
        await b.render({canvas:canvases[1],canvasContext:canvases[1].getContext('2d'),viewport:vb,background:'white',intent:'print'}).promise;
        const x=Math.ceil(left)+2,y=Math.ceil(top)+2,w=Math.floor(piece.largura*72/25.4*2)-4,h=Math.floor(piece.altura*72/25.4*2)-4,images=canvases.map(c=>c.getContext('2d').getImageData(x,y,w,h));
        let above2=0,max=0;for(let i=0;i<images[0].data.length;i++){const d=Math.abs(images[0].data[i]-images[1].data[i]);if(d>2)above2++;max=Math.max(max,d);}
        const comparison=document.createElement('canvas');canvases.push(comparison);comparison.width=w*2+12;comparison.height=h;const ctx=comparison.getContext('2d');ctx.fillStyle='white';ctx.fillRect(0,0,comparison.width,comparison.height);ctx.putImageData(images[0],0,0);ctx.putImageData(images[1],w+12,0);
        return {components:images[0].data.length,above2,max,fraction:above2/images[0].data.length,png:comparison.toDataURL('image/png').split(',')[1]};
    }finally{for(const c of canvases)c.width=c.height=0;await Promise.all(tasks.map(t=>t.destroy()));}
}
