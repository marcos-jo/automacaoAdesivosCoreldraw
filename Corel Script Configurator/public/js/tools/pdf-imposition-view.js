import {mmParaPt} from './pdf-imposition-layout.js';

// Sobreposições apenas de inspeção. O PDF de produção não recebe guias/hachuras.
export function desenharInspecaoMontagem(canvas,viewport,result,number,{guias=true}={}){
    const ctx=canvas.getContext('2d'),p=result.plan,g=result.geometria;
    const stage=g?.palco||{x:0,y:0,width:p.folha.largura,height:p.folha.altura};
    const point=(x,y)=>viewport.convertToViewportPoint?viewport.convertToViewportPoint(mmParaPt(x),mmParaPt(y)):[(x-stage.x)/stage.width*canvas.width,(stage.y+stage.height-y)/stage.height*canvas.height];
    const rect=r=>{const a=point(r.x,r.y+r.height),b=point(r.x+r.width,r.y);return [a[0],a[1],b[0]-a[0],b[1]-a[1]];};
    ctx.save();
    if(g){
        const [x,y,w,h]=rect(g.paper);
        ctx.beginPath();ctx.rect(0,0,canvas.width,canvas.height);ctx.rect(x,y,w,h);ctx.clip('evenodd');
        ctx.fillStyle='rgba(110,120,135,.15)';ctx.fillRect(0,0,canvas.width,canvas.height);
        ctx.strokeStyle='rgba(200,20,35,.35)';ctx.lineWidth=1;
        for(let i=-canvas.height;i<canvas.width;i+=14){ctx.beginPath();ctx.moveTo(i,0);ctx.lineTo(i+canvas.height,canvas.height);ctx.stroke();}
        ctx.restore();ctx.save();
        ctx.strokeStyle=g.sheets[number-1]?.foraFolha.length?'#d22235':'#263445';ctx.lineWidth=2;ctx.strokeRect(x,y,w,h);
        ctx.font='12px sans-serif';ctx.fillStyle='#263445';ctx.fillText(`${p.folha.largura} × ${p.folha.altura} mm`,x+4,Math.max(13,y-5));
        ctx.setLineDash([6,4]);ctx.lineWidth=1;ctx.strokeStyle='#a87900';ctx.strokeRect(...rect(g.margin));
        if(g.printable){ctx.strokeStyle='#8a31a1';ctx.setLineDash([2,3]);ctx.strokeRect(...rect(g.printable));}
    }
    if(guias){ctx.strokeStyle='#0077cc';ctx.lineWidth=1;ctx.setLineDash([4,4]);for(const piece of p.folhas[number-1])ctx.strokeRect(...rect({x:piece.x,y:piece.y,width:piece.largura,height:piece.altura}));}
    ctx.restore();
}

// Mostra a folha mesmo durante geração/erro; nunca apresenta arte antiga como nova.
export function desenharFolhaVazia(canvas,config){
    const w=config.largura,h=config.altura;if(!Number.isFinite(w)||!Number.isFinite(h)||w<10||h<10||w>5000||h>5000)return;
    const pad=Math.max(8,Math.min(w,h)*.04),stage={x:-pad,y:-pad,width:w+2*pad,height:h+2*pad};
    const scale=Math.min(480/stage.width,400/stage.height);canvas.width=Math.ceil(stage.width*scale);canvas.height=Math.ceil(stage.height*scale);
    const ctx=canvas.getContext('2d');ctx.fillStyle='white';ctx.fillRect(0,0,canvas.width,canvas.height);
    const m=config.margens||{};
    desenharInspecaoMontagem(canvas,{}, {plan:{folha:{largura:w,altura:h},folhas:[[]]},geometria:{palco:stage,paper:{x:0,y:0,width:w,height:h},margin:{x:m.esquerda||0,y:m.inferior||0,width:w-(m.esquerda||0)-(m.direita||0),height:h-(m.superior||0)-(m.inferior||0)},sheets:[{foraFolha:[]}]}},1,{guias:false});
}
