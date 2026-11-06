// Diagnóstico e encaixe em milímetros; não altera o plano nem o PDF original.
import {limiteEscalaCompacta} from './pdf-imposition-compact.js';
const TOL=.001;
const rectangle=(x,y,width,height)=>({x,y,width,height});
export function limitesMontagem(plan,index){
    const rects=plan.folhas[index].map(p=>rectangle(p.x-plan.sangria,p.y-plan.sangria,p.largura+2*plan.sangria,p.altura+2*plan.sangria));
    for(const line of [...plan.folhas[index].flatMap(p=>p.cortes),...(plan.cortesGrupo[index]||[])]){
        const r=plan.marcas.espessura/2;
        rects.push(rectangle(Math.min(line.x1,line.x2)-r,Math.min(line.y1,line.y2)-r,Math.abs(line.x2-line.x1)+2*r,Math.abs(line.y2-line.y1)+2*r));
    }
    return rects;
}
export function bordasUltrapassadas(rects,frame){
    const edges=new Set();for(const r of rects){if(r.x<frame.x-TOL)edges.add('esquerda');if(r.x+r.width>frame.x+frame.width+TOL)edges.add('direita');if(r.y<frame.y-TOL)edges.add('inferior');if(r.y+r.height>frame.y+frame.height+TOL)edges.add('superior');}return [...edges];
}
export function areaImprimivelConhecida(plan,source){
    const g=source?.geometry||source;
    if(!g||![g.width,g.height,g.printableWidth,g.printableHeight].every(n=>Number.isFinite(n)&&n>0)||![g.left,g.top].every(n=>Number.isFinite(n)&&n>=0)||g.left+g.printableWidth>g.width+.1||g.top+g.printableHeight>g.height+.1||Math.abs(g.width-plan.folha.largura)>.1||Math.abs(g.height-plan.folha.altura)>.1)return null;
    return {...rectangle(g.left,g.height-g.top-g.printableHeight,g.printableWidth,g.printableHeight),printer:source?.printer||'última consulta'};
}
export function analisarMontagem(plan,printableSource=null){
    const paper=rectangle(0,0,plan.folha.largura,plan.folha.altura),margin=rectangle(plan.util.x,plan.util.y,plan.util.largura,plan.util.altura),printable=areaImprimivelConhecida(plan,printableSource);
    const sheets=plan.folhas.map((pieces,i)=>{
        const rects=limitesMontagem(plan,i),foraFolha=bordasUltrapassadas(rects,paper),foraMargens=bordasUltrapassadas(rects,margin),foraImprimivel=printable?bordasUltrapassadas(rects,printable):null;
        const areaTotalArtes=pieces.reduce((sum,p)=>sum+p.largura*p.altura,0);
        const areaArtesNaFolha=pieces.reduce((sum,p)=>sum+Math.max(0,Math.min(p.x+p.largura,paper.width)-Math.max(p.x,0))*Math.max(0,Math.min(p.y+p.altura,paper.height)-Math.max(p.y,0)),0);
        return {foraFolha,foraMargens,foraImprimivel,areaTotalArtes,areaArtesNaFolha,areaVazia:Math.max(0,paper.width*paper.height-areaArtesNaFolha),aproveitamento:100*areaArtesNaFolha/(paper.width*paper.height),rects};
    });
    const all=sheets.flatMap(s=>s.rects),pad=Math.max(8,Math.min(paper.width,paper.height)*.04);
    const bounds=all.reduce((b,r)=>({left:Math.min(b.left,r.x),bottom:Math.min(b.bottom,r.y),right:Math.max(b.right,r.x+r.width),top:Math.max(b.top,r.y+r.height)}),{left:0,bottom:0,right:paper.width,top:paper.height});
    const left=bounds.left-pad,bottom=bounds.bottom-pad,right=bounds.right+pad,top=bounds.top+pad;
    // Metadados da prévia interna, nunca tamanho do PDF de produção.
    const palco=rectangle(left,bottom,right-left,top-bottom);
    if(palco.width>10000||palco.height>10000)throw Error('A área de trabalho da prévia excede 10.000 mm. Reduza a grade ou as medidas.');
    return {paper,margin,printable,sheets,palco,requerConfirmacao:sheets.some(s=>s.foraFolha.length||s.foraMargens.length||s.foraImprimivel?.length)};
}
export function sugerirAjusteMontagem(plan,config){
    const limitar=value=>config.modo==='nup'&&config.escalaModo==='reduzir'?Math.min(value,1/Math.max(...plan.folhas.flat().map(p=>p.escala))):value;
    const ext=plan.marcas.ativas?Math.max(plan.sangria,plan.marcas.afastamento+plan.marcas.comprimento+plan.marcas.espessura/2):plan.sangria;
    if(plan.distribuicao==='compacto'){
        const factor=limitar(Math.min(...plan.folhas.map(pieces=>limiteEscalaCompacta(pieces,{colunas:plan.colunas,linhas:plan.linhas,ordem:config.ordem??'horizontal',gapX:Number(config.espacoX)||0,gapY:Number(config.espacoY)||0},plan.util.largura-2*ext,plan.util.altura-2*ext)))),pieces=plan.folhas.flat();
        if(!Number.isFinite(factor)||factor<=0)throw Error('Intervalos, margens e marcas consomem a área útil. Ajuste-os ou aumente a folha.');
        if(pieces.some(p=>plan.sangria>p.sangriaDisponivel*p.escala*factor+.000001))throw Error('O encaixe reduziria a sangria abaixo da solicitada.');
        const total=(config.fatorMontagem??1)*factor;if(total<.000001||total>1000)throw Error('Escala de encaixe além do limite da ferramenta.');
        return {factor,fatorMontagem:total,escalaMin:Math.min(...pieces.map(p=>p.escala))*factor*100,escalaMax:Math.max(...pieces.map(p=>p.escala))*factor*100,config:{...config,automatico:false,colunas:plan.colunas,linhas:plan.linhas,fatorMontagem:total}};
    }
    const group=['grupo','bordas'].includes(plan.marcas.modo),reserve=group?plan.sangria:ext,outer=group?ext-reserve:0;
    const w=(plan.util.largura-2*outer-(plan.colunas-1)*(Number(config.espacoX)||0))/plan.colunas-2*reserve;
    const h=(plan.util.altura-2*outer-(plan.linhas-1)*(Number(config.espacoY)||0))/plan.linhas-2*reserve;
    if(w<=0||h<=0)throw Error('Margens, espaços e marcas já consomem a área útil. Reduza-os ou escolha papel maior.');
    const pieces=plan.folhas.flat();const factor=limitar(Math.min(w/Math.max(...pieces.map(p=>p.largura)),h/Math.max(...pieces.map(p=>p.altura))));
    if(!Number.isFinite(factor)||factor<=0)throw Error('Não há escala proporcional válida para este encaixe.');
    if(pieces.some(p=>plan.sangria>p.sangriaDisponivel*p.escala*factor+.000001))throw Error('O encaixe reduziria a sangria real abaixo da solicitada. Reduza a sangria solicitada ou use uma origem com mais sangria.');
    const total=(config.fatorMontagem??1)*factor;
    if(total<.000001||total>1000)throw Error('Escala de encaixe além do limite da ferramenta.');
    return {factor,fatorMontagem:total,escalaMin:Math.min(...pieces.map(p=>p.escala))*factor*100,escalaMax:Math.max(...pieces.map(p=>p.escala))*factor*100,
        config:{...config,automatico:false,colunas:plan.colunas,linhas:plan.linhas,fatorMontagem:total}};
}
