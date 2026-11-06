import {marcasIndividuais,marcasDoGrupo} from './pdf-imposition-crop-marks.js';
const EPS=1e-7;
const marcasInvalidas=message=>Object.assign(Error(message),{code:'COMPACT_MARKS'});
// Faixas de dimensões reais: linhas no preenchimento horizontal, colunas no vertical.
function medidas(items,{colunas,linhas,ordem='horizontal',gapX,gapY}){
    const vertical=ordem==='vertical',limit=vertical?linhas:colunas,strips=[];
    for(let i=0;i<items.length;i+=limit){const pieces=items.slice(i,i+limit);strips.push({pieces,w:vertical?Math.max(...pieces.map(p=>p.largura)):pieces.reduce((sum,p)=>sum+p.largura,0),h:vertical?pieces.reduce((sum,p)=>sum+p.altura,0):Math.max(...pieces.map(p=>p.altura)),gx:vertical?0:(pieces.length-1)*gapX,gy:vertical?(pieces.length-1)*gapY:0});}
    return {strips,width:vertical?strips.reduce((s,r)=>s+r.w,0):Math.max(...strips.map(r=>r.w)),height:vertical?Math.max(...strips.map(r=>r.h)):strips.reduce((s,r)=>s+r.h,0),vertical,gapX,gapY};
}
export function limiteEscalaCompacta(items,options,width,height){
    const m=medidas(items,options);
    if(m.vertical)return Math.min((width-(m.strips.length-1)*m.gapX)/m.width,...m.strips.map(r=>(height-r.gy)/r.h));
    return Math.min((height-(m.strips.length-1)*m.gapY)/m.height,...m.strips.map(r=>(width-r.gx)/r.w));
}
function tamanho(items,options,scale){const m=medidas(items,options);return m.vertical?{width:m.width*scale+(m.strips.length-1)*m.gapX,height:Math.max(...m.strips.map(r=>r.h*scale+r.gy))}:{width:Math.max(...m.strips.map(r=>r.w*scale+r.gx)),height:m.height*scale+(m.strips.length-1)*m.gapY};}
// Busca exaustiva até 8 peças; acima disso, feixe determinístico de 32 candidatos.
// Agregados e caminho persistente evitam copiar listas inteiras a cada página.
function orientar(items,options,width,height,auto,scale,adjust,ceiling=1){
    if(!auto)return items.map(p=>({...p,giroAutomatico:0}));
    const vertical=options.ordem==='vertical',limit=vertical?options.linhas:options.colunas,beam=items.length<=8?2**items.length:32;
    const sizeA=vertical?height:width,sizeB=vertical?width:height,gapA=vertical?options.gapY:options.gapX,gapB=vertical?options.gapX:options.gapY;
    let states=[{count:0,a:0,b:0,completedB:0,maxA:0,fitA:Infinity,rotations:0,node:null}];
    for(let i=0;i<items.length;i++){
        const next=[];
        for(const state of states)for(const turn of [0,90]){
            const p=items[i],w=turn?p.altura:p.largura,h=turn?p.largura:p.altura;
            let {count,a,b,completedB,maxA,fitA}=state;
            if(count===limit){fitA=Math.min(fitA,(sizeA-(count-1)*gapA)/a);maxA=Math.max(maxA,a*scale+(count-1)*gapA);completedB+=b;count=0;a=0;b=0;}
            count++;a+=vertical?h:w;b=Math.max(b,vertical?w:h);
            const strips=Math.floor(i/limit)+1,totalB=completedB+b,realA=Math.max(maxA,a*scale+(count-1)*gapA),realB=totalB*scale+(strips-1)*gapB;
            const fit=Math.min(ceiling,fitA,(sizeA-(count-1)*gapA)/a,(sizeB-(strips-1)*gapB)/totalB);
            next.push({count,a,b,completedB,maxA,fitA,rotations:state.rotations+(turn?1:0),node:{turn,previous:state.node},score:adjust?-fit:Math.max(realA/sizeA,realB/sizeB),area:realA*realB});
        }
        next.sort((a,b)=>Math.abs(a.score-b.score)>EPS?a.score-b.score:Math.abs(a.area-b.area)>EPS?a.area-b.area:a.rotations-b.rotations);states=next.slice(0,beam);
    }
    const turns=[];for(let n=states[0].node;n;n=n.previous)turns.push(n.turn);turns.reverse();
    return items.map((p,i)=>({...p,largura:turns[i]?p.altura:p.largura,altura:turns[i]?p.largura:p.altura,rotacao:(p.rotacao+turns[i])%360,giroAutomatico:turns[i]}));
}
function validarMarcas(pieces,lines,marcas,bleed){
    if(!lines.length)return;
    const r=marcas.espessura/2;
    if(['grupo','bordas'].includes(marcas.modo)){if(marcas.afastamento<bleed+r-EPS)throw Error('Afaste as marcas externas pelo menos a sangria mais metade do traço, para não invadir a reserva da arte.');return;}
    // Faixas podem ser irregulares. Consultar apenas vizinhos por grade espacial,
    // sem produto cartesiano de 10.000 peças × todas as marcas.
    const sizes=pieces.map(p=>Math.max(p.largura,p.altura)+2*bleed).sort((a,b)=>a-b),step=Math.max(1,sizes[Math.floor(sizes.length/2)],sizes.at(-1)/32),grid=new Map();let checks=0;
    for(const [i,p]of pieces.entries()){const minX=Math.floor((p.x-bleed)/step),maxX=Math.floor((p.x+p.largura+bleed)/step),minY=Math.floor((p.y-bleed)/step),maxY=Math.floor((p.y+p.altura+bleed)/step);for(let x=minX;x<=maxX;x++)for(let y=minY;y<=maxY;y++){const key=x+':'+y;if(!grid.has(key))grid.set(key,[]);grid.get(key).push(i);}}
    for(const line of lines){const x1=Math.min(line.x1,line.x2)-r,x2=Math.max(line.x1,line.x2)+r,y1=Math.min(line.y1,line.y2)-r,y2=Math.max(line.y1,line.y2)+r,near=new Set();for(let x=Math.floor(x1/step);x<=Math.floor(x2/step);x++)for(let y=Math.floor(y1/step);y<=Math.floor(y2/step);y++)for(const i of grid.get(x+':'+y)||[])near.add(i);
        for(const i of near){if(++checks>2000000)throw marcasInvalidas('Validação de marcas individuais muito extensa. Use marcas externas ou reduza as peças por folha.');const p=pieces[i];if(x1<p.x+p.largura+bleed-EPS&&x2>p.x-bleed+EPS&&y1<p.y+p.altura+bleed-EPS&&y2>p.y-bleed+EPS)throw marcasInvalidas('Marcas cruzariam arte ou sangria no compacto. Use Contorno externo + divisões, afaste as marcas ou aumente os intervalos.');}
    }
}
export function planejarNupCompacto({config,paginas,folha,margem,util,gapX,gapY,sangria,marcas,escalaModo,escala,fator,colunas,linhas,permitirExcesso}){
    if(paginas.some(p=>!p))throw Error('N-up compacto não utiliza vagas 0. Selecione páginas reais ou use Em células.');
    const capacidade=colunas*linhas,totalFolhas=Math.ceil(paginas.length/capacidade);
    if(capacidade>10000||totalFolhas>1000)throw Error('A grade ultrapassa o limite de posições/folhas.');
    const options={colunas,linhas,ordem:config.ordem??'horizontal',gapX,gapY},width=util.largura-2*sangria,height=util.altura-2*sangria;
    if(width<=0||height<=0)throw Error('A sangria não deixa área útil para o compacto.');
    const oriented=[];
    const encaixar=['ajustar','reduzir','preencher'].includes(escalaModo),ceiling=escalaModo==='preencher'?Infinity:1;
    for(let i=0;i<paginas.length;i+=capacidade)oriented.push(orientar(paginas.slice(i,i+capacidade),options,width,height,config.rotacaoAutomatica,escala,encaixar,ceiling));
    const automatic=encaixar?Math.min(ceiling,...oriented.map(items=>limiteEscalaCompacta(items,options,width,height))):escala;
    if(!Number.isFinite(automatic)||automatic<=0)throw Error('Os intervalos não deixam espaço para encaixar as peças. Reduza-os ou escolha uma folha maior.');
    const s=escalaModo==='reduzir'?Math.min(1,automatic*fator):automatic*fator,folhas=[];
    for(const [sheet,items]of oriented.entries()){
        const scaled=items.map(p=>({...p,largura:p.largura*s,altura:p.altura*s,escala:s}));
        const m=medidas(scaled,options),box=tamanho(scaled,options,1);
        if(box.width>10000||box.height>10000)throw Error('O bloco compacto excede 10.000 mm. Reduza a grade, os intervalos ou a escala.');
        if(!permitirExcesso&&(box.width>width+EPS||box.height>height+EPS))throw Error('O bloco compacto excede a área útil. Ajuste escala ou papel.');
        const usedCols=m.vertical?m.strips.length:Math.min(colunas,items.length),usedRows=m.vertical?Math.min(linhas,items.length):m.strips.length;
        if(sangria>0&&((usedCols>1&&gapX<2*sangria-EPS)||(usedRows>1&&gapY<2*sangria-EPS)))throw Error('Os intervalos do compacto são entre cortes. Para não sobrepor sangrias, use pelo menos duas vezes a sangria entre peças, ou use sangria zero.');
        const left=util.x+sangria+(config.centralizar===false?0:(width-box.width)/2),top=util.y+util.altura-sangria-(config.centralizar===false?0:(height-box.height)/2);
        const pieces=[];let x=left,y=top,position=sheet*capacidade;
        for(const strip of m.strips){let a=m.vertical?top:left;for(const p of strip.pieces){if(sangria>p.sangriaDisponivel*s+EPS)throw Error(`Página ${p.origem}: a escala reduz a sangria disponível abaixo da solicitada.`);const piece={...p,indice:position++,x:m.vertical?x:a,y:(m.vertical?a:y)-p.altura,cortes:[]};if(marcas.modo==='individuais')piece.cortes=marcasIndividuais(piece.x,piece.y,piece.largura,piece.altura,marcas);pieces.push(piece);a+=m.vertical?-(p.altura+gapY):p.largura+gapX;}if(m.vertical)x+=strip.w+gapX;else y-=strip.h+gapY;}
        folhas.push(pieces);
    }
    const cortesGrupo=folhas.map(pieces=>marcasDoGrupo(pieces,marcas));
    for(const [i,pieces]of folhas.entries())validarMarcas(pieces,[...pieces.flatMap(p=>p.cortes),...cortesGrupo[i]],marcas,sangria);
    if(!permitirExcesso)for(const [i,pieces]of folhas.entries())for(const line of [...pieces.flatMap(p=>p.cortes),...cortesGrupo[i]]){const r=marcas.espessura/2;if(Math.min(line.x1,line.x2)-r<util.x-EPS||Math.max(line.x1,line.x2)+r>util.x+util.largura+EPS||Math.min(line.y1,line.y2)-r<util.y-EPS||Math.max(line.y1,line.y2)+r>util.y+util.altura+EPS)throw Error('Marcas externas excedem a área útil. Pré-visualize o excesso para decidir sobre corte ou ajuste.');}
    return {folha,util,margem,sangria,marcas,colunas,linhas,capacidade,quantidade:paginas.length,totalFolhas,folhas,cortesGrupo,distribuicao:'compacto',escalaAutomatica:automatic,giroAutomaticoPermitido:!!config.rotacaoAutomatica};
}
