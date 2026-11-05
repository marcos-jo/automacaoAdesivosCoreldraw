import {marcasDoGrupo} from './pdf-imposition-crop-marks.js';
const EPS=1e-7;
export const formatosLivreto={A5:{largura:148,altura:210},A4:{largura:210,altura:297}};
const papeis={A4:{largura:297,altura:210},A3:{largura:420,altura:297}};
function numero(v,label,min,max){const n=typeof v==='string'?Number(v.trim().replace(',','.')):v;if(!Number.isFinite(n)||n<min||n>max)throw Error(`${label}: valor inválido (${min} a ${max}).`);return n;}
export function paginarLivreto(count){
    if(!Number.isInteger(count)||count<1||count>2000)throw Error('Livreto: informe de 1 a 2.000 páginas de origem.');
    const total=Math.ceil(count/4)*4,slot=n=>({logica:n,origem:n<=count?n:null}),folhas=[];
    for(let k=0;k<total/4;k++)folhas.push({numero:k+1,frente:[slot(total-2*k),slot(1+2*k)],verso:[slot(2+2*k),slot(total-1-2*k)]});
    // Simula os dois lados das folhas dobradas/aninhadas; é a ordem usada pela prévia de leitura.
    const leitura=[...folhas.flatMap(f=>[f.frente[1],f.verso[0]]),...folhas.slice().reverse().flatMap(f=>[f.verso[1],f.frente[0]])];
    const aberturas=[[null,leitura[0]]];for(let i=1;i<leitura.length-1;i+=2)aberturas.push([leitura[i],leitura[i+1]]);aberturas.push([leitura.at(-1),null]);
    return {originais:count,total,adicionadas:total-count,brancas:Array.from({length:total-count},(_,i)=>count+i+1),folhas,leitura,aberturas};
}
export function girarRetangulo(r,width,height,quarters){let p={...r},w=width,h=height;for(let i=0;i<quarters;i++){p={...p,x:p.y,y:w-p.x-p.width,width:p.height,height:p.width};[w,h]=[h,w];}return p;}
const dentro=(a,b)=>a.x>=b.x-EPS&&a.y>=b.y-EPS&&a.x+a.width<=b.x+b.width+EPS&&a.y+a.height<=b.y+b.height+EPS;
export function planejarLivreto(config,fontes){
    const final=formatosLivreto[config.final??'A5'],canonical=papeis[config.papel??'A4'];if(!final||!canonical)throw Error('Escolha os formatos disponíveis: final A5/A4 e papel A4/A3.');
    const orientation=config.orientacao??'paisagem';if(!['paisagem','retrato'].includes(orientation))throw Error('Orientação de folha inválida.');
    const reverse=config.verso??'normal';if(!['normal','girar180'].includes(reverse))throw Error('Orientação de verso inválida.');
    const mode=config.escalaModo??'real';if(!['real','ajustar','personalizada'].includes(mode))throw Error('Método de escala inválido.');
    const manual=mode==='personalizada'?numero(config.escala,'Escala',1,1000)/100:1;
    const turn=numero(config.rotacao??0,'Giro do conteúdo',0,270);if(turn%90)throw Error('Giro do conteúdo deve ser ortogonal.');
    const margin=Object.fromEntries(['esquerda','direita','superior','inferior'].map(k=>[k,numero(config.margensConteudo?.[k]??0,`Margem ${k}`,0,100)]));
    const avail={width:final.largura-margin.esquerda-margin.direita,height:final.altura-margin.superior-margin.inferior};if(avail.width<=0||avail.height<=0)throw Error('As margens não deixam área para o conteúdo.');
    const bleed=numero(config.sangria??0,'Sangria',0,50),marks={modo:config.marcasCorte?'grupo':'sem',ativas:!!config.marcasCorte||!!config.marcasDobra,afastamento:numero(config.afastamento??3,'Afastamento',0,50),comprimento:numero(config.comprimento??3,'Comprimento',.1,50),espessura:numero(config.espessura??.1,'Traço',.01,2)};
    if(marks.ativas&&marks.afastamento<bleed+marks.espessura/2-EPS)throw Error('Afaste as marcas para fora da sangria, incluindo metade do traço.');
    const dx=numero(config.registroX??0,'Registro do verso X',-20,20),dy=numero(config.registroY??0,'Registro do verso Y',-20,20),pagination=paginarLivreto(fontes.length);
    const q=orientation==='retrato'?1:0,paper={largura:q?canonical.altura:canonical.largura,altura:q?canonical.largura:canonical.altura};
    const fold=canonical.largura/2,frameY=(canonical.altura-final.altura)/2,canonicalFrames=[{x:fold-final.largura,y:frameY,width:final.largura,height:final.altura},{x:fold,y:frameY,width:final.largura,height:final.altura}];
    const issues=new Set(),faces=[],sheets=[],frames=[],groupMarks=[],readingByLogical=new Map();
    for(const f of canonicalFrames)if(!dentro(f,{x:0,y:0,width:canonical.largura,height:canonical.altura}))issues.add('O formato final não cabe no papel dobrado. Selecione papel maior; não será reduzido o formato final silenciosamente.');
    const sourcePlacements=fontes.map((f,i)=>{
        if(!f)throw Error('Livreto exige todas as páginas reais na ordem de leitura.');
        const w=numero(f.largura,`Página ${i+1}: largura`,.01,5000),h=numero(f.altura,`Página ${i+1}: altura`,.01,5000),sw=turn%180?h:w,sh=turn%180?w:h;
        const s=mode==='ajustar'?Math.min(avail.width/sw,avail.height/sh):manual;
        if(bleed>(f.sangriaDisponivel??0)*s+EPS)throw Error(`Página ${i+1}: sangria solicitada não disponível após escala.`);
        if(sw*s>avail.width+EPS||sh*s>avail.height+EPS)issues.add(`Página ${i+1}: conteúdo excede a área do formato final; escolha ajuste proporcional, reduza a escala ou as margens.`);
        return {origem:f.id??i+1,logica:i+1,largura:sw*s,altura:sh*s,escala:s,rotacao:turn,sangriaDisponivel:f.sangriaDisponivel??0,x:margin.esquerda+(avail.width-sw*s)/2,y:margin.inferior+(avail.height-sh*s)/2};
    });
    for(const sheet of pagination.folhas)for(const side of ['frente','verso']){
        const extra=side==='verso'&&reverse==='girar180'?2:0,quarters=(q+extra)%4,transform=r=>girarRetangulo(r,canonical.largura,canonical.altura,quarters),physicalFrames=canonicalFrames.map(transform),pieces=[];
        const face={folha:sheet.numero,lado:side,pdf:faces.length+1,giro:quarters*90,slots:sheet[side].map((slot,i)=>({...slot,moldura:physicalFrames[i]}))};
        for(const [i,slot]of sheet[side].entries()){
            if(!slot.origem){readingByLogical.set(slot.logica,null);continue;}
            const p=sourcePlacements[slot.origem-1],f=canonicalFrames[i],rect=transform({x:f.x+p.x,y:f.y+p.y,width:p.largura,height:p.altura});if(side==='verso'){rect.x+=dx;rect.y+=dy;}
            if(!dentro(rect,physicalFrames[i]))issues.add(`Página ${slot.logica}: posição/registro excede o corte final; o conteúdo não pode ser descartado.`);
            const clip=transform({x:i?fold:f.x-bleed,y:f.y-bleed,width:final.largura+bleed,height:final.altura+2*bleed});
            pieces.push({...p,x:rect.x,y:rect.y,largura:rect.width,altura:rect.height,rotacao:(turn+quarters*90)%360,indice:slot.logica-1,cortes:[],clip});
            const restored=girarRetangulo(rect,paper.largura,paper.altura,(4-quarters)%4);
            readingByLogical.set(slot.logica,{...p,x:restored.x-f.x,y:restored.y-f.y,indice:slot.logica-1,cortes:[],clip:{x:0,y:0,width:final.largura,height:final.altura}});
        }
        const cut=marks.modo==='grupo'?marcasDoGrupo(canonicalFrames.map(f=>({x:f.x,y:f.y,largura:f.width,altura:f.height})),marks):[];
        if(config.marcasDobra){const a=marks.afastamento,l=marks.comprimento;cut.push({x1:fold,y1:frameY-a,x2:fold,y2:frameY-a-l,dobra:true},{x1:fold,y1:frameY+final.altura+a,x2:fold,y2:frameY+final.altura+a+l,dobra:true});}
        const rotatePoint=(x,y)=>{const r=transform({x,y,width:0,height:0});return {x:r.x,y:r.y};};
        groupMarks.push(cut.map(line=>{const a=rotatePoint(line.x1,line.y1),b=rotatePoint(line.x2,line.y2);return {...line,x1:a.x,y1:a.y,x2:b.x,y2:b.y};}));
        faces.push(face);sheets.push(pieces);frames.push(physicalFrames);
    }
    const leitura={folha:{largura:2*final.largura,altura:final.altura},folhas:[],molduras:[],cortesGrupo:[],sangria:bleed,totalFolhas:pagination.aberturas.length};
    for(const opening of pagination.aberturas){const pieces=[];for(const [i,slot]of opening.entries()){const p=slot&&readingByLogical.get(slot.logica);if(p)pieces.push({...p,x:p.x+i*final.largura,clip:{x:i*final.largura,y:0,width:final.largura,height:final.altura}});}leitura.folhas.push(pieces);leitura.molduras.push([{x:0,y:0,width:final.largura,height:final.altura},{x:final.largura,y:0,width:final.largura,height:final.altura}]);leitura.cortesGrupo.push([]);}
    return {tipo:'livreto',folha:paper,util:{x:0,y:0,largura:paper.largura,altura:paper.altura},margem:{esquerda:0,direita:0,superior:0,inferior:0},sangria:bleed,marcas:marks,colunas:q?1:2,linhas:q?2:1,capacidade:2,quantidade:pagination.total,totalFolhas:faces.length,folhasFisicas:pagination.folhas.length,folhas:sheets,cortesGrupo:groupMarks,molduras:frames,faces,paginacao:pagination,leitura,final,margensConteudo:margin,bloqueios:[...issues],verso:reverse,orientacao:orientation,aposDobra:{largura:canonical.largura/2,altura:canonical.altura}};
}
