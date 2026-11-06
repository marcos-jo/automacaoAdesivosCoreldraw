// Somente apresentação: recebe o diagnóstico pronto, sem recalcular geometria ou liberar saída.
const tipos=[
    ['foraFolha','Possível corte fora da folha física','possível corte fora do papel'],
    ['foraMargens','Margens excedidas','margens excedidas'],
    ['foraImprimivel','Fora da área imprimível consultada','fora da área imprimível']
];
export function intervalosFolhas(numbers){
    const values=[...new Set(numbers)].sort((a,b)=>a-b),ranges=[];
    for(let i=0;i<values.length;i++){const first=values[i];let last=first;while(values[i+1]===last+1)last=values[++i];ranges.push(first===last?String(first):`${first}–${last}`);}
    return ranges.join(', ');
}
export function agruparAvisosMontagem(geometry){
    const sheets=geometry?.sheets||[],affected=new Set(),groups=[];
    for(const [key,title,short]of tipos){
        const occurrences=[],patterns=new Map();
        for(const [i,sheet]of sheets.entries()){
            const edges=[...new Set(sheet[key]||[])];if(!edges.length)continue;
            const number=i+1;affected.add(number);occurrences.push({folha:number,bordas:edges});
            const signature=edges.slice().sort().join('|');if(!patterns.has(signature))patterns.set(signature,{bordas:edges,folhas:[]});patterns.get(signature).folhas.push(number);
        }
        if(occurrences.length)groups.push({tipo:key,titulo:title,resumo:short,ocorrencias:occurrences,padroes:[...patterns.values()]});
    }
    return {total:sheets.length,afetadas:affected.size,grupos:groups,resumo:affected.size?`${affected.size} de ${sheets.length} folha(s) com avisos: ${groups.map(g=>g.resumo).join('; ')}.`:''};
}
export function textoAvisosMontagem(model){
    return model.grupos.map(group=>`${group.titulo}: `+group.padroes.map(p=>`folhas ${intervalosFolhas(p.folhas)} — bordas ${p.bordas.join(', ')}`).join('; ')+'.').join('\n');
}
let sequence=0;
export class AvisosMontagem {
    constructor(get){
        this.get=get;this.button=get('avisos-ver');this.panel=get('avisos-detalhes');
        this.panel.id=`montagem-avisos-${++sequence}`;this.button.setAttribute('aria-controls',this.panel.id);
        this.panel.setAttribute('aria-label','Detalhes dos avisos da montagem');
        this.button.onclick=()=>this.expandir(this.panel.hidden);
        get('avisos-recolher').onclick=()=>this.expandir(false,true);
        this.panel.addEventListener('keydown',e=>{if(e.key==='Escape'){e.preventDefault();e.stopPropagation();this.expandir(false,true);}});
    }
    expandir(open,returnFocus=false){
        this.panel.hidden=!open;this.button.setAttribute('aria-expanded',String(open));this.button.textContent=open?'Ocultar detalhes':'Ver detalhes';
        if(open){this.panel.focus({preventScroll:true});this.panel.scrollIntoView({block:'nearest'});}else if(returnFocus)this.button.focus();
    }
    limpar(){this.expandir(false);this.get('avisos-lista').replaceChildren();}
    atualizar(model){
        if(!model.afetadas){this.limpar();return;}
        const doc=this.panel.ownerDocument,content=[];
        for(const group of model.grupos){
            const block=doc.createElement('div'),title=doc.createElement('h4'),list=doc.createElement('ul');
            title.textContent=`${group.titulo} · ${group.ocorrencias.length} folha(s)`;block.append(title,list);
            for(const pattern of group.padroes){const row=doc.createElement('li');row.textContent=`Folhas ${intervalosFolhas(pattern.folhas)} — bordas: ${pattern.bordas.join(', ')}.`;list.append(row);}
            content.push(block);
        }
        this.get('avisos-lista').replaceChildren(...content);
    }
}
