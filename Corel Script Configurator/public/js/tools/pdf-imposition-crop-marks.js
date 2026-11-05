// Segmentos em mm. Compartilhados sem alterar as fórmulas dos modos anteriores.
const epsilon=1e-7;
export function marcasIndividuais(x,y,largura,altura,marcas){
    const cortes=[],a=marcas.afastamento,l=marcas.comprimento;
    for(const X of [x,x+largura])for(const Y of [y,y+altura]){
        const sx=X===x?-1:1,sy=Y===y?-1:1;
        cortes.push({x1:X+sx*a,y1:Y,x2:X+sx*(a+l),y2:Y});
        cortes.push({x1:X,y1:Y+sy*a,x2:X,y2:Y+sy*(a+l)});
    }return cortes;
}
export function marcasDoGrupo(pecas,marcas){
    if(!['grupo','bordas'].includes(marcas.modo)||!pecas.length)return [];
    const esquerda=Math.min(...pecas.map(p=>p.x)),direita=Math.max(...pecas.map(p=>p.x+p.largura));
    const inferior=Math.min(...pecas.map(p=>p.y)),superior=Math.max(...pecas.map(p=>p.y+p.altura));
    const linhas=[],a=marcas.afastamento,l=marcas.comprimento;
    if(marcas.modo==='bordas'){
        const unicos=valores=>valores.sort((a,b)=>a-b).filter((v,i,sorted)=>i===0||v-sorted[i-1]>epsilon);
        for(const X of unicos(pecas.flatMap(p=>[p.x,p.x+p.largura]))){linhas.push({x1:X,y1:inferior-a,x2:X,y2:inferior-a-l});linhas.push({x1:X,y1:superior+a,x2:X,y2:superior+a+l});}
        for(const Y of unicos(pecas.flatMap(p=>[p.y,p.y+p.altura]))){linhas.push({x1:esquerda-a,y1:Y,x2:esquerda-a-l,y2:Y});linhas.push({x1:direita+a,y1:Y,x2:direita+a+l,y2:Y});}
        return linhas;
    }
    for(const X of [esquerda,direita])for(const Y of [inferior,superior]){
        const sx=X===esquerda?-1:1,sy=Y===inferior?-1:1;
        linhas.push({x1:X+sx*a,y1:Y,x2:X+sx*(a+l),y2:Y});linhas.push({x1:X,y1:Y+sy*a,x2:X,y2:Y+sy*(a+l)});
    }return linhas;
}
