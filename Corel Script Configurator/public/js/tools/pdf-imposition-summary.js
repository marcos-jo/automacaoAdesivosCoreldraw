const fmt=n=>Number.isFinite(n)?n.toLocaleString('pt-BR',{maximumFractionDigits:2}):'não informado';
const percent=n=>Number.isFinite(n)?n.toLocaleString('pt-BR',{maximumFractionDigits:4})+'%':'não informada';
const dims=(w,h)=>`${fmt(w)} × ${fmt(h)} mm`;
const range=values=>{const min=Math.min(...values),max=Math.max(...values);return Math.abs(max-min)<1e-8?percent(min):`${percent(min)} a ${percent(max)}`;};
export function resumirMontagem(result){
    const p=result.plan,c=result.config||{},factor=c.fatorMontagem??1,rows=[],keys=new Set(),origins=new Map(result.origens?.map(o=>[o.id,o]));
    for(const piece of p.folhas.flat()){
        const key=[piece.origem,piece.largura,piece.altura,piece.escala,piece.rotacao].join(':');if(keys.has(key))continue;keys.add(key);
        const origin=origins.get(piece.origem),turned=piece.rotacao%180;
        const cutW=origin?.largura??(turned?piece.altura:piece.largura)/piece.escala,cutH=origin?.altura??(turned?piece.largura:piece.altura)/piece.escala;
        rows.push({pagina:piece.origem,origem:dims(origin?.larguraPagina??cutW,origin?.alturaPagina??cutH),corte:dims(cutW,cutH),final:dims(piece.largura,piece.altura),escala:piece.escala*100,encaixe:piece.escala/factor*100,orientacao:Math.abs(piece.largura-piece.altura)<1e-7?'quadrada':piece.largura>piece.altura?'paisagem':'retrato',giro:piece.rotacao??0});
    }
    if(!rows.length)return {basico:'Medidas indisponíveis.',detalhes:'',confirmacao:'',efetiva:'não informada'};
    const effective=range(rows.map(r=>r.escala)),hasSize=c.modo==='repeat'&&(c.pecaLargura!=null||c.pecaAltura!=null),auto=c.modo==='nup'&&['ajustar','reduzir','preencher'].includes(c.escalaModo);
    const politica=c.modo==='nup'?({real:'Manter tamanho real (100%)',reduzir:'Reduzir somente se necessário (máximo 100%)',preencher:'Preencher células — ampliação autorizável',ajustar:'Encaixe legado explícito',personalizada:'Escala personalizada explícita'}[c.escalaModo||'real']):null;
    const ampliacao=c.modo==='nup'&&rows.some(r=>r.escala>100.000001);
    const manual=hasSize?'não aplicada; tamanho final informado':auto?'não aplicada':c.escalaModo==='personalizada'?percent(c.escala):'100% (tamanho real)';
    const automatic=auto?range(rows.map(r=>r.encaixe)):'não aplicado';
    const additional=Math.abs(factor-1)>1e-10?percent(factor*100)+' sobre a escala-base':'nenhum';
    const unique=key=>{const values=[...new Set(rows.map(r=>r[key]))];return values.length===1?values[0]:`${values.length} tamanhos diferentes (ver por página)`;};
    const orientations=[...new Set(rows.map(r=>r.orientacao))].join('/'),turns=[...new Set(rows.map(r=>r.giro))].map(n=>n+'°').join('/');
    const mixed=new Set(rows.map(r=>r.final)).size>1;
    const basico=(politica?`Política de escala: ${politica}.\n`:'')+`Origem: ${unique('origem')} · corte usado: ${unique('corte')}\nPeça final: ${unique('final')} · escala efetiva: ${effective}\nOrientação final: ${orientations} · giro na montagem: ${turns}\nManual: ${manual} · encaixe automático: ${automatic}\nAjuste para caber: ${additional}. Escalas relativas ao corte de origem; zoom e driver não incluídos.`+(ampliacao?'\nAMPLIAÇÃO INTENCIONAL: confira as medidas finais antes de gerar ou imprimir.':'')+(p.distribuicao==='compacto'?'\nCompacto: encaixe das artes sem reservar espaço para marcas externas. Confira os avisos de corte.'+(mixed?' Tamanhos mistos podem deixar vazios e exigir refiles separados; divisões projetadas não garantem corte contínuo pela folha.':''):'' );
    const details=rows.slice(0,100).map(r=>`Página ${r.pagina} — origem ${r.origem}; corte ${r.corte}; peça ${r.final}; efetiva ${percent(r.escala)}; ${r.orientacao}; giro na montagem ${r.giro}°.`).join('\n')+(rows.length>100?`\nMais ${rows.length-100} combinações de página/escala. Confira as demais folhas na prévia.`:'');
    const pieces=p.folhas.reduce((sum,sheet)=>sum+sheet.length,0);
    const virada=p.viradaDuplex==='longa'||p.viradaDuplex==='curta'?p.viradaDuplex:null;
    const producao=p.repetirPorPagina
        ? `${p.capacidade} peça(s) por face · ${p.totalFolhas} página(s) no PDF${virada?` · ${p.totalFolhas/2} folha(s) por cópia em duplex, borda ${virada}`:''}. Tiragem em Cópias na impressão.`
        : '';
    const warnings=result.geometria?.sheets.flatMap((s,i)=>[s.foraFolha.length?`Folha ${i+1}: fora do papel (${s.foraFolha.join(', ')}).`:'',s.foraMargens.length?`Folha ${i+1}: fora das margens (${s.foraMargens.join(', ')}).`:'',s.foraImprimivel?.length?`Folha ${i+1}: fora da área imprimível consultada (${s.foraImprimivel.join(', ')}).`:''].filter(Boolean))||[];
    const contagem=producao||`${pieces} peça(s) preenchida(s) em ${p.totalFolhas} folha(s).${c.modo==='nup'?' Não há repetição automática da página para preencher vagas.':''}`;
    const confirmation=`${basico}\nFolha: ${dims(p.folha.largura,p.folha.altura)}. Grade: ${p.colunas} × ${p.linhas}, capacidade ${p.capacidade}. Distribuição: ${p.distribuicao==='compacto'?'compacta em faixas':'células/grade'}.\n${contagem}\n${details}\n${warnings.slice(0,12).join('\n')}${warnings.length>12?'\nHá avisos nas demais folhas; confira a prévia.':''}\nO PDF já estará imposto. Confirme papel e escala do driver (normalmente 100%), sem N-up/livreto adicional.${virada?` Configure também o duplex pela borda ${virada} no driver.`:''}`;
    return {basico,detalhes:details,confirmacao:confirmation,producao,efetiva:effective,rows,manual,automatica:automatic,adicional:additional,politica,ampliacao};
}
