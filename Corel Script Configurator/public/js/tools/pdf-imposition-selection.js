// Seleção da janela: validar antes de gerar, sem alterar ordem ou inventar páginas.
const invalid=message=>Object.assign(new Error(message),{code:'MONTAGEM_SELECAO'});
export function selecionarPaginasMontagem({modo,selecao='todas',texto='',total}){
    if(!Number.isSafeInteger(total)||total<1)throw invalid('Não foi possível determinar as páginas do documento. Reabra a montagem.');
    if(modo==='nup'&&selecao==='todas'){
        if(total>10000)throw invalid('O documento excede 10.000 páginas. Escolha um intervalo menor em Configurações avançadas.');
        return Array.from({length:total},(_,i)=>i+1);
    }
    if(modo==='nup'&&selecao!=='especificas')throw invalid('Escolha todas as páginas ou uma seleção específica.');
    if(!texto.trim())throw invalid('Informe as páginas desejadas ou selecione Todas as páginas do documento.');
    const paginas=[];
    for(const item of texto.split(',')){
        const m=item.trim().match(/^(\d+)(?:\s*-\s*(\d+))?$/);
        if(!m)throw invalid('Seleção inválida. Use páginas e intervalos como 1,3-5, com números inteiros positivos.');
        const a=Number(m[1]),b=Number(m[2]??m[1]);
        if(!Number.isSafeInteger(a)||!Number.isSafeInteger(b)||a<1||b<1)throw invalid('As páginas começam em 1. Zero e números negativos não são páginas válidas.');
        if(b<a)throw invalid(`Intervalo invertido: ${a}-${b}. Informe a página inicial antes da final.`);
        if(a>total||b>total)throw invalid(`Página ${Math.max(a,b)} inexistente. O documento tem ${total} página(s); use números de 1 a ${total}.`);
        if(paginas.length+b-a+1>10000)throw invalid('Selecione no máximo 10.000 páginas/posições por montagem.');
        for(let p=a;p<=b;p++)paginas.push(p);
    }
    if(modo==='repeat'&&paginas.length!==1)throw invalid('Step & Repeat exige exatamente uma arte. Escolha uma página de origem.');
    return paginas;
}
