import {PDFDocument, PDFName, PDFDict, PDFNumber} from '/vendor/pdf-lib/pdf-lib.js';
import {diagnosticarImpressao, prepararParaImpressao, validarActualTextImpressao} from './pdf-print-compatibility.js';
const N=PDFName.of;
const cancel=signal=>{if(signal?.aborted)throw new DOMException('Preparação cancelada','AbortError');};
export class PDFMontagemCompatibilidadeErro extends Error {
    constructor(diagnostico){super('Este PDF contém estruturas especiais que precisam ser preparadas antes da montagem.');this.name='PDFMontagemCompatibilidadeErro';this.diagnostico=diagnostico;}
}
export function diagnosticarMontagem(pdf){
    const base=diagnosticarImpressao(pdf,pdf.getPages().map((_,i)=>i+1));
    const bloqueios=[...base.bloqueios],removiveis=[...base.removiveis];let links=0;
    for(const key of ['AcroForm','OCProperties','OutputIntents','AA','OpenAction']){
        // Destino simples de abertura já é removível no diagnóstico existente.
        if(pdf.catalog.has(N(key))&&!(key==='OpenAction'&&base.removiveis.some(r=>r.codigo==='abertura')))bloqueios.push({codigo:key,descricao:`${key}: aparência, interatividade ou cor ainda sem preparação vetorial homologada para montagem.`});
    }
    for(const [index,page] of pdf.getPages().entries())for(const ref of page.node.Annots()?.asArray()||[]){
        const annotation=pdf.context.lookup(ref,PDFDict),type=annotation.lookup(N('Subtype'))?.toString();
        const flags=annotation.lookupMaybe(N('F'),PDFNumber)?.asNumber()??0;
        if(type==='/Link'&&!(flags&4)&&!annotation.has(N('AP'))&&!annotation.has(N('AA'))){links++;}
        else bloqueios.push({codigo:'anotacao-montagem',descricao:`Página ${index+1}: ${type||'anotação desconhecida'} imprimível, com aparência AP ou não suportada. Não será apagada; prepare sua aparência por um fluxo específico antes de montar.`});
    }
    if(links)removiveis.push({codigo:'links-nao-imprimiveis',descricao:`${links} link(s) sem flag de impressão e sem aparência AP; a interatividade será removida somente na cópia.`});
    const impeditivos=bloqueios.filter(b=>b.codigo!=='actualtext');
    return {...base,bloqueios,removiveis,links,podePreparar:removiveis.length>0&&impeditivos.length===0,necessitaPreparacao:removiveis.length>0||bloqueios.length>0};
}
export async function prepararCopiaMontagem(bytes,{signal,confirmado=false,validar}={}){
    cancel(signal);
    if(!confirmado)throw Error('Confirme explicitamente a preparação da cópia.');
    if(typeof validar!=='function')throw Error('Comparação de texto e aparência obrigatória antes de montar.');
    const pdf=await PDFDocument.load(new Uint8Array(bytes),{updateMetadata:false});
    const diag=diagnosticarMontagem(pdf);
    if(!diag.podePreparar)throw new PDFMontagemCompatibilidadeErro(diag);
    // Mantém a prova por MCID/ActualText já usada no redimensionamento.
    await validarActualTextImpressao(pdf,diag,validar,{signal});
    if(diag.bloqueios.length)throw new PDFMontagemCompatibilidadeErro(diag);
    prepararParaImpressao(pdf,diag);
    for(const page of pdf.getPages()){
        const annots=page.node.Annots();if(!annots)continue;
        // O diagnóstico autorizou somente links sem flag Print, AP ou AA.
        page.node.delete(N('Annots'));
    }
    cancel(signal);const depois=await pdf.save();
    const remaining=diagnosticarMontagem(await PDFDocument.load(depois,{updateMetadata:false}));
    if(remaining.necessitaPreparacao)throw new PDFMontagemCompatibilidadeErro(remaining);
    const proof=await validar({antes:new Uint8Array(bytes),depois,substitutos:diag.actualText,signal});
    cancel(signal);
    if(!proof?.compativel)throw Error(proof?.motivo||'Não foi comprovada a conservação de texto e aparência. Cópia recusada.');
    return {bytes:depois,diagnostico:diag,validacao:proof,removidos:diag.removiveis.map(r=>r.descricao)};
}
