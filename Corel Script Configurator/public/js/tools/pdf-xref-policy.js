// Limitação conhecida da reserialização atual: o contexto pode reter revisões
// com o mesmo número de objeto e gerações distintas. Não escolher uma geração
// nem reparar/renumerar referências aqui. Ler os bytes originais é permitido.
export function diagnosticarSerializacao(pdf) {
    const numeros = new Map();
    for (const [ref] of pdf.context.enumerateIndirectObjects()) {
        if (numeros.has(ref.objectNumber)) return { compativel: false, numero: ref.objectNumber,
            referencias: [numeros.get(ref.objectNumber), ref.toString()] };
        numeros.set(ref.objectNumber, ref.toString());
    }
    return { compativel: true };
}

export function exigirSerializacaoCompativel(pdf, operacao = 'reescrever o PDF') {
    const diagnostico = diagnosticarSerializacao(pdf);
    if (diagnostico.compativel) return;
    const erro = new Error(`Operação bloqueada: ${operacao}. O documento contém gerações concorrentes do objeto ${diagnostico.numero}; o mecanismo atual não garante referências XRef válidas ao reescrever. Abra e visualize o original; nenhuma saída foi publicada.`);
    erro.codigo = 'PDF_XREF_INCOMPATIVEL';
    throw erro;
}
