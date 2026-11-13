// Apresentação somente: nenhum diagnóstico é inferido do texto de uma exceção.
export function mensagemPDF(erro, { uniao = false } = {}) {
    const codigo = erro?.codigo;
    const mensagens = {
        PDF_RESTRITO: ['PDF protegido', 'Este PDF não pode ser unido. Remova-o da lista.'],
        PDF_SENHA: ['PDF com senha', 'A abertura exige senha. Remova-o da lista; este fluxo não recebe senhas.'],
        PDF_CRIPTOGRAFADO: ['PDF protegido', 'Este PDF usa uma proteção ainda não suportada. Remova-o da lista.'],
        PDF_CAMADAS_INCOMPATIVEIS: ['Camadas incompatíveis', 'Este PDF tem camadas incompatíveis. Remova-o da lista.'],
        PDF_INVALIDO: ['PDF inválido ou incompatível', 'Não foi possível ler este arquivo. Solicite outra cópia.'],
        PDF_XREF_INCOMPATIVEL: ['Apenas visualização', uniao ? 'Este PDF não pode ser unido. Remova-o da lista.' : 'Este PDF permite apenas visualização. Edição e exportação bloqueadas.'],
        PDF_RECORTE_ESTRUTURA: ['Recorte bloqueado', 'Este PDF não permite recorte seguro. O arquivo não foi alterado.'],
        PDF_RECORTE_MARGENS: ['Margens inválidas', 'Reduza as margens para manter uma área de recorte.'],
        PDF_ASSINATURA_PROTEGIDA: ['Alterações bloqueadas', erro?.assinaturaDetectada === true ? 'PDF assinado. Alterações bloqueadas.' : 'Possível assinatura digital. Alterações bloqueadas por segurança.']
    };
    const [titulo, mensagem] = mensagens[codigo] || (erro?.name === 'TimeoutError'
        ? ['Preparação demorada', 'O prazo terminou. Remova este arquivo para continuar a seleção.']
        : ['Operação não concluída', 'Consulte os detalhes ou peça ajuda ao suporte.']);
    return { titulo, mensagem, codigo, tecnico: erro?.motivoTecnico || erro?.message || 'Sem diagnóstico adicional.' };
}

export function mostrarMensagemPDF(host, erro, { operacao, uniao = false } = {}) {
    const info = mensagemPDF(erro, { uniao });
    host.replaceChildren();
    host.textContent = `${info.titulo}. ${info.mensagem}`;
    const detalhes = document.createElement('details'), resumo = document.createElement('summary'), texto = document.createElement('pre');
    resumo.textContent = 'Ver detalhes';
    texto.textContent = [info.codigo && `Código: ${info.codigo}`, operacao && `Operação: ${operacao}`, info.tecnico,
        'Para suporte, informe o código e a operação. Não envie documentos de clientes sem autorização.'].filter(Boolean).join('\n');
    texto.style.whiteSpace = 'pre-wrap'; texto.style.overflowWrap = 'anywhere';
    detalhes.append(resumo, texto);
    if (info.codigo) {
        const copiar = document.createElement('button'); copiar.type = 'button'; copiar.textContent = 'Copiar código';
        copiar.addEventListener('click', async () => {
            try {
                if (!globalThis.navigator?.clipboard?.writeText) throw Error('Área de transferência indisponível');
                await globalThis.navigator.clipboard.writeText(info.codigo);
                copiar.textContent = 'Código copiado';
            } catch { copiar.textContent = 'Selecione e copie o código acima'; }
        });
        detalhes.append(copiar);
    }
    host.append(detalhes);
    return detalhes;
}
