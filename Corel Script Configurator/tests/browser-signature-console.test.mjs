import test from 'node:test';
import assert from 'node:assert/strict';
import { ConsoleAssinaturas, CENARIOS_ASSINATURA, MENSAGEM_GIRO_PROTEGIDO } from './helpers/pdf-signature-console.mjs';

const origem = 'http://localhost:3000';
function ambiente({ codigo = 'PDF_ASSINATURA_PROTEGIDA', mensagem = MENSAGEM_GIRO_PROTEGIDO } = {}) {
    const consoleTeste = new ConsoleAssinaturas();
    const evento = {
        type: 'error', timestamp: 1005,
        args: [{ subtype: 'error', className: 'Error', objectId: 'erro-1', description: `Error: ${mensagem}\n    at exigirOperacaoSemAssinatura (${origem}/js/tools/pdf-signatures.js:159:18)\n    at PDFWorkspace.girarPagina (${origem}/js/tools/pdf-workspace.js:245:14)` }],
        stackTrace: { callFrames: [{ functionName: 'girarPagina', url: `${origem}/js/tools/pdfs.js` }] }
    };
    let relogio = 0;
    const cdp = async method => method === 'Runtime.getProperties'
        ? { result: Object.entries({ message: mensagem, codigo, assinaturaDetectada: true }).map(([name, value]) => ({ name, value: { value } })) }
        : { result: { value: relogio++ % 2 ? 1010 : 1000 } };
    const emitir = (params = evento, method = 'Runtime.consoleAPICalled') => consoleTeste.registrar(method, params);
    const recusa = (acao = () => emitir(), cenario = 'aplicada') => consoleTeste.recusaGiro(cenario, acao, { cdp, origem });
    return { consoleTeste, evento, emitir, recusa };
}

test('console: exige e conserva as três recusas por cenário, código, origem e momento', async () => {
    const a = ambiente();
    for (const cenario of CENARIOS_ASSINATURA) await a.recusa(() => a.emitir(), cenario);
    assert.equal(a.consoleTeste.verificarFinal(true).esperados.length, 3);
    assert.equal(a.consoleTeste.registros.length, 3, 'Eventos não são apagados');
});

test('console: duplicação idêntica na mesma recusa reprova', async () => {
    const a = ambiente();
    await assert.rejects(a.recusa(() => { a.emitir(); a.emitir(); }), /exatamente 1 registro/);
});

test('console: erro adicional no mesmo cenário reprova, mesmo com a recusa correta', async () => {
    const a = ambiente();
    await assert.rejects(a.recusa(() => { a.emitir(); a.emitir({ ...a.evento, args: [{ value: 'falha inesperada' }] }); }), /exatamente 1 registro/);
});

test('console: exceção não tratada nunca corresponde a console.error esperado', async () => {
    const a = ambiente();
    await assert.rejects(a.recusa(() => a.emitir({ exceptionDetails: { text: MENSAGEM_GIRO_PROTEGIDO } }, 'Runtime.exceptionThrown')), /Exceção não tratada/);
});

test('console: mensagem correta com código errado reprova', async () => {
    const a = ambiente({ codigo: 'OUTRO_ERRO' });
    await assert.rejects(a.recusa(), /PDF_ASSINATURA_PROTEGIDA/);
});

test('console: mesmo código com operação/mensagem diferente reprova', async () => {
    const a = ambiente({ mensagem: 'Processamento bloqueado: outra operação.' });
    await assert.rejects(a.recusa(), /Mensagem\/operacao/);
});

test('console: origem diferente na mesma janela não é permitida', async () => {
    const a = ambiente();
    a.evento.stackTrace.callFrames[0].url = `${origem}/js/outro.js`;
    await assert.rejects(a.recusa(), /pdfs\.js/);
});

test('console: mensagem de proteção fora da ação continua inesperada', () => {
    const a = ambiente(); a.emitir();
    assert.throws(() => a.consoleTeste.verificarFinal(false), /Erros JavaScript inesperados/);
});

test('console: recusa sem registro esperado reprova', async () => {
    const a = ambiente();
    await assert.rejects(a.recusa(() => {}), /recebidos 0/);
});

test('console: quarto registro depois das três recusas continua fatal', async () => {
    const a = ambiente();
    for (const cenario of CENARIOS_ASSINATURA) await a.recusa(() => a.emitir(), cenario);
    a.emitir();
    assert.throws(() => a.consoleTeste.verificarFinal(true), /Erros JavaScript inesperados/);
});

test('console: timestamp anterior à ação reprova, mesmo com texto/código/origem corretos', async () => {
    const a = ambiente(); a.evento.timestamp = 999;
    await assert.rejects(a.recusa(), /fora do momento/);
});

test('console: texto correto sem o frame da proteção reprova', async () => {
    const a = ambiente(); a.evento.args[0].description = `Error: ${MENSAGEM_GIRO_PROTEGIDO}\n at outraFuncao`;
    await assert.rejects(a.recusa(), /Origem da proteção/);
});

test('console: conjunto incompleto ou cenário repetido não fecha a validação', async () => {
    const a = ambiente(); await a.recusa();
    assert.throws(() => a.consoleTeste.verificarFinal(true), /Quantidade\/ordem/);
    await assert.rejects(a.recusa(), /Cenário de console repetido/);
});
