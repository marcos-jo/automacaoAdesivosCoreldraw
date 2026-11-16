import assert from 'node:assert/strict';

export const CENARIOS_ASSINATURA = Object.freeze(['aplicada', 'multipla', 'perms']);
export const MENSAGEM_GIRO_PROTEGIDO = 'O documento contém assinatura digital ou evidência estrutural compatível. Processamento bloqueado: girar páginas, para preservar sua integridade. Abra e visualize o original. Validade criptográfica não verificada.';

// Observa eventos CDP; nunca intercepta ou silencia o console da aplicação.
export class ConsoleAssinaturas {
    constructor() {
        this.registros = [];
        this.confirmados = [];
        this.cenario = null;
    }

    registrar(method, params) {
        if (method === 'Runtime.exceptionThrown' || method === 'Runtime.consoleAPICalled' && params.type === 'error') {
            this.registros.push({ method, params, cenario: this.cenario, reconhecido: false });
        }
    }

    async recusaGiro(cenario, acao, { cdp, origem }) {
        assert.ok(CENARIOS_ASSINATURA.includes(cenario), 'Cenário de recusa desconhecido');
        assert.equal(this.cenario, null, 'Recusas de console não podem ser concorrentes');
        assert.ok(!this.confirmados.some(r => r.cenario === cenario), 'Cenário de console repetido');
        const antes = await cdp('Runtime.evaluate', { expression: 'Date.now()', returnByValue: true });
        const inicio = this.registros.length;
        this.cenario = cenario;
        try {
            await acao();
            // A ação aguarda a mensagem e a liberação da UI; a barreira drena os
            // eventos anteriores no mesmo canal CDP antes de validar a janela.
            const depois = await cdp('Runtime.evaluate', { expression: 'Date.now()', returnByValue: true });
            const recebidos = this.registros.slice(inicio);
            assert.equal(recebidos.length, 1, `Console ${cenario}: exigido exatamente 1 registro na recusa de giro; recebidos ${recebidos.length}`);
            const registro = recebidos[0], evento = registro.params;
            assert.equal(registro.cenario, cenario);
            assert.equal(registro.method, 'Runtime.consoleAPICalled', 'Exceção não tratada nunca é recusa esperada');
            assert.equal(evento.type, 'error');
            assert.ok(Number.isFinite(evento.timestamp) && evento.timestamp >= antes.result.value && evento.timestamp < depois.result.value + 1, 'Registro fora do momento da ação');
            assert.equal(evento.args.length, 1, 'Recusa deve registrar somente o Error original');
            const erro = evento.args[0];
            assert.equal(erro.subtype, 'error');
            assert.equal(erro.className, 'Error');
            assert.ok(erro.objectId, 'Error CDP sem identidade');
            const origemConsole = evento.stackTrace?.callFrames?.[0];
            assert.equal(origemConsole?.functionName, 'girarPagina');
            assert.equal(origemConsole?.url, `${origem}/js/tools/pdfs.js`);
            assert.ok(erro.description?.startsWith(`Error: ${MENSAGEM_GIRO_PROTEGIDO}\n`), 'Mensagem/operacao de recusa divergente');
            assert.ok(erro.description.includes(`at exigirOperacaoSemAssinatura (${origem}/js/tools/pdf-signatures.js:`), 'Origem da proteção divergente');
            assert.ok(erro.description.includes(`at PDFWorkspace.girarPagina (${origem}/js/tools/pdf-workspace.js:`), 'Operação de workspace divergente');
            const propriedades = await cdp('Runtime.getProperties', { objectId: erro.objectId, ownProperties: true });
            const valor = nome => propriedades.result.find(p => p.name === nome)?.value?.value;
            assert.equal(valor('message'), MENSAGEM_GIRO_PROTEGIDO);
            assert.equal(valor('codigo'), 'PDF_ASSINATURA_PROTEGIDA');
            assert.equal(valor('assinaturaDetectada'), true);
            assert.equal(this.registros.length - inicio, 1, 'Erro adicional durante a validação da recusa');
            registro.reconhecido = true;
            this.confirmados.push({ cenario, operacao: 'girar páginas', codigo: valor('codigo'), mensagem: valor('message'), origem: origemConsole.url, indice: inicio,
                momento: evento.timestamp, inicio: antes.result.value, fim: depois.result.value });
        } finally {
            this.cenario = null;
        }
    }

    verificarFinal(assinaturasExecutadas) {
        assert.equal(this.cenario, null, 'Expectativa de console ainda aberta');
        assert.deepEqual(this.confirmados.map(r => r.cenario), assinaturasExecutadas ? CENARIOS_ASSINATURA : [], 'Quantidade/ordem das três recusas esperadas');
        const inesperados = this.registros.filter(r => !r.reconhecido);
        assert.deepEqual(inesperados, [], 'Erros JavaScript inesperados no navegador');
        return { esperados: this.confirmados, inesperados: inesperados.length };
    }
}
