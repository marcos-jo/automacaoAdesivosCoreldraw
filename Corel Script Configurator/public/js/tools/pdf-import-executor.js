export function executarImportacao(acao, dados, { signal } = {}) {
    return new Promise((resolve, reject) => {
        signal?.throwIfAborted();
        const worker = new Worker(new URL('./pdf-import-worker.js', import.meta.url), { type: 'module' });
        const timeout = setTimeout(() => terminar(reject, new DOMException('Tempo limite de processamento atingido.', 'TimeoutError')), 60_000);
        const cancelar = () => terminar(reject, signal.reason);
        function terminar(settle, value) {
            clearTimeout(timeout);
            signal?.removeEventListener('abort', cancelar);
            worker.terminate();
            settle(value);
        }
        signal?.addEventListener('abort', cancelar, { once: true });
        worker.onmessage = ({ data }) => {
            if (data.erro) terminar(reject, Object.assign(new Error(data.erro.mensagem), { codigo: data.erro.codigo }));
            else terminar(resolve, data.resultado);
        };
        worker.onerror = () => terminar(reject, new Error('Falha ao iniciar ou executar a preparação do PDF.'));
        worker.onmessageerror = () => terminar(reject, new Error('Não foi possível receber o resultado da preparação.'));
        try { worker.postMessage({ acao, ...dados }); }
        catch (error) { terminar(reject, error); }
    });
}
