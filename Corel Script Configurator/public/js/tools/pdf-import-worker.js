import { PDFDocument } from '/vendor/pdf-lib/pdf-lib.js';
import { PDFCombinacao } from './pdf-combine-session.js';
import { diagnosticarAssinaturas } from './pdf-signatures.js';

self.onmessage = async ({ data }) => {
    const draft = new PDFCombinacao();
    try {
        if (data.acao === 'preparar') {
            const { bytes, nome } = data;
            await draft.add([{ name: nome, size: bytes.length, arrayBuffer: async () => bytes.buffer }]);
            const [source] = draft.sources;
            const previewBytes = source.type === 'PDF' ? undefined : source.previewBytes;
            self.postMessage({ resultado: {
                previewBytes,
                paginas: draft.pages.map(({ index, width, height }) => ({ index, width, height })),
                diagnostico: diagnosticarAssinaturas(source.pdf, { tamanhoArquivo: source.size })
            } }, previewBytes ? [previewBytes.buffer] : []);
        } else if (data.acao === 'combinar') {
            for (const { id, bytes } of data.fontes) {
                draft.sources.push({ id, pdf: await PDFDocument.load(bytes, { updateMetadata: false }) });
            }
            const sources = new Map(draft.sources.map(source => [source.id, source]));
            draft.pages = data.paginas.map(({ origem, index }) => ({ source: sources.get(origem), index }));
            const resultado = await draft.build();
            self.postMessage({ resultado }, [resultado.buffer]);
        } else throw new Error('Operação de importação inválida.');
    } catch (error) {
        self.postMessage({ erro: { codigo: error.codigo, mensagem: error.cause?.codigo ? error.cause.message : error.message } });
    }
};
