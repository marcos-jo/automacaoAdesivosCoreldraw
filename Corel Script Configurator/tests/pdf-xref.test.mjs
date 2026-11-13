import test from 'node:test';
import assert from 'node:assert/strict';
import { PDFDocument } from '../public/vendor/pdf-lib/pdf-lib.js';
import * as pdfjs from '../public/vendor/pdfjs/pdf.js';
import { PDFWorkspace } from './helpers/pdf-resize-modules.mjs';
import { criarPDFGeracoes } from './helpers/pdf-xref-fixture.mjs';
import { PDFCombinacao } from './helpers/pdf-combine-modules.mjs';
import { motorUrl, impressaoUrl, copiaImpressaoUrl, assinaturasUrl } from './helpers/pdf-resize-modules.mjs';
import { readFile } from 'node:fs/promises';
import { criarPDFAssinado } from './helpers/pdf-signature-fixture.mjs';
let codigoLegado = await readFile(new URL('../public/js/tools/pdf-editor.js', import.meta.url), 'utf8');
for (const [de, para] of Object.entries({ '/vendor/pdf-lib/pdf-lib.js': new URL('../public/vendor/pdf-lib/pdf-lib.js', import.meta.url).href, './pdf-signatures.js': assinaturasUrl, './pdf-xref-policy.js': new URL('../public/js/tools/pdf-xref-policy.js', import.meta.url).href })) codigoLegado = codigoLegado.replaceAll(JSON.stringify(de), JSON.stringify(para));
const legado = await import('data:text/javascript;base64,' + Buffer.from(codigoLegado).toString('base64'));
pdfjs.GlobalWorkerOptions.workerSrc = new URL('../public/vendor/pdfjs/pdf.worker.js', import.meta.url).href;
const arquivo = bytes => ({ name: 'geracoes.pdf', arrayBuffer: async () => bytes.slice().buffer });
async function ler(bytes) { const task = pdfjs.getDocument({ data: bytes.slice() }); try { const d = await task.promise; const p = await d.getPage(1); await p.getOperatorList(); return d.numPages; } finally { await task.destroy(); } }
async function abrir(bytes = criarPDFGeracoes()) { const ws = new PDFWorkspace(); await ws.abrir(arquivo(bytes)); return ws; }

function suspender() {
    let liberar;
    const espera = new Promise(resolve => { liberar = resolve; });
    return { espera, liberar };
}

for (const perfil of ['comum', 'assinado', 'xref', 'mesmo arquivo']) {
    for (const falhaAntiga of [false, true]) test(`abertura concorrente: ${perfil}, ${falhaAntiga ? 'erro' : 'sucesso'} antigo`, async () => {
        const comum = await PDFDocument.create(); comum.addPage([111, 222]);
        const bytes = perfil === 'assinado' ? await criarPDFAssinado() : perfil === 'xref' ? criarPDFGeracoes() : await comum.save();
        const ws = new PDFWorkspace(), pausa = suspender(), novo = arquivo(bytes);
        const antigo = perfil === 'mesmo arquivo' ? novo : arquivo(await comum.save());
        let leituras = 0;
        const lerArquivo = antigo.arrayBuffer;
        antigo.arrayBuffer = async () => {
            if (++leituras === 1) { await pausa.espera; if (falhaAntiga) throw Error('Leitura antiga falhou'); }
            return lerArquivo();
        };
        const pendente = ws.abrir(antigo);
        const recusada = assert.rejects(pendente, { codigo: 'PDF_ABERTURA_OBSOLETA' });
        await ws.abrir(novo);
        const estado = ws.capturarEstado(), paginas = ws.obterPaginas(), resumo = ws.obterResumoAssinaturas();
        const geracao = ws.geracaoAbertura;
        pausa.liberar(); await recusada;
        assert.deepEqual(ws.capturarEstado(), estado);
        assert.deepEqual(ws.obterPaginas(), paginas);
        assert.deepEqual(ws.obterResumoAssinaturas(), resumo);
        assert.equal(ws.geracaoAbertura, geracao);
        if (perfil === 'xref' || perfil === 'assinado') {
            assert.deepEqual(await ws.obterBytesVisualizacao(), bytes);
            await assert.rejects(ws.obterBytes());
        } else {
            ws.girarPagina(1, 90); assert.equal(await ler(await ws.obterBytes()), 1);
        }
    });
}

test('abertura concorrente: tentativa recente falha e não ressuscita abertura anterior', async () => {
    const ws = await abrir(), estado = ws.capturarEstado(), pausa = suspender();
    const pendente = ws.abrir({ name: 'antiga.pdf', arrayBuffer: async () => { await pausa.espera; return criarPDFGeracoes().buffer; } });
    const recusada = assert.rejects(pendente, { codigo: 'PDF_ABERTURA_OBSOLETA' });
    await assert.rejects(ws.abrir({ name: 'recente-invalida.pdf', arrayBuffer: async () => new ArrayBuffer(0) }));
    assert.deepEqual(ws.capturarEstado(), estado);
    pausa.liberar(); await recusada;
    assert.deepEqual(ws.capturarEstado(), estado);
    assert.deepEqual(await ws.obterBytesVisualizacao(), estado.bytesOriginais);
});

for (const falha of [false, true]) test(`abertura concorrente: ${falha ? 'erro' : 'sucesso'} tardio no load`, async () => {
    const ws = await abrir(), pausa = suspender(), iniciou = suspender();
    const dadosAntigos = criarPDFGeracoes({ concorrentes: false }).buffer;
    const carregar = PDFDocument.load;
    PDFDocument.load = async function (dados, ...args) {
        const pdf = await carregar.call(this, dados, ...args);
        if (dados === dadosAntigos) {
            iniciou.liberar(); await pausa.espera;
            if (falha) throw Error('Load antigo falhou');
        }
        return pdf;
    };
    try {
        const pendente = ws.abrir({ name: 'antigo.pdf', arrayBuffer: async () => dadosAntigos });
        const recusada = assert.rejects(pendente, { codigo: 'PDF_ABERTURA_OBSOLETA' });
        await iniciou.espera; await ws.abrir(arquivo(criarPDFGeracoes()));
        const estado = ws.capturarEstado(); pausa.liberar(); await recusada;
        assert.deepEqual(ws.capturarEstado(), estado);
        assert.deepEqual(await ws.obterBytesVisualizacao(), criarPDFGeracoes());
    } finally { pausa.liberar(); PDFDocument.load = carregar; }
});

test('abertura concorrente: nova tentativa invalida união pendente mesmo quando falha', async () => {
    const pdf = await PDFDocument.create(); pdf.addPage([100, 100]);
    const bytes = await pdf.save(), ws = await abrir(bytes), estado = ws.capturarEstado(), pausa = suspender();
    const uniao = ws.adicionarArquivos([{ name: 'unir.pdf', arrayBuffer: async () => { await pausa.espera; return bytes.slice().buffer; } }]);
    const recusada = assert.rejects(uniao, { codigo: 'PDF_UNIAO_OBSOLETA' });
    await assert.rejects(ws.abrir({ name: 'invalido.pdf', arrayBuffer: async () => new ArrayBuffer(0) }));
    pausa.liberar(); await recusada;
    assert.deepEqual(ws.capturarEstado(), estado);
});

test('abertura concorrente: união publicada invalida abertura antiga', async () => {
    const pdf = await PDFDocument.create(); pdf.addPage([100, 100]);
    const bytes = await pdf.save(), ws = await abrir(bytes), pausa = suspender();
    const pendente = ws.abrir({ name: 'antiga.pdf', arrayBuffer: async () => { await pausa.espera; return criarPDFGeracoes().buffer; } });
    const recusada = assert.rejects(pendente, { codigo: 'PDF_ABERTURA_OBSOLETA' });
    await ws.adicionarArquivos([arquivo(bytes)]);
    const estado = ws.capturarEstado(); pausa.liberar(); await recusada;
    assert.deepEqual(ws.capturarEstado(), estado); assert.equal(ws.getTotalPaginas(), 2);
});

test('XRef: entrada literal abre no PDF.js; save direto da biblioteca gera saída recusada', async () => {
    const bytes = criarPDFGeracoes(); assert.equal(await ler(bytes), 2);
    const pdf = await PDFDocument.load(bytes, { updateMetadata: false });
    await assert.rejects(ler(await pdf.save()), /Inconsistent generation in XRef|Invalid Root reference/);
});
test('XRef: visualização preserva bytes originais sem chamar save', async () => {
    const bytes = criarPDFGeracoes(), original = bytes.slice(), ws = await abrir(bytes); let saves = 0;
    const save = ws.pdf.save.bind(ws.pdf); ws.pdf.save = (...args) => { saves++; return save(...args); };
    const vista = await ws.obterBytesVisualizacao(); assert.deepEqual(vista, original); assert.equal(saves, 0);
    vista[0] ^= 1; assert.deepEqual(await ws.obterBytesVisualizacao(), original); assert.equal(await ler(await ws.obterBytesVisualizacao()), 2); assert.deepEqual(bytes, original);
});
for (const [nome, acao] of [
    ['giro', ws => ws.girarPagina(1, 90)], ['exclusão', ws => ws.excluirPaginas([1])],
    ['ordem', ws => ws.reordenarPaginas([2, 1])], ['duplicação', ws => ws.duplicarPaginas([1])],
    ['recorte', ws => ws.prepararRecorte(1, { x: 10, y: 10, width: 100, height: 50 })],
    ['redimensionamento', ws => ws.prepararRedimensionamento({})], ['exportação', ws => ws.obterBytes()]
]) test(`XRef: ${nome} incompatível recusa antes de alterar ou salvar`, async () => {
    const ws = await abrir(), estado = ws.capturarEstado(), paginas = ws.pdf.getPages().map(p => p.node.toString()); let saves = 0;
    ws.pdf.save = async () => { saves++; throw Error('Não deveria salvar'); };
    await assert.rejects(Promise.resolve().then(() => acao(ws)), e => e.codigo === 'PDF_XREF_INCOMPATIVEL');
    assert.equal(saves, 0); assert.deepEqual(ws.capturarEstado(), estado); assert.deepEqual(ws.pdf.getPages().map(p => p.node.toString()), paginas);
});
test('XRef: origem incompatível não publica união', async () => {
    const pdf = await PDFDocument.create(); pdf.addPage([100, 100]); const ws = await abrir(await pdf.save()), estado = ws.capturarEstado();
    await assert.rejects(ws.adicionarArquivos([arquivo(criarPDFGeracoes())]), e => e.codigo === 'PDF_XREF_INCOMPATIVEL');
    assert.deepEqual(ws.capturarEstado(), estado);
});
test('XRef: geração não zero isolada continua editável e derivado reabre no PDF.js', async () => {
    const ws = await abrir(criarPDFGeracoes({ concorrentes: false })); ws.girarPagina(1, 90);
    const bytes = await ws.obterBytes(); assert.equal(await ler(bytes), 2);
    assert.equal((await PDFDocument.load(bytes)).getPage(0).getRotation().angle, 90);
});
test('XRef: troca e rollback não reutilizam bytes de outro documento/revisão', async () => {
    const bytes = criarPDFGeracoes(), ws = await abrir(bytes), estado = ws.capturarEstado();
    const comum = await PDFDocument.create(); comum.addPage([100, 200]); await ws.abrir(arquivo(await comum.save()));
    assert.equal(await ler(await ws.obterBytesVisualizacao()), 1);
    ws.restaurarEstado(estado); assert.deepEqual(await ws.obterBytesVisualizacao(), bytes);
    ws.registrarAlteracao(); await assert.rejects(ws.obterBytesVisualizacao(), e => e.codigo === 'PDF_XREF_INCOMPATIVEL');
});
test('XRef: rascunho de união recusa origem, inclusive build interno usado pelo worker', async () => {
    const bytes = criarPDFGeracoes(), draft = new PDFCombinacao();
    await assert.rejects(draft.add([arquivo(bytes)]), e => e.codigo === 'PDF_XREF_INCOMPATIVEL'); assert.equal(draft.sources.length, 0);
    draft.sources.push({ pdf: await PDFDocument.load(bytes), id: 1 });
    await assert.rejects(draft.build(), e => e.codigo === 'PDF_XREF_INCOMPATIVEL');
});
test('XRef: exportação legada recusa antes de criar download ou extrair', async () => {
    const documento = { getData: async () => criarPDFGeracoes() };
    await assert.rejects(legado.salvarPdfGirado(documento, new Map(), 'teste.pdf'), e => e.codigo === 'PDF_XREF_INCOMPATIVEL');
    await assert.rejects(legado.salvarPaginasSelecionadas(documento, [1], 'extrair', new Map(), 'teste.pdf'), e => e.codigo === 'PDF_XREF_INCOMPATIVEL');
});
test('XRef: motores internos de preparo/resize recusam antes de mutar a origem', async () => {
    const pdf = await PDFDocument.load(criarPDFGeracoes()), before = pdf.getPages().map(p => p.node.toString());
    const { redimensionarDocumento } = await import(motorUrl), { prepararParaImpressao, validarActualTextImpressao } = await import(impressaoUrl), { reconstruirCopiaImpressao } = await import(copiaImpressaoUrl);
    await assert.rejects(redimensionarDocumento(pdf, {}), e => e.codigo === 'PDF_XREF_INCOMPATIVEL');
    assert.throws(() => prepararParaImpressao(pdf, {}), e => e.codigo === 'PDF_XREF_INCOMPATIVEL');
    await assert.rejects(validarActualTextImpressao(pdf, {}, null), e => e.codigo === 'PDF_XREF_INCOMPATIVEL');
    await assert.rejects(reconstruirCopiaImpressao(pdf, {}), e => e.codigo === 'PDF_XREF_INCOMPATIVEL');
    assert.deepEqual(pdf.getPages().map(p => p.node.toString()), before);
});
