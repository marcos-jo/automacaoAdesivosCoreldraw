import test from 'node:test';
import assert from 'node:assert/strict';
import { setImmediate as tick } from 'node:timers/promises';
import { Elemento } from './helpers/pdf-dom.mjs';
import { mensagemPDF, mostrarMensagemPDF } from '../public/js/tools/pdf-feedback.js';
import { PDFDocument, PDFName } from '../public/vendor/pdf-lib/pdf-lib.js';
import { PDFCombinacao } from './helpers/pdf-combine-modules.mjs';
import { PDFWorkspace } from './helpers/pdf-resize-modules.mjs';

for (const [codigo, esperado] of [
    ['PDF_RESTRITO', 'Este PDF não pode ser unido. Remova-o da lista.'],
    ['PDF_CAMADAS_INCOMPATIVEIS', 'Este PDF tem camadas incompatíveis. Remova-o da lista.'],
    ['PDF_XREF_INCOMPATIVEL', 'Este PDF permite apenas visualização. Edição e exportação bloqueadas.'],
    ['PDF_RECORTE_ESTRUTURA', 'Este PDF não permite recorte seguro. O arquivo não foi alterado.'],
    ['PDF_SENHA', 'A abertura exige senha. Remova-o da lista; este fluxo não recebe senhas.'],
    ['PDF_CRIPTOGRAFADO', 'Este PDF usa uma proteção ainda não suportada. Remova-o da lista.']
]) test(`mensagem por código: ${codigo}`, () => {
    const r = mensagemPDF({ codigo, message: 'diagnóstico técnico original' });
    assert.equal(r.mensagem, esperado); assert.equal(r.codigo, codigo);
    assert.equal(r.tecnico, 'diagnóstico técnico original');
});

test('assinatura aplicada e inconclusiva permanecem distintas', () => {
    assert.equal(mensagemPDF({ codigo: 'PDF_ASSINATURA_PROTEGIDA', assinaturaDetectada: true }).mensagem, 'PDF assinado. Alterações bloqueadas.');
    assert.match(mensagemPDF({ codigo: 'PDF_ASSINATURA_PROTEGIDA', assinaturaDetectada: false }).mensagem, /^Possível assinatura digital/);
    assert.match(mensagemPDF({ codigo: 'PDF_ASSINATURA_PROTEGIDA' }).mensagem, /^Possível assinatura digital/);
});

test('não classifica exceções por texto nem divulga caminhos na mensagem principal', () => {
    const r = mensagemPDF(Error('XRef assinatura StructTreeRoot PDF_RESTRITO C:\\cliente\\privado.pdf'));
    assert.equal(r.codigo, undefined); assert.equal(r.titulo, 'Operação não concluída');
    assert.doesNotMatch(r.mensagem, /cliente|XRef|assinatura|restri/i);
    assert.match(r.tecnico, /privado.pdf/);
});

test('detalhes recolhidos, cópia somente do código e recolhimento não mudam estado', async () => {
    const anterior = globalThis.document, navigatorOriginal = Object.getOwnPropertyDescriptor(globalThis, 'navigator');
    const copiados = [];
    globalThis.document = { createElement: tag => new Elemento(tag) };
    Object.defineProperty(globalThis, 'navigator', { configurable: true, value: { clipboard: { writeText: async texto => { copiados.push(texto); } } } });
    try {
        const host = new Elemento('div');
        const detalhes = mostrarMensagemPDF(host, { codigo: 'PDF_XREF_INCOMPATIVEL', message: 'Gerações concorrentes do objeto 3' }, { operacao: 'Exportar PDF' });
        assert.ok(!detalhes.open); assert.equal(detalhes.querySelector('summary').textContent, 'Ver detalhes');
        detalhes.open = true;
        assert.match(detalhes.querySelector('pre').textContent, /Código: PDF_XREF_INCOMPATIVEL/);
        assert.match(detalhes.querySelector('pre').textContent, /Operação: Exportar PDF/);
        detalhes.querySelector('button').click(); await tick();
        assert.deepEqual(copiados, ['PDF_XREF_INCOMPATIVEL']);
        assert.equal(detalhes.querySelector('button').textContent, 'Código copiado');
        detalhes.open = false;
        mostrarMensagemPDF(host, Error('Erro sem código'));
        assert.equal(host.querySelectorAll('details').length, 1);
        assert.equal(host.querySelector('button'), undefined);
    } finally {
        globalThis.document = anterior;
        if (navigatorOriginal) Object.defineProperty(globalThis, 'navigator', navigatorOriginal); else delete globalThis.navigator;
    }
});

test('código de camadas vem do catálogo e não do nome; união comum continua disponível', async () => {
    const pdf = await PDFDocument.create(); pdf.addPage([200, 300]);
    const comum = await pdf.save(), draft = new PDFCombinacao();
    await draft.add([{ name: 'camadas-PDF_RESTRITO.pdf', arrayBuffer: async () => comum.slice().buffer }]);
    assert.equal(draft.pages.length, 1);
    pdf.catalog.set(PDFName.of('OCProperties'), pdf.context.obj({}));
    const bloqueado = await pdf.save();
    await assert.rejects(draft.add([{ name: 'original.pdf', arrayBuffer: async () => bloqueado.slice().buffer }]), e => {
        assert.equal(e.codigo, 'PDF_CAMADAS_INCOMPATIVEIS'); assert.equal(e.arquivo, 'original.pdf');
        assert.doesNotMatch(e.message, /original.pdf/); return true;
    });
    assert.equal(draft.pages.length, 1);
});

test('recorte marcado incompatível mantém a recusa e acrescenta somente código de diagnóstico', async () => {
    const pdf = await PDFDocument.create(); pdf.addPage([200, 300]);
    pdf.catalog.set(PDFName.of('StructTreeRoot'), pdf.context.obj({}));
    const bytes = await pdf.save(), ws = new PDFWorkspace();
    await ws.abrir({ name: 'marcado.pdf', arrayBuffer: async () => bytes.slice().buffer });
    const estado = ws.capturarEstado();
    await assert.rejects(ws.prepararRecorte(1, { x: 10, y: 10, width: 100, height: 150 }), { codigo: 'PDF_RECORTE_ESTRUTURA' });
    assert.deepEqual(ws.capturarEstado(), estado);
});

test('falha da área de transferência preserva diagnóstico e oferece cópia manual', async () => {
    const anterior = globalThis.document, navigatorOriginal = Object.getOwnPropertyDescriptor(globalThis, 'navigator');
    globalThis.document = { createElement: tag => new Elemento(tag) };
    Object.defineProperty(globalThis, 'navigator', { configurable: true, value: {} });
    try {
        const host = new Elemento('div'); mostrarMensagemPDF(host, { codigo: 'PDF_RESTRITO', message: 'Restrição confirmada' });
        host.querySelector('button').click(); await tick();
        assert.equal(host.querySelector('button').textContent, 'Selecione e copie o código acima');
        assert.match(host.querySelector('pre').textContent, /Restrição confirmada/);
    } finally {
        globalThis.document = anterior;
        if (navigatorOriginal) Object.defineProperty(globalThis, 'navigator', navigatorOriginal); else delete globalThis.navigator;
    }
});
