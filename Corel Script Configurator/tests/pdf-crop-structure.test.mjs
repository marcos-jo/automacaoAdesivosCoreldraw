import test from 'node:test';
import assert from 'node:assert/strict';
import { PDFDocument, PDFName, PDFRef, PDFString, PDFHexString, decodePDFRawStream } from '../public/vendor/pdf-lib/pdf-lib.js';
import { PDFWorkspace, estruturaRecorteUrl } from './helpers/pdf-resize-modules.mjs';
import { criarEstruturaRecorte } from './helpers/pdf-crop-structure-fixture.mjs';
const { validarEstruturaRecorte, conferirEstruturaRecorte } = await import(estruturaRecorteUrl);
const N = PDFName.of, caixa = { x: 20, y: 30, width: 650, height: 450 };
const arquivo = bytes => ({ name: 'marcado.pdf', arrayBuffer: async () => bytes.slice().buffer });
async function abrir(f) { const bytes = await f.pdf.save(), ws = new PDFWorkspace(); await ws.abrir(arquivo(bytes)); return { bytes, ws }; }

test('estrutura válida: preparar, cancelar, aplicar, exportar e reabrir conserva manifesto, texto e Tabs', async () => {
    const f = await criarEstruturaRecorte(), antes = validarEstruturaRecorte(f.pdf), { bytes, ws } = await abrir(f), original = bytes.slice(), estado = ws.capturarEstado();
    assert.equal(antes.elementos, 1); assert.equal(antes.mcids, 1);
    const abort = new AbortController(), cancelada = await ws.prepararRecorte(1, caixa, { signal: abort.signal });
    abort.abort(); assert.throws(() => ws.aplicarRecorte(cancelada), /cancelada/); assert.deepEqual(ws.capturarEstado(), estado);
    await assert.rejects(ws.prepararRecorte(1, caixa, { signal: abort.signal }), /cancelada/);
    const prep = await ws.prepararRecorte(1, caixa); assert.deepEqual(ws.capturarEstado(), estado); ws.aplicarRecorte(prep);
    const saida = await ws.obterBytes(), depois = await PDFDocument.load(saida, { updateMetadata: false });
    conferirEstruturaRecorte(antes, depois); assert.deepEqual(depois.getPage(0).getMediaBox(), caixa); assert.deepEqual(depois.getPage(0).getCropBox(), caixa);
    assert.equal(depois.getPage(0).node.lookup(N('Tabs')).toString(), '/S');
    assert.equal(depois.context.lookup(f.elemRef).lookup(N('ActualText')).decodeText(), 'Texto logico');
    assert.equal(depois.context.lookup(f.elemRef).lookup(N('Alt')).decodeText(), 'Descricao');
    assert.equal(depois.catalog.lookup(N('Lang')).decodeText(), 'pt-BR'); assert.deepEqual(bytes, original);
    const reaberta = new PDFWorkspace(); await reaberta.abrir(arquivo(saida)); await reaberta.prepararRecorte(1, { x: 30, y: 40, width: 600, height: 400 });
});

const invalidos = {
    raizDireta: f => f.pdf.catalog.set(N('StructTreeRoot'), f.root),
    raizAusente: f => f.pdf.catalog.set(N('StructTreeRoot'), PDFRef.of(99999)),
    paiAusente: f => f.elem.set(N('P'), PDFRef.of(99999)),
    paiIncorreto: f => f.elem.set(N('P'), f.pagina.ref),
    paginaAusente: f => f.elem.set(N('Pg'), PDFRef.of(99999)),
    paginaIncorreta: f => f.elem.set(N('Pg'), f.rootRef),
    semPagina: f => f.elem.delete(N('Pg')),
    chavePagina: f => f.pagina.node.set(N('StructParents'), f.c.obj(1)),
    indiceFracionario: f => f.elem.set(N('K'), f.c.obj(0.5)),
    indiceNegativo: f => f.elem.set(N('K'), f.c.obj(-1)),
    indiceGrande: f => f.elem.set(N('K'), f.c.obj(10000)),
    donoIncorreto: f => f.pais.set(0, f.rootRef),
    donoAusente: f => f.pais.set(0, PDFRef.of(99999)),
    parentTreeVazia: f => f.pt.set(N('Nums'), f.c.obj([])),
    parentTreeKids: f => f.pt.set(N('Kids'), f.c.obj([])),
    parentTreeDuplicada: f => f.pt.set(N('Nums'), f.c.obj([0, f.pais, 0, f.pais])),
    parentTreeIndice: f => f.pt.set(N('Nums'), f.c.obj([1, f.pais])),
    parentTreeExtra: f => f.pais.push(f.elemRef),
    proximaChave: f => f.root.set(N('ParentTreeNextKey'), f.c.obj(0)),
    mcidSemArvore: f => f.pagina.node.set(N('Contents'), f.c.register(f.c.flateStream('/P <</MCID 1>> BDC EMC'))),
    arvoreSemMcid: f => f.pagina.node.set(N('Contents'), f.c.register(f.c.flateStream('q Q'))),
    mcidDuplicado: f => f.pagina.node.set(N('Contents'), f.c.register(f.c.flateStream('/P <</MCID 0>> BDC EMC /P <</MCID 0>> BDC EMC'))),
    propriedadeDuplicada: f => f.pagina.node.set(N('Contents'), f.c.register(f.c.flateStream('/P <</MCID 0 /MCID 0>> BDC EMC'))),
    propriedadeNomeada: f => f.pagina.node.set(N('Contents'), f.c.register(f.c.flateStream('/P /Props BDC EMC'))),
    inlineImage: f => f.pagina.node.set(N('Contents'), f.c.register(f.c.flateStream('BI ID /P <</MCID 0>> BDC EMC EI'))),
    mcidSomenteString: f => f.pagina.node.set(N('Contents'), f.c.register(f.c.flateStream('(/P <</MCID 0>> BDC EMC) Tj'))),
    mcidSomenteComentario: f => f.pagina.node.set(N('Contents'), f.c.register(f.c.flateStream('% /P <</MCID 0>> BDC EMC\nq Q'))),
    emcSolto: f => f.pagina.node.set(N('Contents'), f.c.register(f.c.flateStream('/P <</MCID 0>> BDC EMC EMC'))),
    semEmc: f => f.pagina.node.set(N('Contents'), f.c.register(f.c.flateStream('/P <</MCID 0>> BDC'))),
    elementoDuplicado: f => f.root.set(N('K'), f.c.obj([f.elemRef, f.elemRef])),
    cicloElemento: f => f.elem.set(N('K'), f.c.obj([f.elemRef])),
    cicloArray: f => { const a = f.c.obj([]), r = f.c.register(a); a.push(r); f.elem.set(N('K'), a); },
    objetoReferenciado: f => f.elem.set(N('K'), f.c.obj({ Type: 'OBJR', Obj: f.pagina.ref })),
    streamReferenciado: f => f.elem.set(N('K'), f.c.obj({ Type: 'MCR', Pg: f.pagina.ref, MCID: 0, Stm: f.pagina.node.get(N('Contents')) })),
    atributoDesconhecido: f => f.elem.set(N('A'), f.c.obj({ O: 'Layout', BBox: [0, 0, 1, 1] })),
    roleMapInvalido: f => f.root.set(N('RoleMap'), f.c.obj({ Textbox: 'Desconhecido' })),
    textoInvalido: f => f.elem.set(N('ActualText'), f.c.obj(1)),
    marcacaoFalsa: f => f.pdf.catalog.set(N('MarkInfo'), f.c.obj({ Marked: false })),
    multiplasPaginas: f => f.pdf.addPage([100, 100]),
    formXObject: f => f.c.register(f.c.flateStream('q Q', { Type: 'XObject', Subtype: 'Form', BBox: [0, 0, 1, 1] })),
    vinculoExtra: f => f.c.register(f.c.obj({ StructParent: 1 })),
    referenciaEmRecurso: f => f.recursos.set(N('Desconhecido'), PDFRef.of(99999)),
    profundidade: f => { let a = f.c.obj([0]); for (let i = 0; i < 70; i++) a = f.c.obj([a]); f.elem.set(N('K'), a); },
    excessoObjetos: f => { for (let i = 0; i < 20001; i++) f.c.register(f.c.obj(i)); }
};
for (const [nome, alterar] of Object.entries(invalidos)) test(`estrutura recusada sem reparo: ${nome}`, async () => {
    const f = await criarEstruturaRecorte(); alterar(f); assert.throws(() => validarEstruturaRecorte(f.pdf), /Recorte bloqueado/);
});

for (const nome of ['AcroForm', 'OCProperties', 'Perms', 'AA']) test(`marcado mantém bloqueio ${nome} sem alterar workspace`, async () => {
    const f = await criarEstruturaRecorte(); f.pdf.catalog.set(N(nome), f.c.obj({})); const { ws } = await abrir(f), estado = ws.capturarEstado();
    await assert.rejects(ws.prepararRecorte(1, caixa), /bloqueado|assinatura/i); assert.deepEqual(ws.capturarEstado(), estado);
});
test('marcado mantém bloqueio de anotações e assinatura DocMDP indireta', async () => {
    for (const assinatura of [false, true]) {
        const f = await criarEstruturaRecorte();
        if (assinatura) f.pdf.catalog.set(N('Perms'), f.c.obj({ DocMDP: f.c.register(f.c.obj({ Type: 'Sig', ByteRange: [0, 10, 20, 10], Contents: PDFHexString.of('010203') })) }));
        else f.pagina.node.set(N('Annots'), f.c.obj([f.c.register(f.c.obj({ Type: 'Annot', Subtype: 'Link', Rect: [0, 0, 10, 10] }))]));
        const { ws } = await abrir(f), estado = ws.capturarEstado(); await assert.rejects(ws.prepararRecorte(1, caixa), /bloqueado|assinatura/i); assert.deepEqual(ws.capturarEstado(), estado);
    }
});
test('prévia vencida de marcado não é adotada e sangria ainda exige aceite', async () => {
    const f = await criarEstruturaRecorte(); f.pagina.setTrimBox(0, 0, 700, 500); const { ws } = await abrir(f), prep = await ws.prepararRecorte(1, caixa);
    assert.throws(() => ws.aplicarRecorte(prep), /sangria/); ws.girarPagina(1, 90); assert.throws(() => ws.aplicarRecorte(prep, { aceitarPerdaSangria: true }), /vencida/);
});
test('manifesto recusa alteração semântica de estrutura, recursos e bytes', async () => {
    for (const alterar of [f => f.elem.set(N('ActualText'), PDFString.of('alterado')), f => f.recursos.set(N('Novo'), f.c.obj(1)), f => { f.stream.contents[0] ^= 1; }]) {
        const f = await criarEstruturaRecorte(), a = validarEstruturaRecorte(f.pdf); alterar(f); assert.throws(() => conferirEstruturaRecorte(a, f.pdf));
    }
});

test('Lang local e MCID escapado no nome são preservados; strings e comentários não contam como marcação', async () => {
    const f = await criarEstruturaRecorte();
    f.pagina.node.set(N('Contents'), f.c.register(f.c.flateStream('% /MCID 999 BDC\n/P <</M#43ID 0 /Lang (pt-BR)>> BDC\n(/MCID 888 BDC EMC \\(literal\\)) Tj\nEMC')));
    const { ws } = await abrir(f), a = validarEstruturaRecorte(ws.pdf);
    ws.aplicarRecorte(await ws.prepararRecorte(1, caixa)); conferirEstruturaRecorte(a, await PDFDocument.load(await ws.obterBytes(), { updateMetadata: false }));
});

test('marcação pode atravessar Contents contíguos sem alterar streams', async () => {
    const f = await criarEstruturaRecorte();
    f.pagina.node.set(N('Contents'), f.c.obj([f.c.register(f.c.flateStream('/P <</MCID 0>> BDC\n')), f.c.register(f.c.flateStream('q Q EMC'))]));
    const { ws } = await abrir(f), a = validarEstruturaRecorte(ws.pdf); ws.aplicarRecorte(await ws.prepararRecorte(1, caixa));
    conferirEstruturaRecorte(a, await PDFDocument.load(await ws.obterBytes(), { updateMetadata: false }));
});

test('exportação de recorte marcado verifica novamente os bytes serializados', async () => {
    const f = await criarEstruturaRecorte(), { ws } = await abrir(f);
    ws.aplicarRecorte(await ws.prepararRecorte(1, caixa)); const estado = ws.capturarEstado(), save = ws.pdf.save.bind(ws.pdf);
    ws.pdf.save = async () => {
        const outro = await PDFDocument.load(await save(), { updateMetadata: false });
        outro.context.lookup(f.elemRef).set(N('ActualText'), PDFString.of('alterado depois de preparar'));
        return outro.save();
    };
    await assert.rejects(ws.obterBytes(), /alterado na serialização/); assert.deepEqual(ws.capturarEstado(), estado);
});

test('cancelamento durante serialização e abertura concorrente preservam o documento atual', async () => {
    for (const cancelar of [true, false]) {
        const f = await criarEstruturaRecorte(), { ws } = await abrir(f), estado = ws.capturarEstado(), abort = new AbortController();
        let liberar, entrou; const gate = new Promise(r => liberar = r), pronto = new Promise(r => entrou = r), save = ws.pdf.save.bind(ws.pdf);
        ws.pdf.save = async () => { entrou(); await gate; return save(); };
        const pendente = ws.prepararRecorte(1, caixa, { signal: abort.signal }); await pronto;
        if (cancelar) abort.abort();
        else { const comum = await PDFDocument.create(); comum.addPage([100, 100]); await ws.abrir(arquivo(await comum.save())); }
        const atual = ws.capturarEstado(); liberar(); await assert.rejects(pendente, /cancelada|mudou/);
        assert.deepEqual(ws.capturarEstado(), cancelar ? estado : atual);
    }
});

// O token vai literalmente no stream comprimido. pdf-lib não recebe esse valor
// como objeto numérico e não pode normalizar um real para inteiro antes do lexer.
async function abrirMcidLiteral(token, esperado = 0) {
    const f = await criarEstruturaRecorte();
    const anteriores = Array.from({ length: esperado }, (_, id) => `/P <</MCID ${id}>> BDC EMC\n`).join('');
    const fonte = anteriores + `/P <</MCID ${token}>> BDC EMC\n`;
    f.elem.set(N('K'), f.c.obj(Array.from({ length: esperado + 1 }, (_, id) => id)));
    for (let id = 1; id <= esperado; id++) f.pais.push(f.elemRef);
    f.pagina.node.set(N('Contents'), f.c.register(f.c.flateStream(fonte)));
    const { ws, bytes } = await abrir(f);
    const s = ws.pdf.context.lookup(ws.pdf.getPage(0).node.get(N('Contents')));
    assert.equal(Buffer.from(decodePDFRawStream(s).decode()).toString('latin1'), fonte, 'token literal deve chegar intacto ao lexer');
    return { ws, bytes, fonte };
}

for (const [token, esperado] of [['0', 0], ['+0', 0], ['-0', 0], ['00', 0], ['+000', 0], ['-000', 0], ['1', 1], ['+1', 1], ['0001', 1], ['+0001', 1], ['7', 7]]) {
    test(`MCID literal inteiro válido ${token}`, async () => {
        const { ws } = await abrirMcidLiteral(token, esperado), antes = validarEstruturaRecorte(ws.pdf);
        assert.equal(antes.mcids, esperado + 1);
        ws.aplicarRecorte(await ws.prepararRecorte(1, caixa));
        conferirEstruturaRecorte(antes, await PDFDocument.load(await ws.obterBytes(), { updateMetadata: false }));
    });
}

const tokensInvalidos = [
    ['0.0'], ['0.'], ['.0'], ['+0.0'], ['-0.0'], ['0.5'],
    ['1.00000000000000001', 1], ['0.99999999999999999', 1],
    ['0.' + '0'.repeat(324) + '1', 0, 'real que sofre underflow para zero'],
    ['9007199254740992'], ['9007199254740993'], ['9007199254740991'], ['10000'],
    ['-1'], ['-0001'], ['(0)'], ['/Zero'], ['null'], ['true'], ['[0]'], ['<< /Value 0 >>'], ['0e0'], [''],
    ['1', 0, 'inteiro sem vínculo em ParentTree']
];
for (const [token, esperado = 0, descricao = token || 'ausente'] of tokensInvalidos) {
    test(`MCID literal recusado: ${descricao}`, async () => {
        const { ws, bytes } = await abrirMcidLiteral(token, esperado), estado = ws.capturarEstado(), original = bytes.slice();
        await assert.rejects(ws.prepararRecorte(1, caixa), /Recorte bloqueado/);
        assert.deepEqual(ws.capturarEstado(), estado); assert.deepEqual(bytes, original);
    });
}
