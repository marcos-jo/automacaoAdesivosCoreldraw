import { PDFDocument, PDFName, PDFString } from '../../public/vendor/pdf-lib/pdf-lib.js';
const N = PDFName.of;
export async function criarEstruturaRecorte() {
    const pdf = await PDFDocument.create(), c = pdf.context, pagina = pdf.addPage([720, 540]);
    pagina.drawText('Conteudo vetorial preservado', { x: 50, y: 100 });
    const recursos = pagina.node.Resources();
    // Envolve o conteúdo que pdf-lib já gerou, conservando fontes e operadores.
    const original = pagina.node.Contents().asArray().map(r => Buffer.from(c.lookup(r).getUnencodedContents()).toString('latin1')).join('\n');
    for (const ref of pagina.node.Contents().asArray()) c.delete(ref);
    const stream = c.flateStream('/P <</MCID 0>> BDC\n' + original + '\nEMC\n');
    pagina.node.set(N('Contents'), c.register(stream));
    pagina.node.set(N('StructParents'), c.obj(0)); pagina.node.set(N('Tabs'), N('S'));
    const root = c.obj({ Type: 'StructTreeRoot' }), rootRef = c.register(root);
    const elem = c.obj({ Type: 'StructElem', S: 'Textbox', P: rootRef, Pg: pagina.ref, K: 0, ActualText: PDFString.of('Texto logico'), Alt: PDFString.of('Descricao') }), elemRef = c.register(elem);
    const pais = c.obj([elemRef]), pt = c.obj({ Nums: [0, c.register(pais)] });
    root.set(N('K'), c.obj([elemRef])); root.set(N('ParentTree'), c.register(pt)); root.set(N('ParentTreeNextKey'), c.obj(1)); root.set(N('RoleMap'), c.obj({ Textbox: 'Sect' }));
    pdf.catalog.set(N('StructTreeRoot'), rootRef); pdf.catalog.set(N('MarkInfo'), c.obj({ Marked: true })); pdf.catalog.set(N('Lang'), PDFString.of('pt-BR'));
    await pdf.flush();
    return { pdf, c, pagina, root, rootRef, elem, elemRef, pais, pt, stream, recursos };
}
