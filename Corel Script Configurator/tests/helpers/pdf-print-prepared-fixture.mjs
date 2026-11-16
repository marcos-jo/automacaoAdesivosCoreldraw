import {PDFDocument, pushGraphicsState, popGraphicsState, concatTransformationMatrix, drawObject} from '../../public/vendor/pdf-lib/pdf-lib.js';

// PDF sintético já preparado: sem campos/assinaturas, página 1 vetorial e
// página 2 com imagem RGB. Não remove assinatura nem simula histórico do app.
// Os 2×2 pixels servem ao contrato estrutural do smoke, não à homologação de DPI.
export async function criarPDFPreparadoSintetico() {
    const pdf=await PDFDocument.create(),vetorial=pdf.addPage([300,400]);
    vetorial.drawText('Documento sintetico pagina 1',{x:10,y:200,size:12});
    const raster=pdf.addPage([300,400]);
    const imagem=pdf.context.register(pdf.context.flateStream(new Uint8Array([
        255,255,255, 0,128,0, 0,0,255, 255,255,255
    ]),{Type:'XObject',Subtype:'Image',Width:2,Height:2,ColorSpace:'DeviceRGB',BitsPerComponent:8}));
    const chave=raster.node.newXObject('FixtureRaster',imagem);
    raster.pushOperators(pushGraphicsState(),concatTransformationMatrix(300,0,0,400,0,0),drawObject(chave),popGraphicsState());
    return pdf.save({updateFieldAppearances:false});
}
