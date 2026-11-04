import {PDFDocument,StandardFonts,degrees,rgb} from '../../public/vendor/pdf-lib/pdf-lib.js';

const points=mm=>mm*72/25.4;
const fixedDate=new Date('2026-11-04T00:00:00Z');
export const printScales=Object.freeze([100,120,87.5]);
function metadata(pdf){
    pdf.setTitle('4PrintShops synthetic print fixture');pdf.setAuthor('4PrintShops tests');
    pdf.setCreator('4PrintShops fixture generator');pdf.setProducer('4PrintShops / pdf-lib');
    pdf.setCreationDate(fixedDate);pdf.setModificationDate(fixedDate);
}
export async function createPaperFixture({width=210,height=297,rotation=0}={}){
    const pdf=await PDFDocument.create();metadata(pdf);
    const page=pdf.addPage([points(width),points(height)]);
    page.setRotation(degrees(rotation));
    page.drawText('4PrintShops SYNTHETIC PRINT TEST',{x:20,y:points(height)-35,size:14});
    page.drawText(`${width} x ${height} mm / rotation ${rotation}`,{x:20,y:points(height)-58,size:11});
    page.drawRectangle({x:30,y:50,width:points(50),height:points(20),color:rgb(0.2,0.4,0.8)});
    return pdf.save();
}
export async function createMixedPrintFixture(baseBytes){
    const pdf=await PDFDocument.load(baseBytes||await createPaperFixture());metadata(pdf);
    pdf.getPage(0).setRotation(degrees(180));
    pdf.addPage([points(300),points(120)]).drawText('LANDSCAPE / 300 x 120 mm',{x:30,y:80,size:15,color:rgb(1,0,0)});
    pdf.addPage([points(100),points(150)]).drawText('CUSTOM / 100 x 150 mm',{x:20,y:120,size:13,color:rgb(0,0.5,0)});
    return pdf.save();
}
export async function createPrintFormFixture(){
    const pdf=await PDFDocument.create();metadata(pdf);
    const page=pdf.addPage([points(210),points(297)]),font=await pdf.embedFont(StandardFonts.Helvetica);
    const field=pdf.getForm().createTextField('test');
    field.setText('PRINTABLE FORM');field.addToPage(page,{x:40,y:500,width:200,height:30});
    pdf.getForm().updateFieldAppearances(font);
    return pdf.save();
}
