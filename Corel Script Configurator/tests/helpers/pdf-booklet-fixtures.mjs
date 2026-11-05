import {PDFDocument,PDFName,rgb,degrees} from '../../public/vendor/pdf-lib/pdf-lib.js';
export async function criarOrigemLivreto(count,{rotation=0,largura=148,altura=210}={}){
    const pdf=await PDFDocument.create();
    for(let i=1;i<=count;i++){
        const p=pdf.addPage([largura*72/25.4,altura*72/25.4]);
        p.node.set(PDFName.of('TrimBox'),pdf.context.obj([0,0,p.getWidth(),p.getHeight()]));
        p.drawRectangle({x:8,y:8,width:p.getWidth()-16,height:p.getHeight()-16,borderWidth:1,borderColor:rgb(.1,.25,.65)});
        p.drawRectangle({x:20,y:25,width:80,height:70,color:rgb(.9,.1,.15),opacity:.5});
        p.drawRectangle({x:50,y:50,width:80,height:70,color:rgb(.1,.2,.9),opacity:.5});
        p.drawText(`PAGINA ${i}`,{x:32,y:p.getHeight()/2,size:35});
        p.drawText(`TOPO ^ | Livreto sintetico ${count} paginas`,{x:20,y:p.getHeight()-30,size:12});
        p.drawText(`BASE | ${i}`,{x:20,y:18,size:12});
        if(rotation)p.setRotation(degrees(rotation));
    }
    return pdf.save();
}
