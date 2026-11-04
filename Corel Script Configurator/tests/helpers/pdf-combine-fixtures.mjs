import {PDFDocument, rgb} from '../../public/vendor/pdf-lib/pdf-lib.js';
import {deflateSync, crc32} from 'node:zlib';
// JPEG RGB 120×180 sintético, retângulos vermelho/azul, gerado com System.Drawing.
const jpeg = Buffer.from('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', 'base64');
function chunk(type, data) {
    const name = Buffer.from(type), size = Buffer.alloc(4), crc = Buffer.alloc(4);
    size.writeUInt32BE(data.length); crc.writeUInt32BE(crc32(Buffer.concat([name, data])));
    return Buffer.concat([size, name, data, crc]);
}
export function png({profile = false,width=120,height=80} = {}) {
    const header = Buffer.alloc(13); header.writeUInt32BE(width); header.writeUInt32BE(height,4); header[8]=8; header[9]=6;
    const raw=Buffer.alloc(height*(1+width*4));
    for(let y=0;y<height;y++)for(let x=0;x<width;x++){const i=y*(1+width*4)+1+x*4;raw[i]=20;raw[i+1]=180;raw[i+2]=50;raw[i+3]=x<width/2?128:255;}
    return Buffer.concat([Buffer.from([137,80,78,71,13,10,26,10]),chunk('IHDR',header),...(profile?[chunk('iCCP',Buffer.from([0]))]:[]),chunk('IDAT',deflateSync(raw)),chunk('IEND',Buffer.alloc(0))]);
}
export function jpg(orientation=1,{icc=false,dpi=0}={}) {
    const exif=Buffer.from([69,120,105,102,0,0,73,73,42,0,8,0,0,0,1,0,18,1,3,0,1,0,0,0,orientation,0,0,0,0,0,0,0]);
    const segment=(marker,data)=>{const length=Buffer.alloc(2);length.writeUInt16BE(data.length+2);return Buffer.concat([Buffer.from([255,marker]),length,data]);};
    const copy=Buffer.from(jpeg);
    const jfif=copy.indexOf(Buffer.from('JFIF\0'));if(jfif>=0){copy[jfif+7]=1;copy.writeUInt16BE(dpi,jfif+8);copy.writeUInt16BE(dpi,jfif+10);}
    return Buffer.concat([copy.subarray(0,2),segment(225,exif),...(icc?[segment(226,Buffer.from('ICC_PROFILE\0test'))]:[]),copy.subarray(2)]);
}
export async function fixtures(){
    const result=[];
    for(const [name,widths]of [['Primeiro A4 A3.pdf',[595.2756,841.8898]],['Segundo acentuação.pdf',[420,460]]]){
        const pdf=await PDFDocument.create();for(let i=0;i<widths.length;i++){const p=pdf.addPage([widths[i],i?1190.5512:841.8898]);p.drawText(name.replace(/[^a-zA-Z0-9 .]/g,'')+' / '+(i+1),{x:30,y:60,size:16});p.drawRectangle({x:40,y:100,width:80,height:50,color:rgb(0.1,0.3,0.8)});}result.push(new File([await pdf.save()],name,{type:'application/pdf'}));
    }
    result.push(new File([jpg(1)],'Retrato sem DPI.jpg',{type:'image/jpeg'}),new File([jpg(6,{dpi:1})],'Paisagem EXIF inválido DPI.jpeg',{type:'image/jpeg'}),new File([png()],'Transparência.png',{type:'image/png'}));
    return result;
}
