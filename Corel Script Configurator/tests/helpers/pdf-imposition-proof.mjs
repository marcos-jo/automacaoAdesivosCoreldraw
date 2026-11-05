import assert from 'node:assert/strict';
import {PDFDocument,PDFName,decodePDFRawStream} from '../../public/vendor/pdf-lib/pdf-lib.js';
const mult=(a,b)=>[a[0]*b[0]+a[2]*b[1],a[1]*b[0]+a[3]*b[1],a[0]*b[2]+a[2]*b[3],a[1]*b[2]+a[3]*b[3],a[0]*b[4]+a[2]*b[5]+a[4],a[1]*b[4]+a[3]*b[5]+a[5]];
const identity=()=>[1,0,0,1,0,0],point=(m,x,y)=>[m[0]*x+m[2]*y+m[4],m[1]*x+m[3]*y+m[5]];
const near=(a,b)=>assert.ok(Math.abs(a-b)<.00001,`${a} != ${b}`);
// Confere a transformação realmente serializada, incluindo Matrix/BBox do Form.
export async function conferirTransformacoes(bytes,plan){
    const pdf=await PDFDocument.load(bytes);assert.equal(pdf.getPageCount(),plan.totalFolhas);let checked=0;
    for(const [index,page]of pdf.getPages().entries()){
        near(page.getWidth(),plan.folha.largura*72/25.4);near(page.getHeight(),plan.folha.altura*72/25.4);
        const stream=page.node.Contents().asArray().map(ref=>Buffer.from(decodePDFRawStream(pdf.context.lookup(ref)).decode()).toString()).join('\n');
        const stack=[],boxes=[];let matrix=identity();
        const expression=/(-?[\d.]+)\s+(-?[\d.]+)\s+(-?[\d.]+)\s+(-?[\d.]+)\s+(-?[\d.]+)\s+(-?[\d.]+)\s+cm\b|\bq\b|\bQ\b|\/([^\s]+)\s+Do\b/g;
        for(const match of stream.matchAll(expression)){
            if(match[1]!==undefined){matrix=mult(matrix,match.slice(1,7).map(Number));continue;}if(match[0]==='q'){stack.push(matrix.slice());continue;}if(match[0]==='Q'){matrix=stack.pop();continue;}
            const form=page.node.Resources().lookup(PDFName.of('XObject')).lookup(PDFName.of(match[7]));assert.equal(form.dict.get(PDFName.of('Subtype')).toString(),'/Form');
            const bbox=form.dict.lookup(PDFName.of('BBox')).asArray().map(v=>v.asNumber()),fm=form.dict.lookup(PDFName.of('Matrix'))?.asArray().map(v=>v.asNumber())||identity(),m=mult(matrix,fm);
            const points=[[bbox[0],bbox[1]],[bbox[2],bbox[1]],[bbox[0],bbox[3]],[bbox[2],bbox[3]]].map(([x,y])=>point(m,x,y));
            boxes.push([Math.min(...points.map(p=>p[0])),Math.min(...points.map(p=>p[1])),Math.max(...points.map(p=>p[0])),Math.max(...points.map(p=>p[1]))]);
        }
        assert.equal(boxes.length,plan.folhas[index].length);for(const [i,piece]of plan.folhas[index].entries()){const bleed=plan.sangria||0,expected=[piece.x-bleed,piece.y-bleed,piece.x+piece.largura+bleed,piece.y+piece.altura+bleed].map(n=>n*72/25.4);expected.forEach((value,k)=>near(boxes[i][k],value));checked++;}
    }
    return checked;
}
