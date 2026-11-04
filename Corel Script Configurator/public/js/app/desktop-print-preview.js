import {PDFPreview} from '../tools/pdf-preview.js';

// Snapshot visual independente. Nunca fornece pixels ao mecanismo de impressão.
export class PreviaDocumentoImpressao {
    constructor(host){
        this.host=host;this.preview=new PDFPreview(host);this.closed=false;this.page=1;
        this.preview.calcularEscala=base=>Math.min(1.3,700/base.width,900/base.height);
    }
    async open(snapshot){
        const candidate=await this.preview.preparar(snapshot.bytes);
        if(!candidate||this.closed)return false;
        if(!await this.preview.adotar(candidate))return false;
        return this.show(Math.min(Math.max(snapshot.page||1,1),candidate.documento.numPages));
    }
    async show(page){
        const document=this.preview.documento;
        if(this.closed||!document||!Number.isInteger(page)||page<1||page>document.numPages)return false;
        if(!await this.preview.mostrar(page))return false;
        if(this.closed)return false;
        const canvas=this.host.querySelector('canvas'),base=this.preview.exibicao.base;
        this.page=page;
        this.source={width:base.width*25.4/72,height:base.height*25.4/72,page,pages:document.numPages};
        this.image=canvas.toDataURL('image/png');
        return true;
    }
    async close(){this.closed=true;this.source=null;this.image=null;await this.preview.destruir();}
}
