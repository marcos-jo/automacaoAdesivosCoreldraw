const positive=value=>typeof value==='number'&&Number.isFinite(value)&&value>0&&value<100000;
export function calcularPreviaImpressao(paper,scale=100,printable=null){
    if(!paper||!positive(paper.width)||!positive(paper.height)||![100,120].includes(scale))throw new Error('Dimensões ou escala inválidas.');
    const width=210*scale/100,height=297*scale/100,left=(paper.width-width)/2,top=(paper.height-height)/2;
    const validPrintable=printable&&[printable.width,printable.height,printable.printableWidth,printable.printableHeight].every(positive)
        &&[printable.left,printable.top].every(n=>typeof n==='number'&&Number.isFinite(n)&&n>=0)
        &&printable.left+printable.printableWidth<=printable.width+0.5&&printable.top+printable.printableHeight<=printable.height+0.5;
    const compatible=!!validPrintable&&Math.abs(printable.width-paper.width)<1&&Math.abs(printable.height-paper.height)<1;
    const crosses=(x,y,w,h)=>left<x-0.001||top<y-0.001||left+width>x+w+0.001||top+height>y+h+0.001;
    return {paper:{width:paper.width,height:paper.height},content:{left,top,width,height},scale,
        margins:{left:Math.max(0,left),right:Math.max(0,left),top:Math.max(0,top),bottom:Math.max(0,top)},
        outsidePaper:crosses(0,0,paper.width,paper.height),outsidePrintable:compatible?crosses(printable.left,printable.top,printable.printableWidth,printable.printableHeight):null,
        printable:compatible?{left:printable.left,top:printable.top,width:printable.printableWidth,height:printable.printableHeight}:null,
        geometryMismatch:!!validPrintable&&!compatible};
}

export class PrintConfigPresentation {
    constructor(){this.generation=0;this.state=null;}
    begin(){return ++this.generation;}
    accept(ticket,state){if(ticket!==this.generation)return false;if(this.state?.session===state.session&&state.revision<this.state.revision)return false;this.state=state;return true;}
    clear(){++this.generation;this.state=null;}
}
