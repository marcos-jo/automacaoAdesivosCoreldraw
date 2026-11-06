import {PDFDocument,PDFName,PDFNumber} from '/vendor/pdf-lib/pdf-lib.js';
import {planejarCopiaImpressao,reconstruirCopiaImpressao} from './pdf-print-copy.js';
import {estimarRaster,rasterizarPaginaImpressao,conferirCopiaImpressao} from './pdf-print-raster.js';
import {diagnosticarMontagem} from './pdf-imposition-compatibility.js';
import {aguardarMontagem} from './pdf-imposition-preview.js';

const cancelar=signal=>{if(signal?.aborted)throw new DOMException('Preparação de imagem cancelada','AbortError');};
const LIMITE=256*1024*1024;
export const avisoImagemMontagem='Todas as páginas serão convertidas em pixels RGB opacos sobre branco. Texto deixará de ser pesquisável e vetores perderão sua natureza vetorial. Camadas, links, campos, acessibilidade e assinaturas ativas não serão transportados; a cópia não preserva a validade de assinatura. Perfis ICC e intenções de saída não são preservados como no original; cores e transparências podem diferir no RIP. Mais DPI não recupera detalhes ausentes. O preparo é limitado pela memória e não altera o workspace.';

// Preflight sem canvas/worker: mantém todos os bloqueios do preparador existente.
export function planejarImagemMontagem(pdf,dpi){
    if(!pdf.getPageCount())throw Error('O documento não contém páginas para preparar.');
    const plano=planejarCopiaImpressao(pdf,{modo:'inteiro',dpi,paginas:[]});
    let total=0;
    const paginas=pdf.getPages().map((p,i)=>{
        const m=p.getMediaBox(),c=p.getCropBox(),unit=p.node.lookupMaybe(PDFName.of('UserUnit'),PDFNumber)?.asNumber()??1;
        const rotation=((p.getRotation().angle%360)+360)%360;
        if(!Number.isFinite(unit)||unit<=0||unit>75000||![0,90,180,270].includes(rotation))throw Error(`Página ${i+1}: rotação/UserUnit não suportado para imagem.`);
        const w=(Math.min(m.x+m.width,c.x+c.width)-Math.max(m.x,c.x))*unit,h=(Math.min(m.y+m.height,c.y+c.height)-Math.max(m.y,c.y))*unit;
        const size=estimarRaster(rotation%180?h:w,rotation%180?w:h,dpi);
        total+=size.pixels*3;
        if(total>LIMITE)throw Error('A estimativa RGB das páginas excede 256 MiB. Escolha 300 DPI ou divida o documento; não haverá redução automática.');
        return {pagina:i+1,...size};
    });
    return {plano,paginas,totalRGB:total};
}
export async function avaliarImagemMontagem(bytes){
    if(typeof CompressionStream==='undefined'||typeof DecompressionStream==='undefined')throw Error('Compressão local indisponível neste navegador.');
    const pdf=await PDFDocument.load(new Uint8Array(bytes),{updateMetadata:false});
    return [300,600].map(dpi=>{try{return {dpi,disponivel:true,...planejarImagemMontagem(pdf,dpi)};}catch(e){return {dpi,disponivel:false,motivo:e.message};}});
}

export async function prepararImagemMontagem(bytes,{dpi,confirmado=false,signal,progresso=()=>{},load=()=>import('/vendor/pdfjs/pdf.js')}={}){
    cancelar(signal);
    if(!confirmado)throw Error('Confirme explicitamente a preparação da cópia como imagem.');
    const original=new Uint8Array(bytes),copia=await PDFDocument.load(original.slice(),{updateMetadata:false});
    const estimativa=planejarImagemMontagem(copia,dpi);
    const pdfjs=await load();cancelar(signal);pdfjs.GlobalWorkerOptions.workerSrc='/vendor/pdfjs/pdf.worker.js';
    const candidatos=[];
    const abrir=async data=>{
        cancelar(signal);const worker=new pdfjs.PDFWorker();let task;
        try{task=pdfjs.getDocument({data:data.slice(),worker});}catch(e){worker.destroy();throw e;}
        candidatos.push({task,worker});
        return aguardarMontagem(task.promise,{signal,timeout:30000});
    };
    let preparada;
    try{
        const fonte=await abrir(original);
        preparada=await reconstruirCopiaImpressao(copia,{modo:'inteiro',dpi,paginas:[]},{signal,progresso,rasterizar:(n,plano)=>rasterizarPaginaImpressao(fonte,n,plano,{signal,progresso})});
        cancelar(signal);const resultado=await preparada.pdf.save({updateFieldAppearances:false});
        if(resultado.length>LIMITE)throw Error('Cópia como imagem excede 256 MiB; divida o documento.');
        const diag=diagnosticarMontagem(await PDFDocument.load(resultado,{updateMetadata:false}));
        if(diag.necessitaPreparacao)throw Error('A cópia como imagem ainda contém recursos incompatíveis. Nenhuma cópia foi aplicada.');
        const candidato=await abrir(resultado);
        const validacao=await conferirCopiaImpressao(fonte,candidato,preparada,{signal,progresso});
        cancelar(signal);
        if(!validacao.compativel)throw Error('A cópia como imagem não passou na comparação de aparência/dimensões.');
        return {bytes:resultado,validacao,imagem:{dpi,paginas:estimativa.paginas.length,area:'Área visível (CropBox), mantendo as caixas e sua orientação visual',aviso:avisoImagemMontagem}};
    }finally{
        preparada?.imagens.clear();
        // Dono único de cada tarefa/worker; destroy nunca é chamado no documento.
        for(const {task,worker}of candidatos){let timer;try{await Promise.race([Promise.resolve().then(()=>task.destroy()).catch(()=>{}),new Promise(r=>{timer=setTimeout(r,2000);})]);}finally{clearTimeout(timer);worker.destroy();}}
    }
}
