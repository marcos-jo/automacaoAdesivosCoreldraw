import {PDFDocument, PDFName, PDFArray, PDFNumber, PDFObjectCopier, degrees, rgb} from '/vendor/pdf-lib/pdf-lib.js';
import {planejarMontagem, mmParaPt} from './pdf-imposition-layout.js';
import {diagnosticarMontagem, PDFMontagemCompatibilidadeErro} from './pdf-imposition-compatibility.js';
import {analisarMontagem} from './pdf-imposition-bounds.js';
const N = PDFName.of;
const cancel = signal => { if(signal?.aborted)throw new DOMException('Montagem cancelada','AbortError'); };
function caixa(page,key,fallback){
    const raw=page.node.lookupMaybe(N(key),PDFArray);
    if(!raw)return fallback;
    const a=raw.asArray().map(v=>page.doc.context.lookup(v,PDFNumber).asNumber());
    if(a.length!==4||!a.every(Number.isFinite)||a[2]<=a[0]||a[3]<=a[1])throw Error(`${key} inválida.`);
    return {x:a[0],y:a[1],width:a[2]-a[0],height:a[3]-a[1]};
}
function contem(a,b){return b.x>=a.x-1e-6&&b.y>=a.y-1e-6&&b.x+b.width<=a.x+a.width+1e-6&&b.y+b.height<=a.y+a.height+1e-6;}
export async function gerarMontagem(bytes,config,{signal,previsualizarExcesso=false}={}){
    cancel(signal);
    const source=await PDFDocument.load(new Uint8Array(bytes));
    const compatibility=diagnosticarMontagem(source);
    if(compatibility.necessitaPreparacao)throw new PDFMontagemCompatibilidadeErro(compatibility);
    const selected=config.paginas;
    if(!Array.isArray(selected)||!selected.length||selected.length>10000||selected.some(n=>n!==null&&(!Number.isInteger(n)||n<1||n>source.getPageCount())))throw Error('Seleção de páginas inválida. Use 0 no campo da interface para uma posição vazia.');
    const avisos=new Set(['Montagem não altera o papel nem a escala do driver. Ao imprimir, desative N-up/livreto e confira escala 100%.','Fidelidade de ICC, cores especiais e sobreimpressão ainda requer conferência no RIP.']);
    const infos=new Map();
    for(const number of new Set(selected.filter(Boolean))){
        const page=source.getPage(number-1);
        if(page.node.Annots()?.size())throw Error(`Página ${number}: anotações ou formulários podem ter aparência impressa. Prepare uma cópia antes de montar.`);
        const media=page.getMediaBox(),crop=caixa(page,'CropBox',media),trim=caixa(page,'TrimBox',crop),bleed=caixa(page,'BleedBox',trim);
        if(!contem(media,crop)||!contem(crop,trim)||!contem(crop,bleed)||!contem(bleed,trim))throw Error(`Página ${number}: caixas gráficas incompatíveis.`);
        const unidade=page.node.lookupMaybe(N('UserUnit'),PDFNumber)?.asNumber()??1;
        if(!Number.isFinite(unidade)||unidade<=0||unidade>75000)throw Error('UserUnit inválido.');
        const rotation=((page.getRotation().angle%360)+360)%360;
        if(rotation%90)throw Error('Rotação não ortogonal não suportada.');
        if(!page.node.has(N('TrimBox')))avisos.add('Origem sem TrimBox explícita: o corte usa CropBox. Confira as dimensões finais antes de produzir.');
        if(!page.node.has(N('BleedBox')))avisos.add('Origem sem BleedBox explícita: sangria não disponível; nenhuma borda será inventada.');
        const factor=unidade*25.4/72;
        infos.set(number,{page,trim,bleed,unidade,rotation,id:number,larguraPagina:(rotation%180?media.height:media.width)*factor,alturaPagina:(rotation%180?media.width:media.height)*factor,largura:(rotation%180?trim.height:trim.width)*factor,altura:(rotation%180?trim.width:trim.height)*factor,sangriaDisponivel:Math.max(0,Math.min(trim.x-bleed.x,trim.y-bleed.y,bleed.x+bleed.width-trim.x-trim.width,bleed.y+bleed.height-trim.y-trim.height))*factor});
    }
    const plan=planejarMontagem(config,selected.map(n=>n===null?null:infos.get(n)),{permitirExcesso:previsualizarExcesso});
    const geometria=previsualizarExcesso?analisarMontagem(plan):null;
    const output=await PDFDocument.create(),copier=PDFObjectCopier.for(source.context,output.context),cache=new Map();
    for(let i=0;i<plan.folhas.length;i++){
        cancel(signal);const sheet=output.addPage([mmParaPt(plan.folha.largura),mmParaPt(plan.folha.altura)]);
        for(const placement of plan.folhas[i]){
            const info=infos.get(placement.origem),{page,trim,unidade}=info;
            const scale=placement.escala*unidade;
            const padding=mmParaPt(plan.sangria)/scale;
            const box={left:trim.x-padding,bottom:trim.y-padding,right:trim.x+trim.width+padding,top:trim.y+trim.height+padding};
            const key=`${placement.origem}:${padding}`;
            let embedded=cache.get(key);
            if(!embedded){
                page.pushOperators(); // Normaliza também páginas originalmente vazias na cópia.
                embedded=await output.embedPage(page,box);await embedded.embed();
                const group=page.node.get(N('Group'));
                if(group)output.context.lookup(embedded.ref).dict.set(N('Group'),copier.copy(group));
                cache.set(key,embedded);
            }
            const rotation=(info.rotation+placement.rotacao)%360;
            // O padding é simétrico: não altera o ponto de corte nem a escala.
            const x=mmParaPt(placement.x-plan.sangria),y=mmParaPt(placement.y-plan.sangria);
            const w=(box.right-box.left)*scale,h=(box.top-box.bottom)*scale;
            const offset=rotation===90?[0,w]:rotation===180?[w,h]:rotation===270?[h,0]:[0,0];
            sheet.drawPage(embedded,{x:x+offset[0],y:y+offset[1],width:w,height:h,rotate:degrees(-rotation)});
            for(const line of placement.cortes)sheet.drawLine({start:{x:mmParaPt(line.x1),y:mmParaPt(line.y1)},end:{x:mmParaPt(line.x2),y:mmParaPt(line.y2)},thickness:mmParaPt(plan.marcas.espessura),color:rgb(0,0,0)});
        }
        for(const line of plan.cortesGrupo[i])sheet.drawLine({start:{x:mmParaPt(line.x1),y:mmParaPt(line.y1)},end:{x:mmParaPt(line.x2),y:mmParaPt(line.y2)},thickness:mmParaPt(plan.marcas.espessura),color:rgb(0,0,0)});
        if(i%5===0)await new Promise(resolve=>setTimeout(resolve,0));
    }
    cancel(signal);const result=await output.save();cancel(signal);
    if(result.length>256*1024*1024)throw Error('Montagem acima de 256 MiB. Reduza o lote.');
    const origens=[...infos.values()].map(({id,larguraPagina,alturaPagina,largura,altura})=>({id,larguraPagina,alturaPagina,largura,altura}));
    if(!previsualizarExcesso)return {bytes:result,plan,origens,avisos:[...avisos]};
    cancel(signal);
    const board=await PDFDocument.load(result,{updateMetadata:false}),stage=geometria.palco;
    for(const page of board.getPages()){
        page.setMediaBox(mmParaPt(stage.x),mmParaPt(stage.y),mmParaPt(stage.width),mmParaPt(stage.height));
        page.setCropBox(mmParaPt(stage.x),mmParaPt(stage.y),mmParaPt(stage.width),mmParaPt(stage.height));
    }
    const bytesPrevia=await board.save();cancel(signal);
    return {bytes:result,bytesPrevia,plan,geometria,origens,avisos:[...avisos]};
}
