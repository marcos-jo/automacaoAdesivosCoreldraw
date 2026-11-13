import assert from 'node:assert/strict';
import { readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { criarEstruturaRecorte } from './pdf-crop-structure-fixture.mjs';

// Amostras reais são opcionais e permanecem fora do Git. Nenhum upload.
export async function testarRecorteEstruturadoBrowser(avaliar, evidencias) {
    const fixture = await criarEstruturaRecorte();
    const casos = [{ id: 'sintetico-marcado', bytes: await fixture.pdf.save() }];
    if (process.env.PDF_CROP_REAL_LIBRARY) for (const id of ['CASO-008', 'CASO-009']) {
        casos.push({ id, bytes: await readFile(path.join(process.env.PDF_CROP_REAL_LIBRARY, 'originais-copiados', id + '.pdf')) });
    }
    const resultados = [];
    for (const caso of casos) {
        const r = await avaliar(`(async()=>{
            const bytes=new Uint8Array(${JSON.stringify([...caso.bytes])});
            const {PDFWorkspace}=await import('/js/tools/pdf-workspace.js');
            const {abrirRecorte}=await import('/js/tools/pdf-crop.js');
            const {validarEstruturaRecorte,conferirEstruturaRecorte}=await import('/js/tools/pdf-crop-structure.js');
            const {PDFDocument}=await import('/vendor/pdf-lib/pdf-lib.js');
            const pdfjs=await import('/vendor/pdfjs/pdf.js');pdfjs.GlobalWorkerOptions.workerSrc='/vendor/pdfjs/pdf.worker.js';
            const ws=new PDFWorkspace();await ws.abrir(new File([bytes],'marcado.pdf'));
            const antes=validarEstruturaRecorte(ws.pdf), originalTask=pdfjs.getDocument({data:bytes.slice()}),original=await originalTask.promise;
            const root=document.createElement('div');document.body.append(root);let saidaTask;
            try {
                const p=await original.getPage(1),viewport=p.getViewport({scale:1});
                const render=async(page,vp)=>{const canvas=document.createElement('canvas');canvas.width=Math.ceil(vp.width);canvas.height=Math.ceil(vp.height);await page.render({canvasContext:canvas.getContext('2d'),viewport:vp}).promise;return canvas;};
                const fonte=await render(p,viewport);root.append(fonte);
                const previa={container:root,exibicao:{viewport}};
                const ready=async()=>{for(let n=0;n<600;n++){const b=root.querySelector('[data-aplicar]');if(b&&!b.disabled)return;await new Promise(r=>setTimeout(r,10));}throw Error('Recorte marcado não ficou pronto');};
                const abrir=()=>abrirRecorte({workspace:ws,previa,numero:1,raiz:root,aplicar:async(prep,op)=>ws.aplicarRecorte(prep,op)});
                const rev=ws.revisao;let sessao=abrir();await ready();
                const aviso=root.querySelector('dialog').textContent;
                if(!aviso.includes('tecnologias assistivas')||!aviso.includes('não certifica PDF/UA'))throw Error('Aviso incompleto');
                root.querySelector('[data-cancelar]').click();await sessao;if(ws.revisao!==rev)throw Error('Cancelar modificou o PDF');
                sessao=abrir();await ready();
                for(const i of root.querySelectorAll('.pdf-recorte-campos input'))if(i.name!=='zoom'){i.value=10;i.dispatchEvent(new Event('input'));}
                await ready();root.querySelector('[data-aplicar]').click();await sessao;
                const out=await ws.obterBytes(), re=await PDFDocument.load(out,{updateMetadata:false}),depois=conferirEstruturaRecorte(antes,re),box=re.getPage(0).getCropBox();
                saidaTask=pdfjs.getDocument({data:out.slice()});const saidaDoc=await saidaTask.promise,rp=await saidaDoc.getPage(1),vp=rp.getViewport({scale:1});
                const esperadoViewport=new viewport.constructor({viewBox:[box.x,box.y,box.x+box.width,box.y+box.height],userUnit:1,scale:1,rotation:0,offsetX:0,offsetY:0,dontFlip:false});
                const esperado=await render(p,esperadoViewport),resultado=await render(rp,vp);
                if(esperado.width!==resultado.width||esperado.height!==resultado.height)throw Error('Dimensões de render divergentes');
                const a=esperado.getContext('2d').getImageData(0,0,esperado.width,esperado.height).data,b=resultado.getContext('2d').getImageData(0,0,resultado.width,resultado.height).data;
                let diferentes=0,maximo=0;for(let i=0;i<a.length;i++){const dif=Math.abs(a[i]-b[i]);if(dif)diferentes++;maximo=Math.max(maximo,dif);}
                const treeAntes=await p.getStructTree(),treeDepois=await rp.getStructTree();
                const base64=b=>{let s='';for(let i=0;i<b.length;i+=8192)s+=String.fromCharCode(...b.subarray(i,i+8192));return btoa(s);};
                return {elementosAntes:antes.elementos,elementosDepois:depois.elementos,mcidsAntes:antes.mcids,mcidsDepois:depois.mcids,box,componentes:a.length,diferentes,maximo,arvoreVisualIgual:JSON.stringify(treeAntes)===JSON.stringify(treeDepois),cancelamento:true,aviso:true,bytes:base64(out),pngAntes:fonte.toDataURL(),pngEsperado:esperado.toDataURL(),pngDepois:resultado.toDataURL()};
            }finally{root.remove();if(saidaTask)await saidaTask.destroy();await originalTask.destroy();}
        })()`);
        assert.equal(r.diferentes, 0, caso.id + ': aparência diferente da região original');
        assert.equal(r.arvoreVisualIgual, true, caso.id + ': árvore PDF.js diferente');
        for (const [nome, campo] of [['antes', 'pngAntes'], ['regiao-original', 'pngEsperado'], ['depois', 'pngDepois']]) {
            await writeFile(path.join(evidencias, caso.id + '-' + nome + '.png'), Buffer.from(r[campo].split(',')[1], 'base64'), { flag: 'wx' }); delete r[campo];
        }
        await writeFile(path.join(evidencias, caso.id + '-chrome-recorte.pdf'), Buffer.from(r.bytes, 'base64'), { flag: 'wx' }); delete r.bytes;
        resultados.push({ id: caso.id, ...r });
    }
    await writeFile(path.join(evidencias, 'recorte-estruturado.json'), JSON.stringify(resultados, null, 2), { flag: 'wx' });
    console.log('PASSOU recorte estruturado Chrome: ' + resultados.map(r => r.id + ' (' + r.elementosDepois + ' elementos/' + r.mcidsDepois + ' MCIDs, ' + r.componentes + ' componentes RGBA iguais)').join('; ') + '. Evidências: ' + evidencias);
}
