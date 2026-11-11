import assert from 'node:assert/strict';import {writeFile} from 'node:fs/promises';import path from 'node:path';
import {conferirTransformacoes} from './pdf-imposition-proof.mjs';
export async function testarEscalasMontagem({ev,wait,socket,cdp,folder,screenshot}){
    const snapshots=[],dialogs=[];let accept=true;
    const handler=e=>{const m=JSON.parse(e.data);if(m.method==='Page.javascriptDialogOpening'){dialogs.push(m.params.message);cdp('Page.handleJavaScriptDialog',{accept}).catch(()=>{});}};socket.addEventListener('message',handler);
    const ready=()=>wait(()=>ev('(()=>{const u=window.__scaleUi;if(!u.get("tentar").hidden)throw Error(u.get("status").textContent);return !!u.result&&!u.busy;})()'),'escala atualizada');
    const edit=async(id,value)=>{await ev(`(()=>{const f=window.__scaleUi.get(${JSON.stringify(id)});f.value=${JSON.stringify(String(value))};f.dispatchEvent(new Event('input',{bubbles:true}));})()`);await ready();};
    const capture=async name=>{
        assert.equal(await ev("(async()=>{const {resumirMontagem}=await import('/js/tools/pdf-imposition-summary.js');return window.__scaleUi.get('encaixe').textContent.includes('Escala final das artes: '+resumirMontagem(window.__scaleUi.result).efetiva+'.');})()"),true);
        const s=await ev(`(()=>{const u=window.__scaleUi,r=u.result;let binary='';for(let i=0;i<r.bytes.length;i+=32768)binary+=String.fromCharCode(...r.bytes.subarray(i,i+32768));return {name:${JSON.stringify(name)},modo:u.get('modo').value,grade:u.get('grade').value,fator:u.fatorMontagem,config:r.config,plan:r.plan,resumo:u.get('medidas-resumo').textContent,manualVisivel:u.get('escala').checkVisibility(),advanced:u.get('avancadas').open,pdf:btoa(binary),origens:r.origens};})()`);
        const bytes=Buffer.from(s.pdf,'base64');s.transformacoes=await conferirTransformacoes(bytes,s.plan);await writeFile(path.join(folder,name+'.pdf'),bytes);delete s.pdf;snapshots.push(s);return s;
    };
    const near=(a,b)=>assert.ok(Math.abs(a-b)<1e-7,`${a} != ${b}`);
    try{
        for(const pages of [1,4]){
            await ev(`(async()=>{const {PDFDocument,rgb,degrees}=await import('/vendor/pdf-lib/pdf-lib.js'),pdf=await PDFDocument.create();const sizes=[[395.82,494.59],[297,420],[100,150],[200,100]].slice(0,${pages});for(const [i,[w,h]]of sizes.entries()){const p=pdf.addPage([w*72/25.4,h*72/25.4]);p.drawRectangle({x:0,y:0,width:p.getWidth(),height:p.getHeight(),color:rgb(.3+i*.1,.5,.8)});p.drawText('Prova sintetica '+(i+1),{x:30,y:60,size:16});if(i===2)p.setRotation(degrees(90));}const source=await pdf.save();window.__scaleOriginal=source.slice();const {JanelaMontagem}=await import('/js/tools/pdf-imposition-ui.js');window.__scaleRoot=document.createElement('div');window.__scaleRoot.className='pdf-editor';document.body.append(window.__scaleRoot);window.__scaleUi=new JanelaMontagem(window.__scaleRoot,{snapshot:async()=>source.slice(),imprimir:async r=>{window.__scalePrinted=r.bytes===window.__scaleToPrint;}});await window.__scaleUi.open();})()`);await ready();
            await edit('modoMarcas','sem');
            const real=await capture(`escala-${pages}-repeat-100`);near(real.plan.folhas[0][0].escala,1);assert.equal(real.advanced,false);assert.match(real.resumo,pages===1?/395,82 × 494,59/:/4 tamanhos diferentes/);assert.equal(real.plan.totalFolhas,pages);near(real.plan.folhas[0][0].largura,395.82);near(real.plan.folhas[0][0].altura,494.59);assert.equal(real.manualVisivel,false);assert.equal(real.config.margens.esquerda,0);
            await edit('escalaModo','personalizada');await edit('escala','120');const repeat=await capture(`escala-${pages}-repeat-120`);near(repeat.plan.folhas[0][0].escala,1.2);assert.equal(repeat.manualVisivel,false); // avançadas recolhidas
            await edit('modo','nup');assert.equal(await ev('window.__scaleUi.get("escalaModo").value'),'real');await edit('escalaModo','preencher');if(pages===4)await edit('paginas','1-4');for(const id of ['esquerda','direita','superior','inferior'])await edit(id,'10');
            const nup=await capture(`escala-${pages}-nup-auto`);assert.equal(nup.config.escalaModo,'preencher');assert.equal(nup.grade,'4');assert.equal(nup.manualVisivel,false);assert.match(nup.resumo,/Manual: não aplicada/);near(nup.plan.folhas[0][0].largura,126.4);assert.equal(nup.plan.folhas[0].length,pages);
            await edit('modo','repeat');const returned=await capture(`escala-${pages}-repeat-retorno`);assert.equal(returned.grade,real.grade);assert.equal(returned.config.paginas.length,pages);assert.equal(returned.fator,1);assert.deepEqual(returned.plan,repeat.plan);
            await edit('modo','nup');await edit('escalaModo','personalizada');await edit('escala','120');await capture(`escala-${pages}-nup-120`);
            await ev('window.__scaleUi.fit()');await ready();const fit=await capture(`escala-${pages}-nup-fit`);assert.ok(fit.fator<1);assert.match(fit.resumo,/Ajuste para caber: /);assert.equal(fit.grade,'custom');await screenshot(`escala-${pages}-encaixe-explicito`);
            await edit('modo','repeat');near((await capture(`escala-${pages}-repeat-sem-herdar-fit`)).plan.folhas[0][0].escala,1.2);
            await edit('modo','nup');const restored=await capture(`escala-${pages}-nup-fit-restaurado`);assert.deepEqual(restored.plan,fit.plan);near(restored.fator,fit.fator);
            for(let i=0;i<3;i++){await ev('window.__scaleUi.fit()');await ready();near(await ev('window.__scaleUi.result.plan.folhas[0][0].escala'),fit.plan.folhas[0][0].escala);}
            await ev('window.__scaleUi.get("desfazer-ajuste").click()');await ready();const removed=await capture(`escala-${pages}-nup-remover-fit`);assert.equal(removed.fator,1);assert.equal(removed.grade,'4');near(removed.plan.folhas[0][0].escala,1.2);
            await ev('window.__scaleUi.get("tamanho-real").click()');await ready();near((await capture(`escala-${pages}-nup-voltar-100`)).plan.folhas[0][0].escala,1);
            await edit('modo','repeat');await edit('escalaModo','medidas');await edit('pecaLargura','90');near((await capture(`escala-${pages}-repeat-medidas`)).plan.folhas[0][0].largura,90);
            await ev('window.__scaleUi.get("tamanho-real").click()');await ready();const reset=await capture(`escala-${pages}-repeat-tamanho-real`);near(reset.plan.folhas[0][0].escala,1);assert.equal(reset.config.pecaLargura,undefined);assert.equal(reset.fator,1);
            await ev('window.__scaleUi.fit()');await ready();const fittedRepeat=await capture(`escala-${pages}-repeat-fit`);await ev('window.__scaleUi.get("desfazer-ajuste").click()');await ready();assert.equal(await ev('window.__scaleUi.get("grade").value'),real.grade);near(await ev('window.__scaleUi.result.plan.folhas[0][0].escala'),1);
            await ev('window.__scaleUi.fit()');await ready();near(await ev('window.__scaleUi.result.plan.folhas[0][0].escala'),fittedRepeat.plan.folhas[0][0].escala);
            accept=false;await ev('window.__scaleUi.save()');assert.equal(await ev('window.__scaleUi.busy'),false);assert.equal(await ev('window.__scaleUi.canOutput()'),true);accept=true;
            assert.equal(await ev('window.__scaleUi.originalBytes.every((v,i)=>v===window.__scaleOriginal[i])'),true);assert.equal(await ev('document.querySelector("#pdf-alteracoes-pendentes").hidden'),false);
            await ev('window.__scaleToPrint=window.__scaleUi.result.bytes;window.__scaleUi.get("imprimir").click()');await wait(()=>ev('window.__scalePrinted===true'),'bytes físicos encaminhados ao callback de teste');await ev('window.__scaleUi.destroy();window.__scaleRoot.remove();window.__scalePrinted=false');
        }
        assert.ok(dialogs.some(t=>t.startsWith('Conferir montagem antes de gerar o PDF?')&&t.includes('Peça final:')&&t.includes('Grade:')));
        assert.ok(dialogs.some(t=>t.startsWith('Imprimir a montagem já imposta?')&&t.includes('escala efetiva:')));
        await writeFile(path.join(folder,'escalas-e-matrizes.json'),JSON.stringify({snapshots,confirmacoes:dialogs.length,sourceUnchanged:true,physicalPrint:false},null,2));
    }finally{socket.removeEventListener('message',handler);await ev('window.__scaleUi?.destroy();window.__scaleRoot?.remove()').catch(()=>{});}
}
