import assert from "node:assert/strict";
import { testarOrientacoesBrowser } from "./pdf-orientation-browser.mjs";

export async function testarUXBrowser({ avaliar, clicar, pronto, esperarTela, cdp, imagem, baixar, confirmar, etapa = 1 }) {
    await avaliar(`(async()=>{const {PDFDocument}=await import('/vendor/pdf-lib/pdf-lib.js');const p=await PDFDocument.create();p.addPage([600,900]);p.addPage([900,600]);
        const d=new DataTransfer();d.items.add(new File([await p.save()],'ux.pdf',{type:'application/pdf'}));const e=document.querySelector('#pdf-arquivo');e.files=d.files;e.dispatchEvent(new Event('change',{bubbles:true}));})()`);
    await pronto();
    const cabe = async () => {
        await esperarTela(`(()=>{const a=document.querySelector('#pdf-visualizacao'),c=a.querySelector('canvas');return c&&c.height<=a.clientHeight-30&&c.width<=a.clientWidth-30})()`, 'Página inteira ajustada');
        assert.equal(await avaliar(`document.documentElement.scrollHeight<=innerHeight+1`), true, 'Documento não exige rolagem da janela');
    };
    await cabe(); await imagem('16-ux-encaixe');
    await clicar('#pdf-proxima'); await esperarTela(`document.querySelector('#pdf-pagina-atual').textContent==='2 / 2'`, 'Paisagem'); await cabe();
    await clicar('#pdf-zoom-real'); await esperarTela(`document.querySelector('#pdf-visualizacao canvas').width===900`, '100%');
    const alvo = await avaliar(`(()=>{const a=document.querySelector('#pdf-visualizacao').getBoundingClientRect();return {x:a.x+a.width/2,y:a.y+a.height/2}})()`);
    const ancora = await avaliar(`(()=>{const r=document.querySelector('#pdf-visualizacao canvas').getBoundingClientRect();return {x:(${alvo.x}-r.left)/r.width,y:(${alvo.y}-r.top)/r.height}})()`);
    await cdp('Input.dispatchMouseEvent', {type:'mouseWheel', ...alvo, deltaX:0, deltaY:-180});
    await esperarTela(`document.querySelector('#pdf-visualizacao canvas').width>1000`, 'Zoom pela roda');
    const erroAncora = await avaliar(`(()=>{const r=document.querySelector('#pdf-visualizacao canvas').getBoundingClientRect();return Math.max(Math.abs(r.left+${ancora.x}*r.width-${alvo.x}),Math.abs(r.top+${ancora.y}*r.height-${alvo.y}))})()`);
    assert.ok(erroAncora < 3, 'Zoom conserva o ponto sob o cursor, salvo arredondamento do canvas');
    const antesPan = await avaliar(`document.querySelector('#pdf-visualizacao').scrollLeft`);
    await cdp('Input.dispatchMouseEvent', {type:'mousePressed', ...alvo, button:'left',clickCount:1});
    await cdp('Input.dispatchMouseEvent', {type:'mouseMoved',x:alvo.x-50,y:alvo.y-40,button:'left',buttons:1});
    await cdp('Input.dispatchMouseEvent', {type:'mouseReleased',x:alvo.x-50,y:alvo.y-40,button:'left',clickCount:1});
    assert.ok(await avaliar(`document.querySelector('#pdf-visualizacao').scrollLeft`) > antesPan, 'Arraste desloca a página ampliada');
    await clicar('#pdf-zoom-pagina'); await cabe();
    await cdp('Emulation.setDeviceMetricsOverride',{width:1024,height:600,deviceScaleFactor:1,mobile:false}); await cabe(); await imagem('17-ux-janela-pequena');
    await cdp('Emulation.setDeviceMetricsOverride',{width:1440,height:1000,deviceScaleFactor:1,mobile:false});
    if (etapa >= 2) {
        for (const total of [100, 350]) {
            await avaliar(`(async()=>{const {PDFDocument,rgb}=await import('/vendor/pdf-lib/pdf-lib.js');const p=await PDFDocument.create();for(let i=0;i<${total};i++){const a=p.addPage([400+i,600]);a.drawText('Pagina '+(i+1),{x:30,y:400});a.drawRectangle({x:20,y:20,width:80,height:60,color:rgb(.1,.4,.7)});}const d=new DataTransfer();d.items.add(new File([await p.save()],'grupo${total}.pdf'));const e=document.querySelector('#pdf-arquivo');e.files=d.files;e.dispatchEvent(new Event('change',{bubbles:true}));})()`);
            await esperarTela(`!document.querySelector('#pdf-arquivo').disabled||document.querySelector('#pdf-confirmar-descarte').open`, 'Importar grupo');
            if (await avaliar(`document.querySelector('#pdf-confirmar-descarte').open`)) await clicar('#pdf-descartar-sair');
            await pronto();
            const selecionar = async (n, modificadores = {}) => {
                await avaliar(`(()=>{const p=document.querySelector('#pdf-miniaturas');p.scrollTop=${(n - 1) * 166};p.dispatchEvent(new Event('scroll'))})()`);
                await esperarTela(`!!document.querySelector('#pdf-miniaturas button[data-pagina="${n}"]')`, 'Miniatura virtualizada');
                await avaliar(`document.querySelector('#pdf-miniaturas button[data-pagina="${n}"]').dispatchEvent(new MouseEvent('click',{bubbles:true,...${JSON.stringify(modificadores)}}))`);
                await esperarTela(`document.querySelector('#pdf-pagina-atual').textContent==='${n} / ${total}'`, 'Página ativa da seleção');
            };
            await selecionar(5); await selecionar(7, { shiftKey:true }); await selecionar(6, { ctrlKey:true }); await selecionar(9, { ctrlKey:true });
            await avaliar(`document.querySelector('#pdf-mover-destino').value='50'`); await clicar('#pdf-mover-aplicar'); await pronto();
            assert.equal(await avaliar(`document.querySelector('#pdf-pagina-atual').textContent`), `52 / ${total}`);
            assert.equal(await avaliar(`document.querySelectorAll('#pdf-miniaturas button[aria-pressed="true"]').length`), 3);
            await avaliar(`(()=>{const p=document.querySelector('#pdf-miniaturas'),b=p.querySelector('button[data-pagina="50"]');window.uxTransfer=new DataTransfer();window.uxScroll=p.scrollTop;b.dispatchEvent(new DragEvent('dragstart',{bubbles:true,dataTransfer:uxTransfer}));p.dispatchEvent(new DragEvent('dragover',{bubbles:true,dataTransfer:uxTransfer,clientY:p.getBoundingClientRect().bottom-2}));})()`);
            await esperarTela(`document.querySelector('#pdf-miniaturas').scrollTop>uxScroll+60`, 'Autoscroll do grupo');
            await avaliar(`(()=>{const p=document.querySelector('#pdf-miniaturas');p.scrollTop=0;p.dispatchEvent(new Event('scroll'));const y=p.getBoundingClientRect().top+1;p.dispatchEvent(new DragEvent('dragover',{bubbles:true,dataTransfer:uxTransfer,clientY:y}));p.dispatchEvent(new DragEvent('drop',{bubbles:true,dataTransfer:uxTransfer,clientY:y}));window.uxTransfer=null;})()`);
            await pronto(); await esperarTela(`document.querySelector('#pdf-pagina-atual').textContent==='3 / ${total}'`, 'Arraste do grupo mantém página ativa');
            assert.ok(await avaliar(`document.querySelectorAll('#pdf-miniaturas button').length`) <= 16);
            if (etapa >= 3) {
                await avaliar(`document.querySelector('#pdf-miniaturas button[data-pagina="2"]').dispatchEvent(new KeyboardEvent('keydown',{bubbles:true,key:'F10',shiftKey:true}))`);
                assert.equal(await avaliar(`document.querySelector('#pdf-paginas-menu').hidden`), false, 'Menu disponível por teclado');
                await avaliar(`document.querySelector('#pdf-duplicar').dispatchEvent(new KeyboardEvent('keydown',{bubbles:true,key:'Escape'}))`);
                assert.equal(await avaliar(`document.querySelector('#pdf-paginas-menu').hidden`), true);
                await avaliar(`document.querySelector('#pdf-miniaturas button[data-pagina="2"]').dispatchEvent(new MouseEvent('contextmenu',{bubbles:true,clientX:300,clientY:400}))`);
                assert.equal(await avaliar(`document.querySelectorAll('#pdf-miniaturas button[aria-pressed="true"]').length`), 3, 'Contexto mantém grupo que inclui a página clicada');
                await clicar('#pdf-duplicar'); await pronto();
                assert.equal(await avaliar(`document.querySelector('#pdf-total').textContent`), String(total + 3));
                assert.equal(await avaliar(`document.querySelector('#pdf-pagina-atual').textContent`), `4 / ${total + 3}`);
                await clicar('#pdf-salvar-copia'); const pdf = await baixar(`grupo${total}_editado.pdf`, total + 3);
                assert.deepEqual(pdf.getPages().slice(0, 6).map(p => p.getWidth()), [404,406,408,404,406,408]);
                const bytes = Buffer.from(await pdf.save()).toString('base64');
                const prova = await avaliar(`(async()=>{const js=await import('/vendor/pdfjs/pdf.js'),t=js.getDocument({data:Uint8Array.from(atob(${JSON.stringify(bytes)}),c=>c.charCodeAt(0))});const canvases=[];try{const d=await t.promise, resultados=[];for(const n of [1,4]){const p=await d.getPage(n),v=p.getViewport({scale:1}),c=document.createElement('canvas');canvases.push(c);c.width=Math.ceil(v.width);c.height=Math.ceil(v.height);await p.render({canvas:c,canvasContext:c.getContext('2d'),viewport:v}).promise;resultados.push({pixels:c.getContext('2d').getImageData(0,0,c.width,c.height).data,texto:(await p.getTextContent()).items.map(i=>i.str).join('')});}return {texto:resultados[0].texto===resultados[1].texto,pixels:resultados[0].pixels.every((v,i)=>v===resultados[1].pixels[i])};}finally{canvases.forEach(c=>c.width=c.height=0);await t.destroy();}})()`);
                assert.deepEqual(prova, { texto:true,pixels:true }, 'Cópia mantém texto extraível e aparência vetorial');
                await clicar('#pdf-paginas-acoes'); await confirmar('#pdf-excluir-selecao', false); await pronto();
                assert.equal(await avaliar(`document.querySelector('#pdf-total').textContent`), String(total + 3));
                await clicar('#pdf-paginas-acoes'); await confirmar('#pdf-excluir-selecao', true); await pronto();
                assert.equal(await avaliar(`document.querySelector('#pdf-total').textContent`), String(total));
            }
        }
    }
    await clicar('[data-close-tool="pdfs"]');
    await esperarTela(`!document.querySelector('#pdf-ferramenta')||document.querySelector('#pdf-confirmar-descarte').open`, 'Encerrar UX');
    if (await avaliar(`!!document.querySelector('#pdf-confirmar-descarte')?.open`)) await clicar('#pdf-descartar-sair');
    await esperarTela(`!document.querySelector('#pdf-ferramenta')`, 'Encerrar UX confirmado');
    await clicar('[data-tool="pdfs"]'); await esperarTela(`!!document.querySelector('#pdf-arquivo')`, 'Retornar PDFs');
    if (etapa >= 4) {
        const antes = await avaliar(`Number(document.querySelector('#pdf-divisor-direita').getAttribute('aria-valuenow'))`);
        await avaliar(`document.querySelector('#pdf-divisor-direita').dispatchEvent(new KeyboardEvent('keydown',{bubbles:true,key:'ArrowLeft'}))`);
        assert.equal(await avaliar(`Number(document.querySelector('#pdf-divisor-direita').getAttribute('aria-valuenow'))`), antes + 10);
        const ponto = await avaliar(`(()=>{const r=document.querySelector('#pdf-divisor-direita').getBoundingClientRect();return {x:r.x+r.width/2,y:r.y+r.height/2}})()`);
        await cdp('Input.dispatchMouseEvent',{type:'mousePressed',...ponto,button:'left',clickCount:1});
        await cdp('Input.dispatchMouseEvent',{type:'mouseMoved',x:ponto.x-30,y:ponto.y,button:'left',buttons:1});
        await cdp('Input.dispatchMouseEvent',{type:'mouseReleased',x:ponto.x-30,y:ponto.y,button:'left',clickCount:1});
        const salvo = await avaliar(`JSON.parse(localStorage.getItem('4printshops.pdf.paineis.v1')).direita`);
        assert.equal(salvo, antes + 40);
        await cdp('Page.reload'); await esperarTela(`!!document.querySelector('[data-tool="pdfs"]')`, 'Reabrir sistema');
        await clicar('[data-tool="pdfs"]'); await esperarTela(`!!document.querySelector('#pdf-divisor-direita')`, 'Preferência recuperada');
        assert.equal(await avaliar(`Number(document.querySelector('#pdf-divisor-direita').getAttribute('aria-valuenow'))`), salvo);
        await avaliar(`localStorage.removeItem('4printshops.pdf.paineis.v1')`);
        await clicar('[data-close-tool="pdfs"]'); await esperarTela(`!document.querySelector('#pdf-ferramenta')`, 'Repor padrões');
        await clicar('[data-tool="pdfs"]'); await esperarTela(`!!document.querySelector('#pdf-arquivo')`, 'PDFs padrão');
    }
    if (etapa >= 5) {
        await testarOrientacoesBrowser({ avaliar, clicar, pronto, esperarTela, imagem, baixar });
        await avaliar(`(async()=>{const {PDFDocument,degrees}=await import('/vendor/pdf-lib/pdf-lib.js'),p=await PDFDocument.create();for(let i=0;i<350;i++){const a=p.addPage([210*72/25.4,297*72/25.4]);if(i===349)a.setRotation(degrees(90));}p.addPage([297*72/25.4,420*72/25.4]);p.addPage([297*72/25.4,420*72/25.4]);const d=new DataTransfer();d.items.add(new File([await p.save()],'dimensoes.pdf'));const e=document.querySelector('#pdf-arquivo');e.files=d.files;e.dispatchEvent(new Event('change',{bubbles:true}));})()`);
        await pronto();
        const topo = await avaliar(`document.querySelector('#pdf-visualizacao').getBoundingClientRect().top`);
        await clicar('#pdf-informacoes');
        assert.equal(await avaliar(`document.querySelectorAll('#pdf-lista-paginas > li').length`), 3);
        assert.match(await avaliar(`document.querySelector('#pdf-lista-paginas summary').textContent`), /A4.*349 página/);
        await clicar('#pdf-lista-paginas summary');
        assert.equal(await avaliar(`document.querySelector('#pdf-lista-paginas details').open`), true);
        assert.match(await avaliar(`document.querySelector('#pdf-lista-paginas p').textContent`), /1–349/);
        assert.equal(await avaliar(`document.querySelector('#pdf-visualizacao').getBoundingClientRect().top`), topo, 'Informações não desloca a prévia');
        await imagem('18-ux-dimensoes-agrupadas');
        await clicar('[data-close-tool="pdfs"]'); await esperarTela(`!document.querySelector('#pdf-ferramenta')`, 'Fechar dimensões');
        await clicar('[data-tool="pdfs"]'); await esperarTela(`!!document.querySelector('#pdf-arquivo')`, 'PDFs após grupos');
    }
    if (etapa >= 6) {
        await clicar('#pdf-ferramenta-redimensionar');
        assert.equal(await avaliar(`document.querySelectorAll('#pdf-redim-vincular').length`), 1);
        assert.equal(await avaliar(`document.querySelector('#pdf-redim-vincular').getAttribute('aria-label')`), 'Proporção livre');
        assert.equal(await avaliar(`getComputedStyle(document.querySelector('#pdf-redim-vincular [data-icon="lock"]')).display`), 'none');
        assert.notEqual(await avaliar(`getComputedStyle(document.querySelector('#pdf-redim-vincular [data-icon="lock-open"]')).display`), 'none');
        assert.equal(await avaliar(`document.querySelector('#pdf-redim-proporcao').textContent`), '210 × 297 mm · ≈ 1:1,41');
        await avaliar(`document.querySelector('#pdf-divisor-direita').dispatchEvent(new KeyboardEvent('keydown',{bubbles:true,key:'Home'}))`);
        const alinhado = await avaliar(`(()=>{const l=document.querySelector('#pdf-redim-largura').getBoundingClientRect(),c=document.querySelector('#pdf-redim-vincular').getBoundingClientRect(),a=document.querySelector('#pdf-redim-altura').getBoundingClientRect(),p=document.querySelector('#pdf-painel-ferramentas'),labels=[...document.querySelectorAll('.pdf-redim-dimensoes label')].map(e=>e.getBoundingClientRect());return l.right<=c.left&&c.right<=a.left&&Math.abs(l.top-c.top)<1&&Math.abs(a.top-c.top)<1&&Math.abs(labels[0].top-labels[1].top)<1&&p.scrollWidth<=p.clientWidth;})()`);
        assert.equal(alinhado, true, 'Campos e cadeado alinhados no painel mínimo, sem overflow horizontal');
        await imagem('20-ux-cadeado-livre');
        await clicar('#pdf-redim-vincular');
        assert.equal(await avaliar(`document.querySelector('#pdf-redim-vincular').getAttribute('aria-label')`), 'Proporção vinculada');
        assert.notEqual(await avaliar(`getComputedStyle(document.querySelector('#pdf-redim-vincular [data-icon="lock"]')).display`), 'none');
        assert.equal(await avaliar(`getComputedStyle(document.querySelector('#pdf-redim-vincular [data-icon="lock-open"]')).display`), 'none');
        await imagem('21-ux-cadeado-vinculado');
        const entrada = async (id, valor, tipo = 'input') => avaliar(`(()=>{const e=document.querySelector(${JSON.stringify(id)});e.value=${JSON.stringify(valor)};e.dispatchEvent(new Event(${JSON.stringify(tipo)},{bubbles:true}));})()`);
        await entrada('#pdf-redim-largura','420'); assert.equal(await avaliar(`document.querySelector('#pdf-redim-altura').value`),'594');
        await entrada('#pdf-redim-orientacao','paisagem','change'); assert.equal(await avaliar(`document.querySelector('#pdf-redim-largura').value`),'594');
        await entrada('#pdf-redim-formato','A5','change'); assert.equal(await avaliar(`document.querySelector('#pdf-redim-altura').value`),'210');
        await clicar('#pdf-redim-vincular'); await entrada('#pdf-redim-largura','123,4567890123'); await entrada('#pdf-redim-altura','234,5678901234');
        await entrada('#pdf-formato-nome','UX preciso'); await clicar('#pdf-formato-salvar');
        await esperarTela(`document.querySelector('#pdf-redim-formato').value.startsWith('custom-')`, 'Formato preciso salvo');
        const id = await avaliar(`document.querySelector('#pdf-redim-formato').value`);
        await entrada('#pdf-redim-formato','A4','change'); await entrada('#pdf-redim-formato',id,'change');
        assert.equal(await avaliar(`document.querySelector('#pdf-redim-largura').value`),'123,4567890123');
        assert.equal(await avaliar(`document.querySelector('#pdf-redim-altura').value`),'234,5678901234');
        for (const [modo, descricao] of [['ajustar','Cabe inteiro; pode sobrar margem.'],['centralizar','Mantém o tamanho original do conteúdo.'],['preencher','Preenche toda a folha; pode cortar bordas.']]) {
            await entrada('#pdf-redim-modo',modo,'change'); assert.equal(await avaliar(`document.querySelector('#pdf-redim-modo-descricao').textContent`), descricao);
        }
        await avaliar(`document.querySelector('#pdf-redim-largura').scrollIntoView({block:'center'})`);
        await imagem('19-ux-proporcao'); await avaliar(`localStorage.removeItem('4printshops.pdf.formatos.v1');localStorage.removeItem('4printshops.pdf.paineis.v1')`);
        await clicar('[data-close-tool="pdfs"]'); await esperarTela(`!document.querySelector('#pdf-ferramenta')`, 'Fechar proporção');
        await clicar('[data-tool="pdfs"]'); await esperarTela(`!!document.querySelector('#pdf-arquivo')`, 'PDFs após proporção');
    }
    console.log('PASSOU: UX etapa '+etapa+', visualização; seleção/movimento/ações em 100/350 conforme etapa; painéis, grupos e proporção.');
}
