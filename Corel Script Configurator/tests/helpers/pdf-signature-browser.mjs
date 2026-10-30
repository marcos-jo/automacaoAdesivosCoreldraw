import assert from "node:assert/strict";
import { PDFDocument, PDFName } from "../../public/vendor/pdf-lib/pdf-lib.js";
import { unlink, readFile } from "node:fs/promises";
import path from "node:path";
import { criarPDFAssinado, adicionarAssinatura } from "./pdf-signature-fixture.mjs";
import { criarFixtureImpressao, textoImpressao } from "./pdf-print-fixture.mjs";
import { criarConjuntoAssinaturasReais, verificarAssinaturasCMS } from "./pdf-real-signatures.mjs";
import { createHash } from "node:crypto";
import { assinaturasUrl } from "./pdf-resize-modules.mjs";
const { verificarSemAssinaturas } = await import(assinaturasUrl);
const nome = PDFName.of;

export async function testarAssinaturasBrowser({ avaliar, clicar, pronto, aguardar, esperarTela, confirmar, baixar, imagem, temporario }) {
    await avaliar(`(async()=>{
        window.downloadsImpressao=0;const clique=HTMLAnchorElement.prototype.click;HTMLAnchorElement.prototype.click=function(){if(this.download)downloadsImpressao++;return clique.call(this);};
        const lib=await import('/vendor/pdf-lib/pdf-lib.js'),pdf=await lib.PDFDocument.create();pdf.addPage([100,100]);
        const js=await import('/vendor/pdfjs/pdf.js');js.GlobalWorkerOptions.workerSrc='/vendor/pdfjs/pdf.worker.js';
        const t=js.getDocument({data:await pdf.save()});try{const p=await(await t.promise).getPage(1),proto=Object.getPrototypeOf(p),render=proto.render;
            proto.render=function(o){if(o.intent==='print'&&o.canvas.className==='pdf-pagina-canvas')window.escalaPreviaImpressao=o.viewport.scale;
                if(window.falharImpressao&&o.intent==='print'){window.falharImpressao=false;return {promise:Promise.reject(new Error('Falha de render injetada no teste')),cancel(){}};}return render.call(this,o);};
        }finally{await t.destroy();}
    })()`);
    const importar = async (bytes, arquivo) => {
        await avaliar(`(()=>{const d=new DataTransfer();d.items.add(new File([Uint8Array.from(atob(${JSON.stringify(Buffer.from(bytes).toString("base64"))}),c=>c.charCodeAt(0))],${JSON.stringify(arquivo)},{type:'application/pdf'}));const e=document.querySelector('#pdf-arquivo');e.files=d.files;e.dispatchEvent(new Event('change',{bubbles:true}));})()`);
        await esperarTela(`!document.querySelector('#pdf-arquivo').disabled||document.querySelector('#pdf-confirmar-descarte').open`, "Importação/confirmar descarte");
        if (await avaliar(`document.querySelector('#pdf-confirmar-descarte').open`)) await clicar("#pdf-descartar-sair");
        await pronto();
    };
    const abrir = async () => {
        await clicar("#pdf-assinaturas-resolver"); await esperarTela(`document.querySelector('#pdf-impressao-dialogo').open`, "Preparo aberto");
        assert.equal(await avaliar(`document.querySelector('#pdf-impressao-detalhes').open`), false, "explicações recolhidas a cada abertura");
        assert.equal(await avaliar(`document.querySelector('#pdf-impressao-controles-revisao').hidden`), true);
        assert.equal(await avaliar(`document.querySelector('#pdf-impressao-resultado-tecnico').textContent`), "");
    };
    const preparar = async (sucesso = true) => {
        await clicar("#pdf-impressao-prever");
        const texto = await aguardar(async () => avaliar(`(()=>{const t=document.querySelector('#pdf-impressao-progresso').textContent;return !document.querySelector('#pdf-impressao-prever').disabled&&t&&!/^(Preparando|Renderizando|Conferindo)/.test(t)?t:false})()`), "Proposta finalizada", 120000);
        const pronta = await avaliar(`!document.querySelector('#pdf-impressao-conferi').disabled`);
        assert.equal(pronta, sucesso, texto); return texto;
    };
    const fechar = async () => { await clicar("#pdf-impressao-cancelar"); await esperarTela(`!document.querySelector('#pdf-impressao-dialogo').open`, "Resolver fechado"); await pronto(); };
    const marcarConferida = async () => { if (!await avaliar(`document.querySelector('#pdf-impressao-conferi').checked`)) await clicar("#pdf-impressao-conferi"); };
    const conferirRestantes = async () => {
        while (!await avaliar(`document.querySelector('#pdf-impressao-proxima').disabled`)) {
            const anterior = await avaliar(`document.querySelector('#pdf-impressao-revisadas').textContent`);
            await clicar("#pdf-impressao-proxima");
            await esperarTela(`document.querySelector('#pdf-impressao-revisadas').textContent!==${JSON.stringify(anterior)}`, "Conferir próxima página");
        }
        await marcarConferida();
    };
    const aplicarExportar = async (nomeArquivo, total) => {
        await conferirRestantes();
        const antes = await avaliar('downloadsImpressao');
        await avaliar(`(()=>{const c=document.querySelector('#pdf-impressao-previa canvas');window.capturaProposta={w:c.width,h:c.height,n:Number(document.querySelector('#pdf-impressao-pagina').value),escala:escalaPreviaImpressao,pixels:c.getContext('2d').getImageData(0,0,c.width,c.height).data};})()`);
        await confirmar(true);
        await esperarTela(`!document.querySelector('#pdf-impressao-dialogo').open`, "Aplicar cópia"); await pronto();
        assert.equal(await avaliar(`document.querySelector('#pdf-alteracoes-pendentes').hidden`), false);
        assert.equal(await avaliar('downloadsImpressao'), antes, 'Resolver/Aplicar não dispara download');
        try { await unlink(path.join(temporario, nomeArquivo)); } catch (e) { if (e.code !== 'ENOENT') throw e; }
        await clicar("#pdf-salvar-copia"); const pdf = await baixar(nomeArquivo, total);
        verificarSemAssinaturas(pdf);
        const b64 = (await readFile(path.join(temporario, nomeArquivo))).toString('base64');
        const diferencas = await avaliar(`(async()=>{const js=await import('/vendor/pdfjs/pdf.js'),t=js.getDocument({data:Uint8Array.from(atob(${JSON.stringify(b64)}),c=>c.charCodeAt(0))});
            let c;try{const p=await(await t.promise).getPage(capturaProposta.n),v=p.getViewport({scale:capturaProposta.escala});c=document.createElement('canvas');c.width=capturaProposta.w;c.height=capturaProposta.h;
                await p.render({canvas:c,canvasContext:c.getContext('2d'),viewport:v,intent:'print',annotationMode:js.AnnotationMode.ENABLE,background:'rgb(255,255,255)'}).promise;
                return c.getContext('2d').getImageData(0,0,c.width,c.height).data.reduce((n,v,i)=>n+(v!==capturaProposta.pixels[i]),0);
            }finally{if(c)c.width=c.height=0;await t.destroy();}})()`);
        assert.equal(diferencas,0,'Pixels da prévia proposta e do PDF exportado devem coincidir');
        assert.equal(await avaliar(`document.querySelector('#pdf-alteracoes-pendentes').hidden`), true);
        return pdf;
    };

    // Uma assinatura somente na última página, sem rasterizar as primeiras.
    await importar(await criarPDFAssinado(), "assinada.pdf");
    assert.equal(await avaliar(`document.querySelector('#pdf-assinaturas-titulo').textContent`), "Assinatura digital detectada");
    await clicar('#pdf-paginas-acoes'); await clicar('#pdf-duplicar'); await pronto();
    assert.match(await avaliar(`document.querySelector('#pdf-status').textContent`), /Duplicação bloqueada.*assinatura/);
    assert.equal(await avaliar(`document.querySelector('#pdf-total').textContent`), '3');
    assert.equal(await avaliar(`document.querySelector('#pdf-alteracoes-pendentes').hidden`), true);
    assert.match(await avaliar(`document.querySelector('#pdf-assinaturas-resumo').textContent`), /visuais atuais: 3/);
    assert.equal(await avaliar(`document.querySelector('#pdf-assinaturas-resolver').textContent`), "Resolver");
    assert.equal(await avaliar(`document.querySelector('#pdf-assinaturas-expandido').hidden`), true);
    await avaliar(`document.querySelector('#pdf-assinaturas-alerta').scrollIntoView({block:'center'})`);
    await imagem("13-aviso-assinatura-compacto");
    await clicar('#pdf-assinaturas-expandir');
    assert.equal(await avaliar(`document.querySelector('#pdf-assinaturas-expandir').getAttribute('aria-expanded')`), 'true');
    assert.equal(await avaliar(`document.querySelector('#pdf-assinaturas-expandido').textContent===document.querySelector('#pdf-assinaturas-detalhes').textContent`), true);
    await clicar('#pdf-assinaturas-expandir');
    assert.equal(await avaliar(`document.querySelector('#pdf-assinaturas-expandido').hidden`), true);
    await clicar('#pdf-assinaturas-expandir'); // Aplicar deverá recolher o histórico da origem.
    await clicar("#pdf-informacoes");
    assert.match(await avaliar(`document.querySelector('#pdf-assinaturas-detalhes').textContent`), /antes de reserializar/);
    await abrir(); assert.equal(await avaliar(`document.querySelector('#pdf-impressao-dpi').value`), "300");
    await imagem('14-modal-impressao-compacto');
    await clicar('#pdf-impressao-detalhes-alternar');
    assert.equal(await avaliar(`document.querySelector('#pdf-impressao-detalhes').open`), true);
    assert.match(await avaliar(`document.querySelector('#pdf-impressao-diagnostico').textContent`), /antes de reserializar/);
    await clicar('#pdf-impressao-detalhes-alternar');
    assert.equal(await avaliar(`document.querySelector('#pdf-impressao-detalhes').open`), false);
    await fechar(); assert.equal(await avaliar(`document.querySelector('#pdf-alteracoes-pendentes').hidden`), true);
    await abrir(); await preparar();
    assert.match(await avaliar(`document.querySelector('#pdf-impressao-progresso').textContent`), /Prévia pronta/);
    assert.match(await avaliar(`document.querySelector('#pdf-impressao-resultado-tecnico').textContent`), /Rasterizadas: 3\. Sem nova rasterização: 2/);
    assert.equal(await avaliar(`document.querySelector('#pdf-impressao-controles-revisao').hidden`), false);
    await marcarConferida(); await confirmar(false);
    assert.equal(await avaliar(`document.querySelector('#pdf-alteracoes-pendentes').hidden`), true);
    await clicar("#pdf-impressao-origem");
    await esperarTela(`document.querySelector('#pdf-impressao-revisadas').textContent.startsWith('Antes do preparo')`, "Workspace antes do preparo com AP");
    await clicar("#pdf-impressao-origem");
    await esperarTela(`document.querySelector('#pdf-impressao-revisadas').textContent.startsWith('Cópia')`, "Cópia de impressão");
    await avaliar(`document.querySelector('#pdf-impressao-previa').scrollIntoView({block:'center'})`);
    await imagem("12-assinaturas-resolver");
    const seletivo = await aplicarExportar("assinada_IMPRESSAO.pdf", 3);
    assert.ok(seletivo.getPage(0).node.Resources().lookup(nome("Font")).keys().length > 0);
    assert.equal(seletivo.getPage(2).node.Resources().lookup(nome("Font"))?.keys().length || 0, 0);
    assert.match(await avaliar(`document.querySelector('#pdf-assinaturas-detalhes').textContent`), /Origem da cópia/);
    assert.equal(await avaliar(`document.querySelector('#pdf-assinaturas-titulo').textContent`), "Cópia para impressão preparada");
    assert.match(await avaliar(`document.querySelector('#pdf-assinaturas-resumo').textContent`), /^Sem assinatura ativa/);
    assert.equal(await avaliar(`document.querySelector('#pdf-assinaturas-resolver').textContent`), "Refazer preparo");
    assert.equal(await avaliar(`document.querySelector('#pdf-assinaturas-expandido').hidden`), true);
    await avaliar(`document.querySelector('#pdf-assinaturas-alerta').scrollIntoView({block:'center'})`);
    await imagem('15-copia-preparada-compacta');
    await clicar('#pdf-assinaturas-expandir');
    assert.match(await avaliar(`document.querySelector('#pdf-assinaturas-expandido').textContent`), /Histórico e diagnóstico das origens/);
    await abrir();
    assert.equal(await avaliar(`document.querySelector('#pdf-impressao-titulo').textContent`), 'Refazer preparo para impressão');
    assert.match(await avaliar(`document.querySelector('#pdf-impressao-contexto').textContent`), /Cópia atual sem assinatura ativa/);
    await clicar('#pdf-impressao-detalhes-alternar');
    assert.match(await avaliar(`document.querySelector('#pdf-impressao-diagnostico').textContent`), /não recupera detalhes perdidos/);
    await fechar();
    assert.equal(await avaliar(`document.querySelector('#pdf-assinaturas-resolver').textContent`), 'Refazer preparo');
    assert.equal(await avaliar(`document.querySelector('#pdf-alteracoes-pendentes').hidden`), true, 'cancelar refazer mantém a versão exportada');
    await importar(await seletivo.save(), "reaberto_IMPRESSAO.pdf");
    assert.equal(await avaliar(`document.querySelector('#pdf-assinaturas-alerta').hidden`), true);

    // Duas assinaturas, anotações e dimensões/orientações mistas. Exige conferir ambas.
    await importar(await criarPDFAssinado({ paginas: [1, 3], anotacao: true, rotacao: 90, tamanhos: [[300, 400], [400, 500], [500, 300]] }), "duas.pdf");
    await abrir(); await preparar(); await marcarConferida();
    assert.equal(await avaliar(`document.querySelector('#pdf-impressao-aplicar').disabled`), true);
    const duas = await aplicarExportar("duas_IMPRESSAO.pdf", 3);
    assert.deepEqual(duas.getPages().map(p => [p.getWidth(), p.getHeight(), p.getRotation().angle]), [[400, 300, 0], [400, 500, 90], [300, 500, 0]]);
    // Operações posteriores continuam acumuladas. Nova união assinada reativa o alerta.
    await clicar("#pdf-girar-direita"); await pronto();
    await clicar("#pdf-ferramenta-organizar");
    await avaliar(`document.querySelector('#pdf-selecao-paginas').value='2'`); await clicar("#pdf-excluir"); await pronto();
    await avaliar(`document.querySelector('#pdf-mover-destino').value='2'`); await clicar("#pdf-mover-aplicar"); await pronto();
    await unlink(path.join(temporario, "duas_IMPRESSAO.pdf"));
    await clicar("#pdf-salvar-copia"); await baixar("duas_IMPRESSAO.pdf", 2);
    await avaliar(`document.querySelector('#pdf-selecao-paginas').value='1'`); await clicar('#pdf-extrair'); await pronto();
    verificarSemAssinaturas(await baixar('duas_IMPRESSAO_extraidas.pdf',1));
    await clicar('#pdf-ferramenta-redimensionar');
    await avaliar(`document.querySelector('#pdf-redim-largura').value='100';document.querySelector('#pdf-redim-altura').value='150';document.querySelector('#pdf-redim-alvo').value='todas';document.querySelector('#pdf-redim-alvo').dispatchEvent(new Event('change',{bubbles:true}))`);
    await clicar('#pdf-redim-prever');await esperarTela(`!document.querySelector('#pdf-redim-aplicar').disabled`,'Redimensionar depois de Resolver');await clicar('#pdf-redim-aplicar');await pronto();
    await clicar('#pdf-ferramenta-unir');await clicar('#pdf-uniao-abrir');
    const nova = await criarPDFAssinado({total:1});
    await avaliar(`(()=>{const d=new DataTransfer();d.items.add(new File([Uint8Array.from(atob(${JSON.stringify(Buffer.from(nova).toString('base64'))}),c=>c.charCodeAt(0))],'nova-assinada.pdf',{type:'application/pdf'}));const e=document.querySelector('#pdf-uniao-arquivos');e.files=d.files;e.dispatchEvent(new Event('change',{bubbles:true}));})()`);
    await clicar('#pdf-uniao-gerar');await pronto();
    assert.equal(await avaliar(`document.querySelector('#pdf-assinaturas-titulo').textContent`),'Assinatura digital detectada');
    assert.equal(await avaliar(`document.querySelector('#pdf-assinaturas-resolver').textContent`), 'Resolver');
    await abrir();await preparar();await fechar();
    assert.equal(await avaliar(`document.querySelector('#pdf-total').textContent`),'3');
    assert.equal(await avaliar(`document.querySelector('#pdf-alteracoes-pendentes').hidden`),false,'Cancelar não perde união/redimensionamento anteriores');
    await abrir();await preparar();await aplicarExportar('duas_IMPRESSAO.pdf',3);

    // Assinatura invisível: nada de supor uma página afetada.
    await importar(await criarPDFAssinado({ invisivel: true }), "invisivel.pdf"); await abrir();
    assert.match(await avaliar(`document.querySelector('#pdf-impressao-contexto').textContent`), /nenhuma determinada/);
    assert.match(await preparar(false), /manualmente|manuais/);
    await avaliar(`document.querySelector('#pdf-impressao-manual').value='2';document.querySelector('#pdf-impressao-manual').dispatchEvent(new Event('input',{bubbles:true}))`);
    await preparar(); await aplicarExportar("invisivel_IMPRESSAO.pdf", 3);

    await importar(await criarPDFAssinado({ vazia: true }), "vazia.pdf");
    assert.equal(await avaliar(`document.querySelector('#pdf-assinaturas-titulo').textContent`), "Campo de assinatura vazio");
    assert.match(await avaliar(`document.querySelector('#pdf-assinaturas-resumo').textContent`), /^0 assinatura/);
    await abrir(); await preparar(); await fechar();
    for (const [tipo, opcoes] of [["sem-ap", { semAparencia: true }], ["sem-print", { imprimivel: false }]]) {
        await importar(await criarPDFAssinado(opcoes), tipo + ".pdf"); await abrir();
        assert.match(await preparar(false), /sem aparência|sem flag/);
        await fechar(); assert.equal(await avaliar(`document.querySelector('#pdf-alteracoes-pendentes').hidden`), true);
    }
    // Catálogo com camadas: seletivo bloqueia, reconstrução inteira é uma alternativa explícita.
    const camadas = await PDFDocument.load(await criarPDFAssinado());
    camadas.catalog.set(nome("OCProperties"), camadas.context.obj({ OCGs: [], D: { Order: [] } }));
    await importar(await camadas.save(), "camadas.pdf"); await abrir();
    assert.match(await preparar(false), /documento inteiro/);
    await avaliar(`document.querySelector('#pdf-impressao-modo').value='inteiro';document.querySelector('#pdf-impressao-modo').dispatchEvent(new Event('change',{bubbles:true}))`);
    await preparar();
    const inteira = await aplicarExportar("camadas_IMPRESSAO.pdf", 3);
    assert.ok(inteira.getPages().every(p => !p.node.Resources().lookup(nome("Font"))?.keys().length));

    // AP formalmente existente mas vazia: falha de renderização não pode virar aprovação.
    const semPixels = await PDFDocument.load(await criarPDFAssinado());
    const widget = semPixels.context.lookup(semPixels.getPage(2).node.Annots().get(0));
    widget.set(nome("AP"), semPixels.context.obj({ N: semPixels.context.register(semPixels.context.flateStream("", { Type: "XObject", Subtype: "Form", BBox: [0, 0, 180, 50], Resources: {} })) }));
    await importar(await semPixels.save(), "sem-pixels.pdf"); await abrir();
    assert.match(await preparar(false), /aparência|pixels/); await fechar();

    // Fonte com origem deslocada, UserUnit e widget NoRotate em página girada.
    const deslocada = await PDFDocument.load(await criarPDFAssinado({ total: 1, rotacao: 270 }));
    const pd = deslocada.getPage(0); pd.setMediaBox(-10, -20, 300, 400); pd.setCropBox(0, -10, 280, 380); pd.setTrimBox(5, 0, 260, 350);
    pd.node.set(nome("UserUnit"), deslocada.context.obj(1.5));
    deslocada.context.lookup(pd.node.Annots().get(0)).set(nome("F"), deslocada.context.obj(4 | 16));
    await importar(await deslocada.save(), "deslocada.pdf"); await abrir(); await preparar();
    const normalizada = await aplicarExportar("deslocada_IMPRESSAO.pdf", 1);
    assert.equal(normalizada.getPage(0).getRotation().angle, 0);
    assert.ok(Math.abs(normalizada.getPage(0).getCropBox().width - 570) < 1e-6);
    assert.ok(Math.abs(normalizada.getPage(0).getCropBox().height - 420) < 1e-6);
    // Falha assíncrona do render libera a interface e permite tentar novamente.
    await importar(await criarPDFAssinado(), 'falha-render.pdf');await abrir();await avaliar('window.falharImpressao=true');
    assert.match(await preparar(false),/Falha de render/);
    assert.equal(await avaliar(`document.querySelector('#pdf-impressao-aplicar').disabled`),true);
    await preparar();await fechar();

    // Fonte TrueType incorporada, símbolos, transparência e PNG 1024². Uma
    // segunda página conserva os mesmos recursos gráficos sem nova rasterização.
    const complexo = await PDFDocument.load(await criarFixtureImpressao({ actualText: textoImpressao }));
    const png = await avaliar(`(()=>{const c=document.createElement('canvas');c.width=c.height=1024;const x=c.getContext('2d'),g=x.createLinearGradient(0,0,1024,1024);g.addColorStop(0,'rgba(255,0,100,.4)');g.addColorStop(1,'rgba(0,70,255,.8)');x.fillStyle=g;x.fillRect(0,0,1024,1024);const s=c.toDataURL();c.width=c.height=0;return s;})()`);
    complexo.getPage(0).drawImage(await complexo.embedPng(png), {x:650,y:270,width:200,height:160});
    const embutida = await complexo.embedPage(complexo.getPage(0));complexo.addPage([960,540]).drawPage(embutida);
    adicionarAssinatura(complexo,{pagina:1});
    await importar(await complexo.save({updateFieldAppearances:false}),'fontes-imagens.pdf');await abrir();await preparar();
    await aplicarExportar('fontes-imagens_IMPRESSAO.pdf',2);
    const conteudoVetorial = (await readFile(path.join(temporario,'fontes-imagens_IMPRESSAO.pdf'))).toString('base64');
    assert.equal(await avaliar(`(async()=>{const j=await import('/vendor/pdfjs/pdf.js'),t=j.getDocument({data:Uint8Array.from(atob(${JSON.stringify(conteudoVetorial)}),c=>c.charCodeAt(0))});try{return (await(await(await t.promise).getPage(2)).getTextContent()).items.map(i=>i.str).join(' ').includes('Gráfica');}finally{await t.destroy();}})()`),true,'Texto incorporado continua extraível na página preservada');

    // Memória/resolução, render e reabertura real em A4/A3, sem baixar automaticamente.
    for (const [w, h, dpi] of [[210, 297, 300], [297, 420, 300], [210, 297, 600], [297, 420, 600]]) {
        const id = `mem-${w}-${dpi}`, bytes = await criarPDFAssinado({ total: 1, tamanhos: [[w * 72 / 25.4, h * 72 / 25.4]] });
        await importar(bytes, id + ".pdf"); await abrir();
        await avaliar(`document.querySelector('#pdf-impressao-dpi').value=${JSON.stringify(String(dpi))};document.querySelector('#pdf-impressao-dpi').dispatchEvent(new Event('change',{bubbles:true}))`);
        if (w === 297 && dpi === 600) { assert.match(await preparar(false), /36 milhões/); await fechar(); continue; }
        await avaliar(`window.quadrosImpressao=0;window.timerImpressao=setInterval(()=>quadrosImpressao++,10)`);
        await preparar();
        const quadros = await avaliar(`clearInterval(timerImpressao);quadrosImpressao`); assert.ok(quadros > 0);
        const resultado = await aplicarExportar(id + "_IMPRESSAO.pdf", 1);
        const recursos = resultado.getPage(0).node.Resources().lookup(nome("XObject"));
        const img = resultado.context.lookup(recursos.values()[0]);
        assert.equal(img.dict.lookup(nome("Width")).asNumber(), Math.ceil(w / 25.4 * dpi));
        assert.equal(img.dict.lookup(nome("Height")).asNumber(), Math.ceil(h / 25.4 * dpi));
        console.log(`Raster ${w}×${h} mm / ${dpi} DPI: ${img.dict.lookup(nome("Width"))}×${img.dict.lookup(nome("Height"))}; canvas ${(img.dict.lookup(nome("Width")).asNumber() * img.dict.lookup(nome("Height")).asNumber() * 4 / 1048576).toFixed(1)} MiB; eventos executados ${quadros}.`);
    }
    // Cancelamento em documento extenso; nenhuma geração obsoleta pode reaparecer.
    await importar(await criarPDFAssinado({ total: 350 }), "grande-assinada.pdf"); await abrir();
    await clicar("#pdf-impressao-prever");
    await esperarTela(`document.querySelector('#pdf-impressao-prever').disabled`, "Preparo grande em andamento");
    await fechar();
    assert.equal(await avaliar(`document.querySelector('#pdf-total').textContent`), "350");
    assert.equal(await avaliar(`document.querySelector('#pdf-alteracoes-pendentes').hidden`), true);
    await clicar("#pdf-girar-direita"); await pronto();
    assert.equal(await avaliar(`document.querySelector('#pdf-alteracoes-pendentes').hidden`), false);
    // Segunda bateria: CMS reais, certificado autoassinado de teste e revisões
    // incrementais. OpenSSL verifica os originais ANTES de o app reserializar.
    const reais = await criarConjuntoAssinaturasReais();
    for (const caso of reais.casos) {
        assert.equal(caso.verificacoes.length, caso.vazio ? 0 : caso.paginas.length);
        await importar(caso.bytes, caso.id + '.pdf');
        assert.equal(await avaliar(`document.querySelector('#pdf-assinaturas-titulo').textContent`), caso.vazio ? 'Campo de assinatura vazio' : 'Assinatura digital detectada');
        await abrir();
        if (caso.invisivel) {
            assert.match(await preparar(false), /manualmente|manuais/);
            await avaliar(`document.querySelector('#pdf-impressao-manual').value=${JSON.stringify(String(caso.paginas[0]))};document.querySelector('#pdf-impressao-manual').dispatchEvent(new Event('input',{bubbles:true}))`);
        }
        await preparar();
        const exportado = await aplicarExportar(caso.id + '_IMPRESSAO.pdf', caso.total);
        const rasterizadas = new Set(caso.paginas);
        for (let i = 0; i < caso.total; i++) {
            const p = exportado.getPage(i), fontes = p.node.Resources().lookup(nome('Font'))?.keys().length || 0;
            if (rasterizadas.has(i + 1)) assert.equal(fontes, 0, 'Página necessária deve ter sido rasterizada');
            else assert.ok(fontes > 0, 'Página não afetada deve preservar fonte/conteúdo vetorial');
        }
        const bytesExportados = await readFile(path.join(temporario, caso.id + '_IMPRESSAO.pdf'));
        assert.equal((await verificarAssinaturasCMS(bytesExportados, reais, caso.id + '-saida')).length, 0);
        assert.equal(createHash('sha256').update(await readFile(caso.arquivo)).digest('hex'), caso.sha256);
        assert.equal((await verificarAssinaturasCMS(await readFile(caso.arquivo), reais, caso.id + '-apos-resolver')).length, caso.verificacoes.length);
        // Reabertura efetiva dos bytes baixados, não apenas do candidato em memória.
        await importar(bytesExportados, caso.id + '_REABERTO.pdf');
        assert.equal(await avaliar(`document.querySelector('#pdf-assinaturas-alerta').hidden`), true);
        console.log(`CMS real ${caso.id}: ${caso.verificacoes.length} assinatura(s) original(is) verificadas antes/depois; cópia independente reaberta sem assinatura; ${caso.paginas.join(',')} rasterizada(s), demais preservadas.`);
    }
    console.log('Certificados de teste e PDFs com CMS reais (chave privada já removida): ' + reais.pasta);
    console.log("PASSOU: assinaturas estruturais, Resolver seletivo/inteiro, conferência/confirmar/cancelar, cópia independente, RGB 300/600 DPI, limites e cancelamento em 350 páginas.");
}
