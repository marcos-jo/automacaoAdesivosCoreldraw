import assert from "node:assert/strict";

export async function testarOrientacoesBrowser({ avaliar, clicar, pronto, esperarTela, imagem, baixar }) {
    const importar = async (retratos, paisagens, outroFormato = false) => {
        await avaliar(`(async()=>{
            const {PDFDocument}=await import('/vendor/pdf-lib/pdf-lib.js'),p=await PDFDocument.create();
            for(let i=0;i<${retratos+paisagens};i++)p.addPage(i<${retratos}?[210*72/25.4,297*72/25.4]:[297*72/25.4,210*72/25.4]);
            if(${outroFormato})p.addPage([297*72/25.4,420*72/25.4]);
            const d=new DataTransfer();d.items.add(new File([await p.save()],'orientacoes.pdf'));
            const e=document.querySelector('#pdf-arquivo');e.files=d.files;e.dispatchEvent(new Event('change',{bubbles:true}));
        })()`);
        await pronto();
    };
    const verificar = async (misturadas, tamanhos = false) => {
        if (misturadas) await esperarTela(`getComputedStyle(document.querySelector('#pdf-informacoes')).backgroundColor==='rgb(170, 118, 25)'`, 'Indicador amarelo após a transição CSS');
        const r = await avaliar(`(()=>{const q=s=>document.querySelector(s),i=q('#pdf-informacoes');return {
            alerta:i.classList.contains('pdf-alerta'),cor:getComputedStyle(i).backgroundColor,
            compacto:q('#pdf-orientacoes-alerta').textContent,oculto:q('#pdf-orientacoes-alerta').hidden,
            avisos:q('#pdf-alertas').textContent,detalhes:q('#pdf-orientacoes-detalhes').textContent
        }})()`);
        assert.equal(r.alerta, misturadas || tamanhos);
        assert.equal(r.oculto, !misturadas);
        assert.equal(r.avisos.includes('Orientações diferentes detectadas'), misturadas);
        assert.equal(r.avisos.includes('tamanhos diferentes'), tamanhos);
        if (misturadas) {
            assert.equal(r.compacto, 'Orientações diferentes detectadas');
            assert.equal(r.cor, 'rgb(170, 118, 25)', 'Indicador usa o amarelo existente');
        }
        await avaliar(`document.querySelector('#pdf-salvar-copia').click()`);
        await esperarTela(`document.querySelector('#app-exportacao-dialogo').open`, 'Diagnóstico Exportar');
        const avisos = await avaliar(`document.querySelector('#app-exportacao-avisos').textContent`);
        assert.equal(avisos.includes('Orientações diferentes detectadas'), misturadas);
        assert.equal(avisos.includes('tamanhos diferentes'), tamanhos);
        if (misturadas) assert.ok(avisos.includes(r.detalhes), 'Mesmas contagens e páginas em Informações/Exportar');
        await clicar('#app-exportacao-cancelar');
        return r;
    };
    for (const [retrato, paisagem] of [[10,0],[0,10],[9,1],[1,9]]) {
        await importar(retrato,paisagem);
        const r = await verificar(!!(retrato && paisagem));
        if (retrato) assert.match(r.detalhes,new RegExp(`Retrato: ${retrato} página\\(s\\)`));
        if (paisagem) assert.match(r.detalhes,new RegExp(`Paisagem: ${paisagem} página\\(s\\)`));
    }
    await importar(9,1,true); await verificar(true,true);
    await importar(10,0);
    const area = await avaliar(`(()=>{const r=document.querySelector('#pdf-visualizacao').getBoundingClientRect();return [r.top,r.height]})()`);
    await clicar('#pdf-girar-direita'); await pronto();
    let r = await verificar(true);
    assert.deepEqual(await avaliar(`(()=>{const r=document.querySelector('#pdf-visualizacao').getBoundingClientRect();return [r.top,r.height]})()`),area,'Aviso compacto não acrescenta uma linha à área técnica nem redimensiona a área da prévia');
    assert.match(r.detalhes,/Retrato: 9 página\(s\) — páginas 2–10.*Paisagem: 1 página\(s\) — páginas 1/);
    await clicar('#pdf-girar-esquerda'); await pronto(); await verificar(false);
    await clicar('#pdf-girar-esquerda'); await pronto(); await verificar(true); // 270°
    await avaliar(`document.querySelector('#pdf-mover-destino').value='10'`);
    await clicar('#pdf-mover-aplicar'); await pronto();
    r = await verificar(true);
    assert.match(r.detalhes,/Retrato: 9 página\(s\) — páginas 1–9.*Paisagem: 1 página\(s\) — páginas 10/);
    await clicar('#pdf-informacoes'); await imagem('22-orientacoes-misturadas');
    // Exportar confirma a ordem e o giro usados pelo mesmo diagnóstico, sem nova transformação.
    await clicar('#pdf-salvar-copia'); await pronto();
    const pdf = await baixar('orientacoes_editado.pdf',10);
    assert.deepEqual(pdf.getPages().map(p=>p.getRotation().angle),[0,0,0,0,0,0,0,0,0,270]);
    await avaliar(`document.querySelector('#pdf-selecao-paginas').value='10'`);
    await clicar('#pdf-excluir'); await pronto(); await verificar(false);
    await clicar('#pdf-ferramenta-unir'); await clicar('#pdf-uniao-abrir');
    await avaliar(`(async()=>{const {PDFDocument}=await import('/vendor/pdf-lib/pdf-lib.js'),p=await PDFDocument.create();p.addPage([297*72/25.4,210*72/25.4]);const d=new DataTransfer();d.items.add(new File([await p.save()],'paisagem.pdf'));const e=document.querySelector('#pdf-uniao-arquivos');e.files=d.files;e.dispatchEvent(new Event('change',{bubbles:true}));})()`);
    await clicar('#pdf-uniao-gerar'); await pronto(); await verificar(true);
    await clicar('#pdf-ferramenta-redimensionar');
    await avaliar(`(()=>{const e=document.querySelector('#pdf-redim-alvo');e.value='todas';e.dispatchEvent(new Event('change',{bubbles:true}))})()`);
    await clicar('#pdf-redim-prever');
    await esperarTela(`!document.querySelector('#pdf-redim-aplicar').disabled`, 'Prévia A4 retrato');
    await clicar('#pdf-redim-aplicar'); await pronto(); await verificar(false);
    await clicar('[data-close-tool="pdfs"]');
    await esperarTela(`document.querySelector('#pdf-confirmar-descarte')?.open`, 'Fechar PDF alterado');
    await clicar('#pdf-descartar-sair');
    await esperarTela(`!document.querySelector('#pdf-ferramenta')`, 'Fechado');
    await clicar('[data-tool="pdfs"]'); await esperarTela(`!!document.querySelector('#pdf-arquivo')`, 'PDFs após orientações');
    console.log('PASSOU: orientações uniformes/mistas nas duas maiorias; amarelo real, formatos independentes, 90°/270°, movimento, exclusão, união, resize e diagnóstico/exportação.');
}
