// Geometria pura em milímetros. Sem PDF, interface, rasterização ou impressão.
import {planejarNupCompacto} from './pdf-imposition-compact.js';
import {marcasIndividuais,marcasDoGrupo} from './pdf-imposition-crop-marks.js';
export const pontosPorMm = 72 / 25.4;
export const mmParaPt = value => value * pontosPorMm;
const epsilon = 1e-7;
function numero(value, name, min = 0, max = 5000) {
    const n = typeof value === 'string' ? Number(value.trim().replace(',', '.')) : value;
    if (!Number.isFinite(n) || n < min || n > max) throw Error(`${name}: valor fora do intervalo ${min}–${max}.`);
    return n;
}
function inteiro(value, name, min, max) {
    const n = numero(value, name, min, max);
    if (!Number.isInteger(n)) throw Error(`${name}: informe um número inteiro.`);
    return n;
}
export function planejarMontagem(config, fontes, {permitirExcesso=false}={}) {
    if (!['repeat', 'nup'].includes(config.modo)) throw Error('Escolha Step & Repeat ou N-up.');
    if (!fontes.length || fontes.length > 10000) throw Error('Selecione de 1 a 10.000 páginas/posições.');
    if (!['horizontal', 'vertical'].includes(config.ordem ?? 'horizontal')) throw Error('Ordem de preenchimento inválida.');
    const folha = { largura: numero(config.largura, 'Largura da folha', 10), altura: numero(config.altura, 'Altura da folha', 10) };
    const margem = Object.fromEntries(['esquerda', 'direita', 'superior', 'inferior'].map(key => [key, numero(config.margens?.[key] ?? 10, `Margem ${key}`)]));
    const util = { x: margem.esquerda, y: margem.inferior, largura: folha.largura - margem.esquerda - margem.direita, altura: folha.altura - margem.superior - margem.inferior };
    if (util.largura <= 0 || util.altura <= 0) throw Error('As margens não deixam área útil.');
    const gapX = numero(config.espacoX ?? 0, 'Espaço horizontal'), gapY = numero(config.espacoY ?? 0, 'Espaço vertical');
    const sangria = numero(config.sangria ?? 0, 'Sangria', 0, 50);
    const modoMarcas = config.modoMarcas ?? (config.marcas === true ? 'individuais' : 'sem');
    if (!['individuais', 'grupo', 'bordas', 'sem'].includes(modoMarcas)) throw Error('Modo de marcas inválido.');
    const marcas = { modo: modoMarcas, ativas: modoMarcas !== 'sem', comprimento: numero(config.comprimento ?? 3, 'Comprimento das marcas', .1, 50), afastamento: numero(config.afastamento ?? 3, 'Afastamento das marcas', 0, 50), espessura: numero(config.espessura ?? .1, 'Espessura das marcas', .01, 2) };
    if (marcas.ativas && marcas.afastamento < sangria) throw Error('Afaste as marcas para fora da sangria.');
    const extensao = marcas.ativas ? Math.max(sangria, marcas.afastamento + marcas.comprimento + marcas.espessura / 2) : sangria;
    const marcasExternas = modoMarcas === 'grupo' || modoMarcas === 'bordas';
    const reserva = marcasExternas ? sangria : extensao;
    const reservaExterna = marcasExternas ? extensao - reserva : 0;
    const rotacao = inteiro(config.rotacao ?? 0, 'Rotação', 0, 270);
    if (rotacao % 90) throw Error('Use rotação de 0°, 90°, 180° ou 270°.');
    const paginas = fontes.map((fonte, index) => {
        if (fonte === null) return null; // Posição vazia explícita de N-up.
        const w = numero(fonte.largura, `Página ${index + 1}: largura`, .01), h = numero(fonte.altura, `Página ${index + 1}: altura`, .01);
        if (sangria > numero(fonte.sangriaDisponivel ?? 0, 'Sangria disponível', 0, 50) + epsilon) throw Error(`Página ${index + 1}: sangria solicitada não existe na origem.`);
        return { origem: fonte.id ?? index + 1, largura: rotacao % 180 ? h : w, altura: rotacao % 180 ? w : h, sangriaDisponivel: fonte.sangriaDisponivel ?? 0, rotacao };
    });
    if (!paginas.some(Boolean)) throw Error('A montagem precisa de uma página de origem.');
    if (config.modo === 'repeat' && (paginas.length !== 1 || !paginas[0])) throw Error('Step & Repeat exige exatamente uma arte.');
    const escalaModo = config.escalaModo ?? (config.modo==='nup'?'real':'ajustar');
    if (!(config.modo==='nup'?['ajustar','real','personalizada','reduzir','preencher']:['ajustar','real','personalizada']).includes(escalaModo)) throw Error('Modo de escala inválido.');
    const encaixar=['ajustar','reduzir','preencher'].includes(escalaModo);
    const escalaNaCelula=(w,h,f)=>{const base=encaixar&&w>0&&h>0?Math.min(w/f.largura,h/f.altura):escala;return escalaModo==='reduzir'?Math.min(1,base*fator):base*fator;};
    const escala = escalaModo === 'personalizada' ? numero(config.escala, 'Escala (%)', 1, 1000) / 100 : 1;
    const fator = numero(config.fatorMontagem ?? 1, 'Fator de encaixe', .000001, 1000);
    if(config.modo==='nup'){
        if(!['celulas','compacto'].includes(config.nupDistribuicao??'celulas'))throw Error('Distribuição N-up inválida.');
        if(config.nupDistribuicao==='compacto')return planejarNupCompacto({config,paginas,folha,margem,util,gapX,gapY,sangria,marcas,escalaModo,escala,fator,colunas:inteiro(config.colunas,'Colunas',1,1000),linhas:inteiro(config.linhas,'Linhas',1,1000),permitirExcesso});
    }
    let colunas, linhas, celulaW, celulaH;
    let orientadas, escalas;
    if (config.modo === 'repeat') {
        const fonte = paginas[0];
        celulaW = config.pecaLargura == null ? fonte.largura * escala : numero(config.pecaLargura, 'Largura final', .01);
        celulaH = config.pecaAltura == null ? celulaW * fonte.altura / fonte.largura : numero(config.pecaAltura, 'Altura final', .01);
        if (Math.abs(celulaW / celulaH - fonte.largura / fonte.altura) > 1e-6) throw Error('O tamanho final deformaria a arte. Preserve a proporção.');
        celulaW *= fator; celulaH *= fator;
        colunas = config.automatico ? Math.floor((util.largura - 2 * reservaExterna + gapX + epsilon) / (celulaW + 2 * reserva + gapX)) : inteiro(config.colunas, 'Colunas', 1, 1000);
        linhas = config.automatico ? Math.floor((util.altura - 2 * reservaExterna + gapY + epsilon) / (celulaH + 2 * reserva + gapY)) : inteiro(config.linhas, 'Linhas', 1, 1000);
        if(permitirExcesso&&config.automatico){colunas=Math.max(1,colunas);linhas=Math.max(1,linhas);}
    } else {
        colunas = inteiro(config.colunas, 'Colunas', 1, 1000); linhas = inteiro(config.linhas, 'Linhas', 1, 1000);
        celulaW = (util.largura - 2 * reservaExterna - (colunas - 1) * gapX) / colunas - 2 * reserva;
        celulaH = (util.altura - 2 * reservaExterna - (linhas - 1) * gapY) / linhas - 2 * reserva;
        if(permitirExcesso){
            const refW=Math.max(.01,celulaW),refH=Math.max(.01,celulaH);
            orientadas=paginas.map(f=>!f?null:config.rotacaoAutomatica&&Math.min(refW/f.altura,refH/f.largura)>Math.min(refW/f.largura,refH/f.altura)?{...f,largura:f.altura,altura:f.largura,rotacao:(f.rotacao+90)%360}:f);
            escalas=orientadas.map(f=>!f?null:escalaNaCelula(celulaW,celulaH,f));
            if(!encaixar||celulaW<=0||celulaH<=0||fator!==1){
                const dimensions=orientadas.map((f,i)=>f?{w:f.largura*escalas[i],h:f.altura*escalas[i]}:null).filter(Boolean);
                // Células crescem para acomodar a arte, em vez de sobrepor páginas
                // ou alterar escala silenciosamente quando a folha é pequena.
                celulaW=Math.max(celulaW,...dimensions.map(d=>d.w));celulaH=Math.max(celulaH,...dimensions.map(d=>d.h));
            }
        }
    }
    if (colunas < 1 || linhas < 1 || celulaW <= 0 || celulaH <= 0) throw Error('A arte/grade não cabe na folha com margens, sangria e marcas.');
    const capacidade = colunas * linhas;
    if (capacidade > 10000) throw Error('A grade ultrapassa 10.000 posições.');
    const ocupadoW = colunas * (celulaW + 2 * reserva) + (colunas - 1) * gapX + 2 * reservaExterna;
    const ocupadoH = linhas * (celulaH + 2 * reserva) + (linhas - 1) * gapY + 2 * reservaExterna;
    if (!permitirExcesso&&(ocupadoW > util.largura + epsilon || ocupadoH > util.altura + epsilon)) throw Error('A grade excede a área útil. Ajuste margens, grade ou papel.');
    const total = config.modo === 'repeat' ? inteiro(config.quantidade ?? capacidade, 'Quantidade', 1, 10000) : paginas.length;
    const totalFolhas = Math.ceil(total / capacidade);
    if (totalFolhas > 1000) throw Error('A montagem ultrapassa 1.000 folhas.');
    const x0 = util.x + (config.centralizar === false ? 0 : (util.largura - ocupadoW) / 2) + reservaExterna;
    const y0 = util.y + (config.centralizar === false ? util.altura - ocupadoH : (util.altura - ocupadoH) / 2) + reservaExterna;
    const folhas = Array.from({ length: totalFolhas }, () => []);
    for (let index = 0; index < total; index++) {
        const pos = index % capacidade; let fonte = config.modo === 'repeat' ? paginas[0] : (orientadas||paginas)[index];
        if (!fonte) continue;
        if (!orientadas&&config.modo === 'nup' && config.rotacaoAutomatica && Math.min(celulaW / fonte.altura, celulaH / fonte.largura) > Math.min(celulaW / fonte.largura, celulaH / fonte.altura)) fonte = {...fonte,largura:fonte.altura,altura:fonte.largura,rotacao:(fonte.rotacao+90)%360};
        const col = config.ordem === 'vertical' ? Math.floor(pos / linhas) : pos % colunas;
        const row = config.ordem === 'vertical' ? pos % linhas : Math.floor(pos / colunas);
        const s = config.modo === 'repeat' ? celulaW / fonte.largura : escalas ? escalas[index] : escalaNaCelula(celulaW,celulaH,fonte);
        const largura = fonte.largura * s, altura = fonte.altura * s;
        if (sangria > fonte.sangriaDisponivel * s + epsilon) throw Error(`Página ${fonte.origem}: a escala reduz a sangria disponível abaixo da solicitada.`);
        if (!permitirExcesso&&(largura > celulaW + epsilon || altura > celulaH + epsilon)) throw Error(`Página ${fonte.origem}: escala não cabe na célula. Nenhum corte automático será aplicado.`);
        const x = x0 + col * (celulaW + 2 * reserva + gapX) + reserva + (celulaW - largura) / 2;
        const y = y0 + (linhas - 1 - row) * (celulaH + 2 * reserva + gapY) + reserva + (celulaH - altura) / 2;
        const cortes = modoMarcas === 'individuais' ? marcasIndividuais(x,y,largura,altura,marcas) : [];
        folhas[Math.floor(index / capacidade)].push({ ...fonte, indice: index, x, y, largura, altura, escala: s, cortes });
    }
    const cortesGrupo = folhas.map(pecas => marcasDoGrupo(pecas,marcas));
    return { folha, util, margem, sangria, marcas, colunas, linhas, capacidade, quantidade: total, totalFolhas, folhas, cortesGrupo };
}
