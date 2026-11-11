const fator = 72 / 25.4;
export function caixaDasMargens(viewport, margens) {
    const { esquerda, direita, superior, inferior } = margens;
    const valores = [esquerda, direita, superior, inferior];
    const k = viewport.scale * fator;
    if (!valores.every(v => Number.isFinite(v) && v >= 0) || (esquerda + direita) * k >= viewport.width || (superior + inferior) * k >= viewport.height) throw Error('Margens inválidas: o recorte precisa ter área positiva.');
    const a = viewport.convertToPdfPoint(esquerda * k, superior * k);
    const b = viewport.convertToPdfPoint(viewport.width - direita * k, viewport.height - inferior * k);
    return { x: Math.min(a[0], b[0]), y: Math.min(a[1], b[1]), width: Math.abs(b[0] - a[0]), height: Math.abs(b[1] - a[1]) };
}

// A imagem é somente a seleção visual; os bytes finais continuam vetoriais.
export async function abrirRecorte({ workspace, previa, numero, raiz, aplicar }) {
    const exibicao = previa.exibicao, fonte = previa.container.querySelector('canvas');
    if (!exibicao?.viewport || !fonte) throw Error('Aguarde a página terminar de carregar.');
    const viewport = exibicao.viewport;
    const versao = workspace.capturarVersao();
    const controller = new AbortController();
    await workspace.prepararRecorte(numero, caixaDasMargens(viewport, { esquerda:0, direita:0, superior:0, inferior:0 }), { signal:controller.signal });
    if (workspace.pdf !== versao.pdf || workspace.revisao !== versao.revisao) throw Error('O documento mudou. Abra o recorte novamente.');
    const dialog = document.createElement('dialog');
    dialog.className = 'pdf-recorte-dialog';
    dialog.innerHTML = `<h3>Recortar página ${numero}</h3><p>Arraste na página ou informe as margens em mm. O conteúdo fora das caixas fica oculto; não é apagado do PDF.</p><div class="pdf-recorte-campos">${['esquerda','direita','superior','inferior'].map(n => `<label>Margem ${n} (mm)<input name="${n}" type="number" min="0" step="0.1" value="0"></label>`).join('')}<label>Zoom da seleção (%)<input name="zoom" type="number" min="25" max="300" value="100"></label></div><p role="status" aria-live="polite"></p><div class="pdf-recorte-imagem"><canvas aria-label="Arraste para selecionar o recorte"></canvas></div><label class="pdf-recorte-aviso" hidden><input type="checkbox">Aceito remover as caixas de acabamento e sangria existentes. A sangria não será preservada.</label><div><button type="button" data-aplicar>Aplicar</button> <button type="button" data-cancelar>Cancelar</button></div>`;
    raiz.append(dialog);
    const canvas = dialog.querySelector('canvas'), ctx = canvas.getContext('2d');
    canvas.width = fonte.width; canvas.height = fonte.height;
    const imagem = document.createElement('canvas'); imagem.width = fonte.width; imagem.height = fonte.height; imagem.getContext('2d').drawImage(fonte, 0, 0);
    const inputs = [...dialog.querySelectorAll('.pdf-recorte-campos input')].filter(i => i.name !== 'zoom');
    const mensagem = dialog.querySelector('[role=status]'), botao = dialog.querySelector('[data-aplicar]'), aviso = dialog.querySelector('.pdf-recorte-aviso');
    let preparacao, geracao = 0, inicio, encerrado = false, aplicando = false, fila = Promise.resolve();
    const margens = () => Object.fromEntries(inputs.map(i => [i.name, Number(i.value === '' ? NaN : i.value)]));
    const ponto = evento => { const r = canvas.getBoundingClientRect(); return [Math.max(0,Math.min(viewport.width,(evento.clientX-r.left)/r.width*viewport.width)), Math.max(0,Math.min(viewport.height,(evento.clientY-r.top)/r.height*viewport.height))]; };
    async function atualizar() {
        const id = ++geracao; preparacao = null; botao.disabled = true; aviso.querySelector('input').checked = false;
        try {
            const m = margens(), caixa = caixaDasMargens(viewport,m), k = viewport.scale * fator;
            ctx.drawImage(imagem,0,0); ctx.fillStyle='rgba(0,0,0,.45)'; ctx.fillRect(0,0,canvas.width,canvas.height);
            const x=m.esquerda*k,y=m.superior*k,w=viewport.width-(m.esquerda+m.direita)*k,h=viewport.height-(m.superior+m.inferior)*k;
            ctx.drawImage(imagem,x,y,w,h,x,y,w,h); ctx.strokeStyle='#0078d4';ctx.lineWidth=2;ctx.strokeRect(x,y,w,h);
            mensagem.textContent = `Dimensões finais: ${(w/k).toFixed(2)} × ${(h/k).toFixed(2)} mm`;
            await new Promise(resolve => setTimeout(resolve, 120));
            if (encerrado || id !== geracao) return;
            const candidato = await (fila = fila.catch(() => {}).then(() => encerrado || id !== geracao ? null : workspace.prepararRecorte(numero, caixa, { signal:controller.signal })));
            if (encerrado || id !== geracao) return;
            if (workspace.pdf !== versao.pdf || workspace.revisao !== versao.revisao) throw Error('Prévia vencida. Cancele e abra novamente.');
            preparacao = candidato; aviso.hidden = !candidato.perdaSangria; botao.disabled = false;
        } catch (erro) { if (!encerrado && id === geracao) mensagem.textContent = erro.message; }
    }
    const finalizar = () => { encerrado=true; ++geracao;controller.abort(); dialog.close();dialog.remove();imagem.width=imagem.height=canvas.width=canvas.height=0; };
    return await new Promise(resolve => {
        dialog.querySelector('[data-cancelar]').onclick=()=>{if (!aplicando) { finalizar();resolve(false); }};
        dialog.oncancel=e=>{e.preventDefault();if (!aplicando) { finalizar();resolve(false); }};
        botao.onclick=async()=>{
            if (!preparacao) return;
            if (preparacao.perdaSangria && !aviso.querySelector('input').checked) { mensagem.textContent='Confirme a perda de sangria antes de aplicar.'; return; }
            aplicando=true;
            dialog.querySelectorAll("input,button").forEach(e=>e.disabled=true);
            try { await aplicar(preparacao,{aceitarPerdaSangria:aviso.querySelector('input').checked});finalizar();resolve(true); }
            catch (erro) { aplicando=false; mensagem.textContent=erro.message; dialog.querySelectorAll("input,[data-cancelar]").forEach(e=>e.disabled=false); }
        };
        inputs.forEach(i=>i.oninput=atualizar);
        dialog.querySelector('[name=zoom]').oninput=e=>{const z=Math.max(25,Math.min(300,Number(e.target.value)||100))/100;canvas.style.width=`${viewport.width*z}px`;canvas.style.height=`${viewport.height*z}px`;};
        canvas.onpointerdown=e=>{inicio=ponto(e);canvas.setPointerCapture(e.pointerId);};
        canvas.onpointermove=e=>{
            if(!inicio)return; const fim=ponto(e),k=viewport.scale*fator;
            const m={esquerda:Math.min(inicio[0],fim[0])/k,direita:(viewport.width-Math.max(inicio[0],fim[0]))/k,superior:Math.min(inicio[1],fim[1])/k,inferior:(viewport.height-Math.max(inicio[1],fim[1]))/k};
            inputs.forEach(i=>i.value=String(m[i.name]));atualizar();
        };
        canvas.onpointerup=canvas.onpointercancel=()=>{inicio=null;};
        dialog.showModal(); atualizar();
    });
}
