export function resumirTrabalho(abas) {
    let pending=false,busy=false;
    for(const aba of abas){
        const state=aba.controle.estadoTrabalho?.();
        if(state){pending ||= state.pending;busy ||= state.busy;}
    }
    return {pending:!!pending,busy:!!busy};
}

export function instalarAtualizacoes(raiz,abas) {
    const desktop=globalThis.window?.desktop4PrintShops;
    if(!desktop?.observarAtualizacoes)return ()=>{};
    const notice=document.createElement('aside');notice.id='desktop-update-notice';notice.className='desktop-update-notice';notice.hidden=true;notice.setAttribute('role','status');
    const text=document.createElement('span'),progress=document.createElement('progress'),details=document.createElement('button'),later=document.createElement('button');
    details.textContent='Ver novidades';later.textContent='Mais tarde';progress.max=100;progress.setAttribute('aria-label','Download da atualização');
    notice.append(text,progress,details,later);raiz.before(notice);
    let alive=true,revision=0;
    const render=state=>{
        revision++;
        notice.hidden=!['available','downloading','ready','error'].includes(state.phase);
        text.textContent=state.phase==='available'?`Nova versão ${state.version} disponível`:state.phase==='ready'?`Versão ${state.version} pronta para instalar`:state.phase==='downloading'?`Baixando atualização: ${Math.round(state.percent||0)}%`:'Não foi possível concluir a atualização.';
        progress.hidden=state.phase!=='downloading';progress.value=state.percent||0;
        details.disabled=state.phase==='downloading';details.textContent=state.phase==='ready'?'Revisar instalação':state.phase==='error'?'Ver detalhes / tentar novamente':'Ver novidades';
    };
    details.onclick=()=>desktop.verificarAtualizacoes();later.onclick=()=>{notice.hidden=true;};
    const stops=[desktop.observarAtualizacoes(render),desktop.consultarTrabalho(()=>resumirTrabalho(abas())),desktop.observarBloqueio(locked=>{raiz.inert=locked;notice.inert=locked;})];
    const initial=revision;
    desktop.estadoAtualizacao().then(state=>{if(alive&&revision===initial)render(state);}).catch(()=>{});
    const destroy=()=>{if(!alive)return;alive=false;for(const stop of stops)stop();raiz.inert=false;notice.remove();window.removeEventListener('pagehide',destroy);};
    window.addEventListener('pagehide',destroy,{once:true});
    return destroy;
}
