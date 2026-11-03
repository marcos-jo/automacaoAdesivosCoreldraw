// Nenhuma rede ou caminho de arquivo no renderer. Navegador não instala listeners.
export function instalarAtualizacaoAssistida(raiz){
    const desktop=globalThis.window?.desktop4PrintShops;
    if(!desktop?.observarAssistida)return ()=>{};
    const notice=document.createElement('aside');notice.id='desktop-assisted-notice';notice.className='desktop-update-notice';notice.hidden=true;notice.setAttribute('role','status');
    const text=document.createElement('span'),progress=document.createElement('progress'),action=document.createElement('button'),later=document.createElement('button');
    progress.max=100;progress.setAttribute('aria-label','Download assistido');later.textContent='Mais tarde';notice.append(text,progress,action,later);raiz.before(notice);
    let alive=true,revision=0,phase='idle';
    function render(state){
        revision++;phase=state.phase;
        notice.hidden=!['available','downloading','ready','error'].includes(phase);
        text.textContent=phase==='available'?`Rede interna: versão ${state.version} disponível`:phase==='ready'?`Instalador ${state.version} verificado — instalação assistida`:phase==='downloading'?`Download assistido: ${Math.round(state.percent||0)}%`:'Rede interna: atualização indisponível. Seu trabalho pode continuar.';
        progress.hidden=phase!=='downloading';progress.value=state.percent||0;
        action.disabled=phase==='downloading';action.textContent=phase==='ready'?'Preparar instalação':phase==='error'?'Verificar novamente':'Ver novidades';
    }
    action.onclick=()=>desktop.atualizacaoAssistida(phase==='ready'?'prepare':phase==='error'?'check':'details').catch(()=>{});
    later.onclick=()=>{notice.hidden=true;};
    const stops=[desktop.observarAssistida(render),desktop.observarBloqueio(locked=>{notice.inert=locked;})];
    const initial=revision;desktop.estadoAssistida().then(state=>{if(alive&&revision===initial)render(state);}).catch(()=>{});
    function destroy(){if(!alive)return;alive=false;for(const stop of stops)stop();notice.remove();window.removeEventListener('pagehide',destroy);}
    window.addEventListener('pagehide',destroy,{once:true});return destroy;
}
