const chords={o:'open',s:'save',p:'print',e:'export'};
// Um único listener no renderer. O Electron mantém o menu nativo desativado.
export function instalarAtalhos({execute,available=()=>true},target=window){
    const handler=event=>{
        if(!(event.ctrlKey||event.metaKey)||event.altKey||event.shiftKey)return;
        const action=chords[event.key.toLowerCase()];if(!action)return;
        event.preventDefault();
        if(event.repeat||event.isComposing||event.target?.closest?.('input,textarea,select,[contenteditable="true"]')||document.querySelector('dialog[open]')||!available(action))return;
        Promise.resolve(execute(action)).catch(error=>console.error(error));
    };
    target.addEventListener('keydown',handler);
    return ()=>target.removeEventListener('keydown',handler);
}
