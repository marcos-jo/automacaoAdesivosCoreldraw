// Único contato opcional do frontend com a ponte restrita do Desktop.
export function informarDocumento({ pending, busy }) {
    globalThis.window?.desktop4PrintShops?.estadoDocumento({ pending: !!pending, busy: !!busy });
}

export function instalarImportacaoDesktop(raiz,{iniciar=()=>true,finalizar=()=>{}}={}) {
    const desktop=globalThis.window?.desktop4PrintShops;
    if(!desktop?.abrirPDF&&!desktop?.importarArquivos)return ()=>{};
    let encerrada=false,escolhendo=false,progress;
    if(desktop.importarArquivos) {
        progress=document.createElement('dialog');progress.id='pdf-import-progress';progress.className='app-exportacao-dialogo';
        progress.innerHTML='<h2>Importar arquivos</h2><p role="status">Selecione PDF, imagens ou Word. Documentos Word serão convertidos localmente. Aguarde a preparação da seleção.</p><button type="button">Cancelar importação</button>';
        raiz.append(progress);
        const cancel=async()=>{progress.querySelector('button').disabled=true;try{await desktop.cancelarImportacao();}catch{}};
        progress.querySelector('button').addEventListener('click',()=>void cancel());
        progress.addEventListener('cancel',e=>{e.preventDefault();void cancel();});
    }
    const selecionar=async evento=>{
        const input=evento.target;
        if(input.tagName!=='INPUT'||input.type!=='file'||!input.accept?.includes('pdf'))return;
        evento.preventDefault();
        if(input.disabled||escolhendo)return;
        const unified=!!desktop.importarArquivos&&input.accept.includes('.doc');
        if(unified&&!iniciar())return;
        escolhendo=true;
        let files;
        try {
            if(unified){raiz.querySelector('#pdf-status').textContent='Importando arquivos…';progress.querySelector('button').disabled=false;progress.showModal();}
            const entries=unified?await desktop.importarArquivos(input.multiple):await desktop.abrirPDF(input.multiple,input.accept.includes('image/'));
            if(encerrada)return;
            if(!entries.length){if(unified)raiz.querySelector('#pdf-status').textContent='Importação cancelada.';return;}
            files=entries.map(entry=>{
                const type=/\.png$/i.test(entry.name)?'image/png':/\.jpe?g$/i.test(entry.name)?'image/jpeg':'application/pdf';
                const file=new File([entry.bytes],entry.name,{type});
                if(entry.wordReport)file.wordReport=entry.wordReport;
                return file;
            });
            if(unified)raiz.querySelector('#pdf-status').textContent='Seleção preparada. Confira os arquivos antes de combinar.';
        }catch(error){if(!encerrada)raiz.querySelector('#pdf-status').textContent='Não foi possível importar: '+error.message;}
        finally {
            escolhendo=false;
            if(unified){progress.close();finalizar();}
        }
        if(!encerrada&&files?.length) {
            // Os metadados de conversão acompanham os próprios Files no evento;
            // não dependem de propriedades preservadas pela cópia do DataTransfer.
            const list=new DataTransfer();for(const file of files)list.items.add(file);
            input.files=list.files;
            input.dispatchEvent(new CustomEvent('change',{bubbles:true,detail:{files}}));
        }
    };
    raiz.addEventListener('click',selecionar);
    return ()=>{encerrada=true;raiz.removeEventListener('click',selecionar);if(escolhendo)void desktop.cancelarImportacao?.();progress?.remove();};
}
