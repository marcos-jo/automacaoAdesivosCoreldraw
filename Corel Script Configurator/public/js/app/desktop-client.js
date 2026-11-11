// Único contato opcional do frontend com a ponte restrita do Desktop.
export function informarDocumento({ pending, busy }) {
    globalThis.window?.desktop4PrintShops?.estadoDocumento({ pending: !!pending, busy: !!busy });
}

export function instalarImportacaoDesktop(raiz,{iniciar=()=>true,finalizar=()=>{},conferencia}={}) {
    const desktop=globalThis.window?.desktop4PrintShops;
    if(!desktop?.abrirPDF&&!desktop?.importarArquivos)return ()=>{};
    let encerrada=false,escolhendo=false,progress,active;
    const asFile=entry=>{
        const type=/\.png$/i.test(entry.name)?'image/png':/\.jpe?g$/i.test(entry.name)?'image/jpeg':'application/pdf';
        const file=new File([entry.bytes],entry.name,{type});if(entry.wordReport)file.wordReport=entry.wordReport;return file;
    };
    const cancel=async()=>{if(active){active.cancelled=true;active.wake?.();}if(progress)progress.querySelector('button').disabled=true;try{await desktop.cancelarImportacao();}catch{}};
    const delivered=async(job,count)=>{
        if(!Number.isInteger(count)||count<1||count>128)throw Error('Contagem de importação inválida.');
        let timer;
        try{await new Promise((resolve,reject)=>{
            job.wake=()=>{
                if(job.cancelled||encerrada)return resolve();
                if(job.error)return reject(job.error);
                if(job.selected&&job.received.size===count)return resolve();
            };
            timer=setTimeout(()=>reject(Error('A entrega da seleção não foi concluída. Cancele e selecione novamente.')),15000);
            job.wake();
        });}finally{clearTimeout(timer);job.wake=null;}
    };
    const unsubscribe=desktop.observarImportacao?.(event=>{
        const job=active;if(!job||job.id!==event.jobId||job.cancelled||encerrada)return;
        try{
            if(event.kind==='selection'){progress.close();conferencia.iniciar(event.files,()=>void cancel());job.selected=true;job.total=event.files.length;job.wake?.();}
            else if(event.kind==='stage') {job.index=event.index;conferencia.etapa(event.index,event.state);}
            else if(event.kind==='entry'){
                if(!Number.isInteger(event.index)||event.index<0||event.index>=job.total||job.received.has(event.index))throw Error('Sequência de importação inválida.');
                job.received.add(event.index);
                job.queue=job.queue.then(async()=>{if(!job.cancelled&&!job.error&&!encerrada)await conferencia.arquivo(event.index,asFile(event.entry));}).catch(error=>{
                    if(job.cancelled||encerrada||job.error)return;job.error=error;job.index=event.index;conferencia.falhar(event.index);job.wake?.();void desktop.cancelarImportacao().catch(()=>{});
                });
                job.wake?.();
            }
        }catch(error){job.error=error;if(job.selected)conferencia.falhar(job.index);job.wake?.();void desktop.cancelarImportacao().catch(()=>{});}
    });
    if(desktop.importarArquivos) {
        progress=document.createElement('dialog');progress.id='pdf-import-progress';progress.className='app-exportacao-dialogo';
        progress.innerHTML='<h2>Importar arquivos</h2><p role="status">Selecione PDF, imagens ou Word. Documentos Word serão convertidos localmente. Aguarde a preparação da seleção.</p><button type="button">Cancelar importação</button>';
        raiz.append(progress);
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
        let files;const job={id:crypto.randomUUID(),queue:Promise.resolve(),received:new Set(),cancelled:false,index:0};active=job;
        try {
            if(unified){raiz.querySelector('#pdf-status').textContent='Importando arquivos…';progress.querySelector('button').disabled=false;progress.showModal();}
            const progressive=unified&&!!desktop.observarImportacao&&!!conferencia;
            const entries=unified?await desktop.importarArquivos(input.multiple,progressive?job.id:undefined):await desktop.abrirPDF(input.multiple,input.accept.includes('image/'));
            if(encerrada)return;
            if(progressive){
                if(job.cancelled){conferencia.cancelar();raiz.querySelector('#pdf-status').textContent='Importação cancelada.';return;}
                if(!entries.count){conferencia.cancelar();raiz.querySelector('#pdf-status').textContent='Importação cancelada.';return;}
                await delivered(job,entries.count);await job.queue;
                if(job.cancelled){conferencia.cancelar();raiz.querySelector('#pdf-status').textContent='Importação cancelada.';return;}
                if(job.error)throw job.error;
                if(job.total!==entries.count)throw Error('A seleção recebida está incompleta.');
                const preparada = await conferencia.concluir();
                raiz.querySelector('#pdf-status').textContent = preparada === false
                    ? 'Confira os arquivos com falha na janela de importação.' : 'Seleção preparada. Confira os arquivos antes de combinar.';
                return;
            }
            if(!entries.length){if(unified)raiz.querySelector('#pdf-status').textContent='Importação cancelada.';return;}
            files=entries.map(asFile);
            if(unified)raiz.querySelector('#pdf-status').textContent='Seleção preparada. Confira os arquivos antes de combinar.';
        }catch(error){
            job.error ||= error;
            if(!encerrada&&!job.cancelled&&job.selected)conferencia.falhar(job.index);
            await job.queue;
            if(!encerrada){
                if(job.cancelled){conferencia?.cancelar();raiz.querySelector('#pdf-status').textContent='Importação cancelada.';}
                else{if(job.selected)conferencia.falhar(job.index);raiz.querySelector('#pdf-status').textContent='Não foi possível importar: '+(job.selected?'falha na preparação da seleção. Nenhum arquivo foi importado.':error.message);}
            }
        }
        finally {
            escolhendo=false;
            if(active===job)active=null;
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
    return ()=>{encerrada=true;if(active){active.cancelled=true;active.wake?.();}unsubscribe?.();raiz.removeEventListener('click',selecionar);if(escolhendo)void desktop.cancelarImportacao?.();progress?.remove();};
}
