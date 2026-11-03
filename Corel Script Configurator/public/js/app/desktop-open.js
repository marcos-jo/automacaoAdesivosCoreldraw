// Mesmo shell/aba/workspace e operações de importação/união existentes.
export function instalarAberturaWindows(abas,interfacePrincipal){
    const bridge=globalThis.window?.desktop4PrintShops;
    if(!bridge?.arquivosWindows)return ()=>{};
    const dialog=document.createElement('dialog');dialog.id='desktop-open-dialog';dialog.className='app-exportacao-dialogo';
    interfacePrincipal.raiz.append(dialog);
    let pending=true,running=false,alive=true,resolveChoice,notifications=0;
    function choose(files,hasDocument,error){
        dialog.replaceChildren();const title=document.createElement('h2'),message=document.createElement('p'),list=document.createElement('ul');
        title.textContent=error?'Não foi possível abrir o PDF':'PDFs recebidos do Windows';
        message.textContent=error|| (hasDocument?'Já existe um documento na aba PDFs. Escolha como receber os arquivos.':'Há vários PDFs. Você pode adicioná-los em um documento ou abrir o primeiro; os demais permanecem na fila.');
        for(const file of files){const item=document.createElement('li');item.textContent=file.name;list.append(item);}
        dialog.append(title,message,list);
        const actions=error?[['retry','Tentar novamente'],['cancel','Dispensar abertura']]:[['open',hasDocument?'Substituir':'Abrir primeiro'],['add','Adicionar ao documento'],['cancel',hasDocument?'Manter o atual':'Dispensar abertura']];
        const result=new Promise(resolve=>{resolveChoice=resolve;});
        for(const [id,label]of actions){const b=document.createElement('button');b.type='button';b.id='desktop-open-'+id;b.textContent=label;b.onclick=()=>finish(id);dialog.append(b);}
        dialog.showModal();dialog.querySelector('#desktop-open-cancel').focus();return result;
    }
    function finish(value){const resolve=resolveChoice;resolveChoice=null;dialog.close();resolve?.(value);}
    dialog.addEventListener('cancel',e=>{e.preventDefault();finish('cancel');});
    async function process(){
        if(!alive||!pending||running||interfacePrincipal.raiz.inert)return;
        running=true;const revision=notifications;
        try{await abas.executar(async()=>{
            const files=await bridge.arquivosWindows();if(!files.length){pending=revision!==notifications;return;}
            if(!await abas.abrir('pdfs'))return;
            const pdf=abas.abas.get('pdfs').controle;if(pdf.estadoTrabalho().busy)return;
            let action=files.length===1&&!pdf.temDocumento()?'open':await choose(files,pdf.temDocumento());
            while(alive){
                const selected=action==='open'?[files[0]]:files,ids=selected.map(f=>f.id);
                if(action==='cancel'){await bridge.concluirArquivosWindows(ids);break;}
                try{
                    interfacePrincipal.raiz.inert=true;
                    const data=await bridge.lerArquivosWindows(ids),inputs=data.map(f=>new File([f.bytes],f.name,{type:'application/pdf'}));
                    if(action==='add')await pdf.adicionarArquivos(inputs);else await pdf.abrirArquivo(inputs[0]);
                    await bridge.concluirArquivosWindows(ids);break;
                }catch(error){
                    // Libera a transação antes de oferecer repetir/dispensar.
                    // A fila conserva os IDs após erros de leitura; o documento é transacional.
                    interfacePrincipal.raiz.inert=false;
                    await bridge.liberarArquivosWindows(ids);
                    const choice=await choose(selected,true,error.message);
                    if(choice==='cancel'){await bridge.concluirArquivosWindows(ids);break;}
                }finally{interfacePrincipal.raiz.inert=false;}
            }
        });}catch(error){interfacePrincipal.informar('Não foi possível receber o PDF: '+error.message);pending=false;}
        finally{running=false;}
    }
    const notify=()=>{notifications++;pending=true;void process();};
    const stop=bridge.observarArquivosWindows(notify),timer=setInterval(()=>void process(),300);
    const destroy=()=>{alive=false;finish('cancel');clearInterval(timer);stop();dialog.remove();window.removeEventListener('pagehide',destroy);};
    window.addEventListener('pagehide',destroy,{once:true});void process();return destroy;
}
