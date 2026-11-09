// Área provisória: nenhum resultado parcial substitui o rascunho ou o workspace.
// Dependências recebidas da janela existente, sem outro motor de importação.
export class PreparacaoConferencia {
    constructor({list,count,host,PDFCombinacao,PDFPreview,wordReports,changed=()=>{}}){
        Object.assign(this,{list,count,host,PDFCombinacao,PDFPreview,wordReports,changed});
        this.batch=null;
    }
    begin(files,previous,cancel){
        if(this.batch)throw Error('Conclua ou cancele a preparação anterior.');
        const draft=new this.PDFCombinacao();
        Object.assign(draft,{sources:[...previous.sources],pages:[...previous.pages],serial:previous.serial,size:previous.size});
        const batch={draft,cancel,items:[],cancelled:false,failed:false,abort:new AbortController(),preview:new this.PDFPreview(this.host)};
        batch.preview.calcularEscala=base=>Math.min(100/base.width,115/base.height);
        this.batch=batch;this.list.hidden=false;this.list.replaceChildren();this.count.hidden=false;
        for(const file of files){
            const row=document.createElement('li');row.className='pdf-import-item';
            const thumb=document.createElement('div');thumb.className='pdf-combinar-miniatura pdf-import-loading';thumb.setAttribute('aria-label','Aguardando preparação');
            const label=document.createElement('span');label.textContent=`${file.name} · ${file.type}`;
            const state=document.createElement('span');state.className='pdf-import-state';state.textContent='Aguardando';
            row.append(thumb,label,state);this.list.append(row);batch.items.push({row,thumb,state,ready:false});
        }
        this.update();return batch;
    }
    current(batch){return !!batch&&this.batch===batch&&!batch.cancelled&&!batch.failed;}
    dispose(batch){return batch.disposal ||= batch.preview.destruir();}
    update(){const b=this.batch;if(!b)return;this.count.textContent=`${b.items.filter(i=>i.ready).length} de ${b.items.length} arquivo(s) pronto(s).${b.failed?' Falha na seleção; nenhum arquivo foi importado.':''}`;this.changed();}
    stage(index,state){
        const b=this.batch,item=b?.items[index];if(!item||!this.current(b)||item.ready)return;
        item.state.textContent=state==='word'?'Convertendo Word':state==='failed'?'Falha':state==='pdf'?'Lendo PDF':'Lendo imagem';
        this.update();
    }
    async accept(index,file,batch=this.batch){
        if(!this.current(batch))return false;
        const item=batch.items[index];if(!item||item.ready)throw Error('Sequência de preparação inválida.');
        item.state.textContent=/\.pdf$/i.test(file.name)?'Lendo PDF':'Preparando imagem';this.update();
        const previous=batch.draft.sources.length;
        await batch.draft.add([file]);
        if(!this.current(batch))return false;
        const source=batch.draft.sources[previous];if(file.wordReport)this.wordReports.set(source,file.wordReport);
        item.state.textContent='Gerando miniaturas';this.update();
        const preview=batch.preview,candidate=await preview.preparar(source.previewBytes,{signal:batch.abort.signal});
        if(!this.current(batch)){if(candidate)await preview.descartar(candidate);return false;}
        if(!candidate||!await preview.adotar(candidate)||!await preview.mostrar(1))throw Error('Não foi possível validar a miniatura.');
        if(!this.current(batch))return false;
        const canvas=this.host.querySelector('canvas');if(!canvas?.width)throw Error('Miniatura indisponível.');
        const img=document.createElement('img');img.src=canvas.toDataURL('image/png');img.alt='Primeira página validada';
        item.thumb.classList.remove('pdf-import-loading');item.thumb.replaceChildren(img);item.thumb.setAttribute('aria-label','Miniatura pronta');
        item.ready=true;item.state.textContent='Pronto';this.update();return true;
    }
    async complete(){
        const b=this.batch;if(!b||!this.current(b)||!b.items.length||b.items.some(i=>!i.ready))throw Error('A seleção ainda não está pronta.');
        const draft=b.draft;await this.dispose(b);if(!this.current(b))return null;
        this.batch=null;this.list.hidden=true;this.list.replaceChildren();this.count.hidden=true;this.changed();return draft;
    }
    fail(index){
        const b=this.batch;if(!b||b.cancelled||b.failed)return;b.failed=true;b.draft=null;b.abort.abort();b.preview.invalidar();
        for(const [i,item]of b.items.entries()){
            item.thumb.classList.remove('pdf-import-loading');
            if(i===index||!item.ready)item.state.textContent=i===index?'Falha':'Não importado';
        }
        void this.dispose(b).catch(()=>{});this.update();
    }
    cancel(){
        const b=this.batch;if(!b)return;b.cancelled=true;this.batch=null;b.draft=null;b.abort.abort();
        b.preview.invalidar();void this.dispose(b).catch(()=>{});
        this.list.hidden=true;this.list.replaceChildren();this.count.hidden=true;this.changed();
        b.cancel?.();
    }
}
