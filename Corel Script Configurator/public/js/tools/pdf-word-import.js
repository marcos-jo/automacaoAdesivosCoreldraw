import {PDFPreview} from './pdf-preview.js';

// A revisão usa uma fonte independente. O workspace só recebe bytes após o aceite.
export class JanelaImportarWord {
    constructor(root) {
        this.dialog=document.createElement('dialog');
        this.dialog.id='pdf-word-dialog';this.dialog.className='pdf-word-dialog app-exportacao-dialogo';
        this.dialog.innerHTML=`<h2>Importar Word</h2>
            <p id="pdf-word-status" role="status" aria-live="polite"></p>
            <p>O original permanece intacto. A conversão local pode alterar fontes, quebras, tabelas e paginação. Compare com o Word ou um PDF de referência.</p>
            <div id="pdf-word-report"></div>
            <nav aria-label="Conferir páginas do Word"><button id="pdf-word-prev" type="button" disabled>Anterior</button>
            <span id="pdf-word-page"></span><button id="pdf-word-next" type="button" disabled>Próxima</button>
            <select id="pdf-word-zoom" aria-label="Zoom da revisão"><option value="fit">Página inteira</option><option value="1">100%</option><option value="1.5">150%</option><option value="2">200%</option></select></nav>
            <div id="pdf-word-preview" class="pdf-word-preview"></div>
            <label><input id="pdf-word-check" type="checkbox" disabled> Conferi as páginas, fontes e quebras; aceito importar este PDF.</label>
            <div class="pdf-word-actions"><button id="pdf-word-cancel" type="button">Cancelar</button>
            <button id="pdf-word-import" type="button" disabled>Usar PDF convertido</button></div>`;
        root.append(this.dialog);
        this.get=id=>this.dialog.querySelector('#pdf-word-'+id);
        this.preview=new PDFPreview(this.get('preview'));
        this.preview.calcularEscala=base=>this.get('zoom').value==='fit'?Math.min(1.4,Math.max(120,this.get('preview').clientWidth-24)/base.width,Math.max(120,this.get('preview').clientHeight-16)/base.height):Number(this.get('zoom').value);
        this.get('cancel').addEventListener('click',()=>void this.cancelar());
        this.dialog.addEventListener('cancel',e=>{e.preventDefault();void this.cancelar();});
        this.get('check').addEventListener('change',()=>{this.get('import').disabled=!this.get('check').checked||this.rendering;});
        this.get('import').addEventListener('click',()=>{if(this.get('check').checked&&!this.rendering&&this.result)this.finish(this.result);});
        this.get('prev').addEventListener('click',()=>void this.mostrar(this.page-1).catch(()=>{}));
        this.get('next').addEventListener('click',()=>void this.mostrar(this.page+1).catch(()=>{}));
        this.get('zoom').addEventListener('change',()=>void this.mostrar(this.page).catch(()=>{}));
    }
    async abrir() {
        const bridge=globalThis.window?.desktop4PrintShops;
        if(!bridge?.importarWord)throw Error('A importação local de Word está disponível no aplicativo Desktop. No navegador, abra um PDF já convertido.');
        if(this.running||this.destroyed)return null;
        this.running=true;this.cancelled=false;this.result=null;this.page=1;
        this.get('report').replaceChildren();this.get('page').textContent='';this.get('zoom').value='fit';
        this.get('check').checked=false;this.get('check').disabled=true;
        this.get('import').disabled=true;this.get('prev').disabled=true;this.get('next').disabled=true;
        this.get('cancel').disabled=false;this.get('status').textContent='Selecione o Word. A conversão pode levar até dois minutos…';
        this.dialog.showModal();this.get('cancel').focus();
        const decision=new Promise(resolve=>{this.resolve=resolve;});
        try {
            this.converting=true;
            const result=await bridge.importarWord();this.converting=false;
            if(this.cancelled||this.destroyed||!result){this.finish(null);return await decision;}
            this.result=result;
            const report=result.report;
            let destination=this.get('report');
            const p=text=>{const element=document.createElement('p');element.textContent=text;destination.append(element);};
            p(`${result.name} — ${report.pages} página(s) no PDF convertido.`);
            p(report.sourcePages?`Word registra ${report.sourcePages} página(s) nos metadados. Esse número pode estar desatualizado.`:'Contagem original de páginas não verificada.');
            if(report.sourcePages&&report.sourcePages!==report.pages)p('Atenção: a contagem do PDF difere dos metadados do Word. Confira as quebras e a página final.');
            const details=document.createElement('details'),summary=document.createElement('summary');
            summary.textContent='Fontes e limites da conferência';details.append(summary);destination.append(details);destination=details;
            p('Fontes declaradas no Word (podem incluir fontes sem uso): '+(report.requestedFonts.join(', ')||'não verificadas neste formato')+'.');
            p('Fontes encontradas no PDF: '+(report.fonts.map(f=>f.name+(f.embedded?' (incorporada)':' (incorporação não confirmada)')).join(', ')||'nenhuma identificada')+'.');
            p('Substituição de fontes e fidelidade ao Word: não verificadas automaticamente. Cores, sangria e qualidade para RIP também não verificadas.');
            const candidate=await this.preview.preparar(result.bytes);
            if(this.cancelled||this.destroyed){await this.preview.descartar(candidate);this.finish(null);return await decision;}
            if(!candidate||candidate.documento.numPages!==report.pages)throw Error('A prévia não corresponde ao PDF convertido.');
            await this.preview.adotar(candidate);
            await this.mostrar(1);
            if(this.cancelled||this.destroyed){this.finish(null);return await decision;}
            this.get('check').disabled=false;
            this.get('status').textContent='Confira todas as páginas necessárias antes de aceitar. A importação substituirá o documento aberto após a confirmação de descarte existente.';
            return await decision;
        } catch(error) {
            this.finish(null);throw error;
        } finally {
            this.converting=false;this.running=false;this.result=null;
            this.dialog.close();
            await this.preview.cancelarRender();
            await this.preview.descartar(this.preview.atual);
            this.preview.atual=null;this.preview.documento=null;this.preview.limparCanvas();
        }
    }
    async mostrar(page) {
        if(this.rendering||!this.result||page<1||page>this.result.report.pages)return;
        this.rendering=true;this.get('prev').disabled=true;this.get('next').disabled=true;this.get('import').disabled=true;
        try {
            if(!await this.preview.mostrar(page))throw Error('Não foi possível mostrar esta página.');
            this.page=page;this.get('check').disabled=false;
            const d=this.result.report.dimensions[page-1];
            this.get('page').textContent=`Página ${page} de ${this.result.report.pages} — ${d.widthMm.toFixed(1)} × ${d.heightMm.toFixed(1)} mm`;
        } catch(error) {
            this.get('check').checked=false;this.get('check').disabled=true;
            this.get('status').textContent='Falha na conferência: '+error.message;
            throw error;
        } finally {
            this.rendering=false;
            this.get('prev').disabled=this.page<=1;this.get('next').disabled=this.page>=this.result.report.pages;
            this.get('import').disabled=!this.get('check').checked||this.get('check').disabled;
        }
    }
    finish(value) {const resolve=this.resolve;this.resolve=null;resolve?.(value);}
    async cancelar() {
        this.cancelled=true;this.get('cancel').disabled=true;
        if(this.converting) {
            this.get('status').textContent='Cancelando a conversão e limpando os arquivos temporários…';
            try{await globalThis.window?.desktop4PrintShops?.cancelarWord();}catch{}
        } else {this.preview.invalidar();this.finish(null);}
    }
    async destruir() {
        this.destroyed=true;await this.cancelar();this.finish(null);
        await this.preview.destruir();this.dialog.remove();
    }
}
