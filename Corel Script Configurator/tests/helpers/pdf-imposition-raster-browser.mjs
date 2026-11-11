// Executada no navegador: PDFs permanentes são sintéticos; original opcional só local.
export async function exercitarRasterMontagem(realBase64,iccBase64){
    const {PDFDocument,PDFName:N,PDFString,rgb,degrees}=await import('/vendor/pdf-lib/pdf-lib.js');
    const {PDFWorkspace}=await import('/js/tools/pdf-workspace.js');
    const {JanelaMontagem}=await import('/js/tools/pdf-imposition-ui.js'),{JanelaLivreto}=await import('/js/tools/pdf-booklet-ui.js');
    const {prepararImagemMontagem}=await import('/js/tools/pdf-imposition-raster.js');
    const pdfjs=await import('/vendor/pdfjs/pdf.js');pdfjs.GlobalWorkerOptions.workerSrc='/vendor/pdfjs/pdf.worker.js';
    const unpack=s=>Uint8Array.from(atob(s),c=>c.charCodeAt(0));
    const pack=b=>{let s='';for(let i=0;i<b.length;i+=32768)s+=String.fromCharCode(...b.subarray(i,i+32768));return btoa(s);};
    const check=(condition,message)=>{if(!condition)throw Error(message);};
    const sleep=ms=>new Promise(r=>setTimeout(r,ms)),until=async(fn,label)=>{for(let i=0;i<1500;i++){if(fn())return;await sleep(20);}throw Error(label+': '+window.rasterUi?.get('status').textContent);};
    const make=async(kind,count=1)=>{
        const p=await PDFDocument.create();for(let i=0;i<count;i++){const page=p.addPage([248,142]);page.drawRectangle({x:0,y:0,width:248,height:142,color:rgb(i%2?.1:.8,.4,i%2?.8:.1)});page.drawText('Synthetic page '+(i+1),{x:12,y:50,size:13});}
        if(kind.startsWith('ocg')||kind==='huge'||kind==='boxes'){const group=p.context.register(p.context.obj({Type:'OCG',Name:PDFString.of('Synthetic layer')})),off=kind==='ocg-off';p.catalog.set(N.of('OCProperties'),p.context.obj({OCGs:[group],D:{BaseState:off?'OFF':'ON',ON:off?[]:[group],OFF:off?[group]:[]}}));for(const page of p.getPages()){const contents=page.node.Contents().asArray();page.node.Resources().set(N.of('Properties'),p.context.obj({Layer:group}));page.node.set(N.of('Contents'),p.context.obj([p.context.register(p.context.stream('/OC /Layer BDC\n')),...contents,p.context.register(p.context.stream('\nEMC'))]));}}
        if(kind==='output'){const profile=p.context.register(p.context.flateStream(unpack(iccBase64),{N:3}));p.catalog.set(N.of('OutputIntents'),p.context.obj([{Type:'OutputIntent',S:'GTS_PDFX',OutputConditionIdentifier:PDFString.of('Synthetic RGB'),DestOutputProfile:profile}]));}
        if(kind==='vector')p.catalog.set(N.of('Outlines'),p.context.obj({}));
        if(kind==='xfa')p.catalog.set(N.of('AcroForm'),p.context.obj({XFA:PDFString.of('unsupported')}));
        if(kind==='boxes'){const page=p.getPage(0);page.setCropBox(10,10,228,122);page.setTrimBox(20,20,208,102);page.setBleedBox(15,15,218,112);page.setRotation(degrees(270));page.node.set(N.of('UserUnit'),p.context.obj(2));}
        if(kind==='huge')p.getPage(0).setSize(10000,10000);
        return p.save();
    };
    const root=document.createElement('div');document.body.append(root);let ui,accept=false;const prompts=[],previousConfirm=window.confirm;
    window.confirm=text=>{prompts.push(text);return accept;};
    const records=[];
    const open=async(Class,bytes,workspace=null)=>{
        if(ui)await ui.destroy();ui=window.rasterUi=new Class(root,{snapshot:async()=>workspace?workspace.obterBytes():bytes.slice(),imprimir:()=>{throw Error('Impressão proibida neste teste');}});await ui.open();
        await until(()=>!ui.busy,'gerar ou bloquear');
        if(!ui.get('compatibilidade').hidden)await until(()=>!ui.get('imagem-status').textContent.startsWith('Verificando'),'preflight');
    };
    try{
        const simple=await make('simple');await open(JanelaMontagem,simple);check(ui.get('imagem-opcoes').hidden,'PDF simples não deve oferecer raster automático');check(ui.canOutput(),'PDF simples deveria gerar');
        for(const kind of ['xfa','huge']){await open(JanelaMontagem,await make(kind));check(ui.get('preparar-imagem').disabled,kind+' não pode ser liberado');check(!ui.canOutput(),kind+' bloqueado');}
        // Uma fonte editada/pendente tem de permanecer exatamente igual depois.
        const synthetic=await make('ocg',3),ws=new PDFWorkspace();await ws.abrir({name:'synthetic.pdf',arrayBuffer:async()=>synthetic.slice().buffer});ws.girarPagina(1,90);const state=ws.capturarEstado(),pendingBytes=await ws.obterBytes();
        const cases=[{name:'ocg-pendente',bytes:pendingBytes,ws},{name:'ocg-off',bytes:await make('ocg-off')},{name:'output-intent',bytes:await make('output')},{name:'caixas-rotacao-userunit',bytes:await make('boxes')}];
        if(realBase64)cases.push({name:'real-local',bytes:unpack(realBase64)});
        for(const item of cases)for(const dpi of [300,600])for(const Class of [JanelaMontagem,JanelaLivreto]){
            await open(Class,item.bytes,item.ws);check(ui.get('preparar').disabled,'Vetor precisa continuar bloqueado');check(ui.imagem.disponivel,'Alternativa deveria estar disponível');check(!ui.get('previa-bloqueada').hidden,'Prévia deve dizer bloqueada');
            if(Class===JanelaMontagem){ui.get('colunas').value='1';ui.get('linhas').value='1';ui.get('modoMarcas').value='bordas';ui.controls();ui.schedule();await until(()=>!ui.busy,'configuração');await until(()=>ui.imagem.disponivel,'oferta nova');}
            ui.get('imagem-dpi').value=String(dpi);ui.imagem.atualizar();const before=ui.bytes.slice(),config=JSON.stringify(ui.config());
            accept=false;await ui.prepare(true);check(!ui.preparing&&ui.bytes.every((v,i)=>v===before[i]),'Recusar confirmação não pode mudar documento');check(!ui.canOutput(),'Recusa deve manter saída bloqueada');
            accept=true;await ui.prepare(true);await until(()=>!ui.busy&&!!ui.result,'montagem da imagem');
            const info=Class===JanelaMontagem?ui.prepared: {imagem:ui.preparoImagem};check(info.imagem?.dpi===dpi,'Resolução escolhida não aplicada');check(JSON.stringify(ui.config())===config,'Configuração alterada pelo preparo');check(ui.get('imagem-opcoes').hidden,'Não oferecer raster recursivo');check(ui.get('previa-bloqueada').hidden,'Bloqueio antigo da prévia');
            const source=await PDFDocument.load(before),copy=await PDFDocument.load(ui.bytes);check(source.getPageCount()===copy.getPageCount(),'Ordem/quantidade');
            const expected=source.getPages().map(p=>({width:p.getCropBox().width*(p.node.lookup(N.of('UserUnit'))?.asNumber()??1),height:p.getCropBox().height*(p.node.lookup(N.of('UserUnit'))?.asNumber()??1),rotation:p.getRotation().angle}));
            for(const [i,p]of copy.getPages().entries()){const e=expected[i],r=((e.rotation%360)+360)%360;check(Math.abs(p.getCropBox().width-(r%180?e.height:e.width))<1e-6&&Math.abs(p.getCropBox().height-(r%180?e.width:e.height))<1e-6,'Dimensões/orientação visual mudaram');}
            const proof={case:item.name,tool:Class===JanelaMontagem?'montagem':'livreto',dpi,expected,preparedSizes:copy.getPages().map(p=>p.getSize()),configPreserved:true,canOutput:ui.canOutput(),preview:ui.dialog.querySelector('canvas').toDataURL().split(',')[1],prepared:pack(ui.bytes),output:pack(ui.result.bytes)};
            if(Class===JanelaMontagem){proof.validation=ui.prepared.validacao;ui.get('modo').value='nup';ui.get('modo').dispatchEvent(new Event('input',{bubbles:true}));await until(()=>!ui.busy&&!!ui.result,'N-up da cópia');check(ui.result.plan.quantidade===copy.getPageCount(),'N-up deve usar todas as páginas');proof.nup=pack(ui.result.bytes);}
            else{ui.get('autorizarBrancas').checked=true;ui.buttons();}
            check(ui.canOutput(),'Resultado validado não liberou exportação');await ui.save();check(ui.get('status').textContent.includes('salvo'),'Exportação pelo botão/função existente falhou');proof.exportStatus=ui.get('status').textContent;
            records.push(proof);
        }
        check(ws.alterado,'Exportar montagem não pode quitar pendência da origem');check(ws.pdf===state.pdf&&ws.revisao===state.revisao,'Workspace substituído');check((await ws.obterBytes()).every((v,i)=>v===pendingBytes[i]),'Workspace mudou');
        // Falha real de renderização na preparação, seguida de retry na mesma janela.
        await open(JanelaMontagem,synthetic);const t=pdfjs.getDocument({data:synthetic.slice()}),doc=await t.promise,page=await doc.getPage(1),proto=Object.getPrototypeOf(page),render=proto.render;await t.destroy();
        const before=ui.bytes.slice();let started=false,cancelled=false;
        try{proto.render=function(){throw Error('Falha de rasterização simulada');};accept=true;await ui.prepare(true);check(ui.get('status').textContent.includes('Nenhuma cópia aplicada'),'Falha não explicada');check(ui.imagem.disponivel&&!ui.canOutput(),'Falha deve permitir repetir sem saída');check(ui.bytes.every((v,i)=>v===before[i]),'Falha publicou cópia');}
        finally{proto.render=render;}
        await ui.prepare(true);await until(()=>!ui.busy&&!!ui.result,'retry de raster');check(ui.canOutput(),'Retry ficou preso');
        for(const Class of [JanelaMontagem,JanelaLivreto]){
            await open(Class,await make('vector'));check(!ui.get('preparar').disabled,'Vetor deveria estar disponível inicialmente');
            try{proto.render=function(){throw Error('Comparação vetorial indisponível');};await ui.prepare();}
            finally{proto.render=render;}
            await until(()=>ui.imagem.disponivel,'alternativa após vetor falhar');check(!ui.canOutput(),'Falha vetorial não pode liberar saída');await ui.prepare(true);await until(()=>!ui.busy&&!!ui.result,'imagem após falha vetorial');
        }
        await open(JanelaMontagem,synthetic);
        try{proto.render=function(){started=true;let reject;return {promise:new Promise((_,r)=>{reject=r;}),cancel(){cancelled=true;reject(new DOMException('Cancelado','AbortError'));}};};ui.prepare(true);await until(()=>started,'início de raster pendente');await ui.close();check(cancelled,'Render antigo não cancelado');check(!ui.dialog.open,'Cancelar deve fechar');}
        finally{proto.render=render;}
        // Cancelamento na carga também precisa destruir recursos próprios.
        const a=new AbortController();a.abort();let aborted=false;try{await prepararImagemMontagem(synthetic,{confirmado:true,dpi:300,signal:a.signal});}catch(e){aborted=e.name==='AbortError';}check(aborted,'Abort inicial');
        return {records,workspacePendingPreserved:true,cancelledRender:cancelled,renderErrorRetry:true,vectorFailureFallbacks:2,prompts:prompts.length,disclaimerConfirmed:prompts.some(p=>p.includes('perfis')||p.includes('Perfis')),noPrinting:true};
    }finally{window.confirm=previousConfirm;if(ui)await ui.destroy();root.remove();delete window.rasterUi;}
}
