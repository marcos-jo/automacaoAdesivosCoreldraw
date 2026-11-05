// Uma tarefa isolada por prévia. Não toca o canvas visível nem o workspace.
export async function aguardarMontagem(promise,{signal,timeout=60000}={}) {
    let reject,timer;const stop=new Promise((_,r)=>{reject=r;});
    const abort=()=>reject(new DOMException('Montagem cancelada','AbortError'));
    signal?.addEventListener('abort',abort,{once:true});if(signal?.aborted)abort();
    timer=setTimeout(()=>reject(Error('A geração da montagem excedeu o tempo limite. Tente novamente com menos peças.')),timeout);
    try{return await Promise.race([promise,stop]);}finally{clearTimeout(timer);signal?.removeEventListener('abort',abort);}
}
export async function renderizarFolhaMontagem(result, number, {signal, load = () => import('/vendor/pdfjs/pdf.js'), createCanvas = () => document.createElement('canvas'), timeout = 30000, cleanupTimeout = 2000, largura=650,altura=650,zoom=1} = {}) {
    let task, worker, render, canvas, cleanup, failed=false;
    let rejectStop, stopped = false;
    const interrupt = new Promise((_, reject) => {rejectStop = reject;});
    // A promise pode rejeitar antes do primeiro await após criar o worker.
    interrupt.catch(() => {});
    const stop = error => {if(stopped)return;stopped=true;rejectStop(error);render?.cancel();};
    const abort = () => stop(new DOMException('Prévia cancelada', 'AbortError'));
    signal?.addEventListener('abort',abort,{once:true});
    const timer=setTimeout(()=>stop(Error('A prévia excedeu o tempo limite. Tente novamente ou reduza a montagem.')),timeout);
    const wait = promise => Promise.race([Promise.resolve(promise),interrupt]);
    const bounded = promise => {let timer;return Promise.race([Promise.resolve(promise),new Promise(resolve=>{timer=setTimeout(resolve,cleanupTimeout);})]).finally(()=>clearTimeout(timer));};
    const dispose = () => cleanup ??= (async()=>{
        try{render?.cancel();await bounded(render?.promise?.catch(()=>{}));if(task)await bounded(Promise.resolve().then(()=>task.destroy()).catch(()=>{}));}
        finally{worker?.destroy();}
    })();
    try{
        if(signal?.aborted)abort();
        const pdfjs=await wait(load());
        pdfjs.GlobalWorkerOptions.workerSrc='/vendor/pdfjs/pdf.worker.js';
        worker=new pdfjs.PDFWorker();
        task=pdfjs.getDocument({data:(result.bytesPrevia||result.bytes).slice(),worker});
        const doc=await wait(task.promise),page=await wait(doc.getPage(number));
        const unit=page.getViewport({scale:1});
        if(![largura,altura,zoom].every(n=>Number.isFinite(n)&&n>0)||zoom>100)throw Error('Enquadramento inválido.');
        const viewport=page.getViewport({scale:Math.min(largura/unit.width,altura/unit.height)*zoom});
        if(viewport.width>12000||viewport.height>12000||viewport.width*viewport.height>36000000)throw Error('Zoom acima do limite de memória da prévia. Reduza a ampliação.');
        canvas=createCanvas();canvas.width=Math.ceil(viewport.width);canvas.height=Math.ceil(viewport.height);
        const context=canvas.getContext('2d');if(!context)throw Error('Não foi possível criar a prévia.');
        render=page.render({canvas,canvasContext:context,viewport,background:'white',intent:'print'});
        await wait(render.promise);
        if(signal?.aborted)abort();if(stopped)await interrupt;
        return {canvas,viewport,pages:doc.numPages};
    }catch(error){failed=true;throw error;}
    finally{clearTimeout(timer);signal?.removeEventListener('abort',abort);await dispose();if(failed&&canvas)canvas.width=canvas.height=0;}
}
