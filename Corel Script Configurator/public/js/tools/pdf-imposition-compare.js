import {validarCompatibilidadeImpressao} from './pdf-print-validation.js';
import {aguardarMontagem} from './pdf-imposition-preview.js';
// Adaptador isolado para o comparador existente. Nunca publica canvases no editor.
export async function compararCopiaMontagem(data){
    const pdfjs=await import('/vendor/pdfjs/pdf.js');pdfjs.GlobalWorkerOptions.workerSrc='/vendor/pdfjs/pdf.worker.js';
    const candidates=new Set();
    const fechar=c=>c.cleanup??=(async()=>{
        let timer;
        try{await Promise.race([Promise.resolve().then(()=>c.task.destroy()).catch(()=>{}),new Promise(resolve=>{timer=setTimeout(resolve,2000);})]);}
        finally{clearTimeout(timer);c.worker.destroy();candidates.delete(c);}
    })();
    try{
        return await validarCompatibilidadeImpressao(data,{
            async abrir(bytes,{signal}){
                if(signal?.aborted)throw new DOMException('Comparação cancelada','AbortError');
                const worker=new pdfjs.PDFWorker();let task;
                try{task=pdfjs.getDocument({data:new Uint8Array(bytes),worker});}
                catch(error){worker.destroy();throw error;}
                const candidate={task,worker};candidates.add(candidate);
                try{candidate.documento=await aguardarMontagem(task.promise,{signal,timeout:30000});return candidate;}
                catch(error){await fechar(candidate);throw error;}
            },fechar
        });
    }finally{await Promise.all([...candidates].map(fechar));}
}
