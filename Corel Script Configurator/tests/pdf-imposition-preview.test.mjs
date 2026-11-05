import test from 'node:test';
import assert from 'node:assert/strict';
import {readFile} from 'node:fs/promises';
import {renderizarFolhaMontagem,aguardarMontagem} from '../public/js/tools/pdf-imposition-preview.js';
const pending=()=>Promise.withResolvers(), sleep=ms=>new Promise(r=>setTimeout(r,ms));
const result={bytes:new Uint8Array([1]),plan:{folha:{largura:100,altura:100},util:{largura:80,altura:80},colunas:1,linhas:1,capacidade:1,quantidade:1,totalFolhas:1,folhas:[[]]},avisos:[]};
const canvas=()=>({width:1,height:1,getContext:()=>({strokeRect(){},setLineDash(){}})});
function fake({loadPending=false,renderPending=false,destroyPending=false,renderError=false}={}){
    const stats={cancel:0,destroy:0,worker:0};const gate=pending();gate.promise.catch(()=>{});
    const render={promise:renderPending?gate.promise:renderError?Promise.reject(Error('Falha sintética')):Promise.resolve(),cancel(){stats.cancel++;gate.reject(new DOMException('Cancelado','AbortError'));}};render.promise.catch(()=>{});
    const page={getViewport:()=>({width:100,height:100}),render:()=>render};
    const task={promise:loadPending?new Promise(()=>{}):Promise.resolve({numPages:1,getPage:async()=>page}),destroy(){stats.destroy++;return destroyPending?new Promise(()=>{}):Promise.resolve();}};
    return {stats,load:async()=>({GlobalWorkerOptions:{},PDFWorker:class{destroy(){stats.worker++;}},getDocument:()=>task})};
}
test('cancelar carregamento que não resolve libera a fila e encerra worker próprio',async()=>{
    const mock=fake({loadPending:true}),abort=new AbortController();const p=renderizarFolhaMontagem(result,1,{signal:abort.signal,load:mock.load,createCanvas:canvas,cleanupTimeout:5});await sleep(0);abort.abort();await assert.rejects(p,{name:'AbortError'});assert.equal(mock.stats.destroy,1);assert.equal(mock.stats.worker,1);
});
test('cancelamento chama RenderTask.cancel e destroy é executado uma vez',async()=>{
    const mock=fake({renderPending:true}),abort=new AbortController();const p=renderizarFolhaMontagem(result,1,{signal:abort.signal,load:mock.load,createCanvas:canvas});await sleep(0);abort.abort();await assert.rejects(p,{name:'AbortError'});assert.ok(mock.stats.cancel);assert.equal(mock.stats.destroy,1);
});
test('timeout e limpeza pendente não bloqueiam uma renderização posterior',async()=>{
    const mock=fake({renderPending:true,destroyPending:true});await assert.rejects(renderizarFolhaMontagem(result,1,{load:mock.load,createCanvas:canvas,timeout:5,cleanupTimeout:5}),/tempo limite/);assert.equal(mock.stats.worker,1);
    const next=fake();assert.ok((await renderizarFolhaMontagem(result,1,{load:next.load,createCanvas:canvas})).canvas);
});
test('falha de renderização é compreensível e nova tentativa pode concluir',async()=>{
    await assert.rejects(renderizarFolhaMontagem(result,1,{load:fake({renderError:true}).load,createCanvas:canvas}),/Falha sintética/);
    assert.equal((await renderizarFolhaMontagem(result,1,{load:fake().load,createCanvas:canvas})).pages,1);
});
test('geração pendente ou abortada não prende atualizações seguintes',async()=>{
    await assert.rejects(aguardarMontagem(new Promise(()=>{}),{timeout:5}),/tempo limite/);const abort=new AbortController();const p=aguardarMontagem(new Promise(()=>{}),{signal:abort.signal});abort.abort();await assert.rejects(p,{name:'AbortError'});assert.equal(await aguardarMontagem(Promise.resolve(42)),42);
});
const asModule=code=>'data:text/javascript;base64,'+Buffer.from(code).toString('base64');
const uiText=await readFile(new URL('../public/js/tools/pdf-imposition-ui.js',import.meta.url),'utf8');
const engine=asModule('export const gerarMontagem=(...args)=>globalThis.__impositionTestGenerate(...args);');
const mapped=uiText.replace("'../app/save-file.js'",JSON.stringify(new URL('../public/js/app/save-file.js',import.meta.url).href)).replace("'./pdf-imposition-preview.js'",JSON.stringify(new URL('../public/js/tools/pdf-imposition-preview.js',import.meta.url).href)).replace("'./pdf-imposition-bounds.js'",JSON.stringify(new URL('../public/js/tools/pdf-imposition-bounds.js',import.meta.url).href)).replace("'./pdf-imposition-view.js'",JSON.stringify(new URL('../public/js/tools/pdf-imposition-view.js',import.meta.url).href)).replace("'./pdf-imposition-engine.js'",JSON.stringify(engine));
const {JanelaMontagem}=await import(asModule(mapped.replace("'./pdf-imposition-settings.js'",JSON.stringify(new URL('../public/js/tools/pdf-imposition-settings.js',import.meta.url).href)).replace("'./pdf-imposition-summary.js'",JSON.stringify(new URL('../public/js/tools/pdf-imposition-summary.js',import.meta.url).href))));
function ui(){const instance=Object.create(JanelaMontagem.prototype),fields=new Map();Object.assign(instance,{version:0,sequence:Promise.resolve(),result,bytes:result.bytes,page:1,dialog:{open:true,querySelector:()=>canvas()},get:id=>{if(!fields.has(id))fields.set(id,{hidden:false});return fields.get(id);},config:()=>({}),render:async()=>{}});return instance;}
test('UI descarta resultado antigo, recupera falha e libera exportação após retry',async()=>{
    const instance=ui(),gate=pending();let calls=0;globalThis.__impositionTestGenerate=async()=>{if(++calls===1)return gate.promise;return result;};
    instance.schedule();await sleep(280);instance.schedule();await sleep(300);await instance.sequence;assert.equal(instance.result,result);assert.equal(instance.get('salvar').disabled,false);gate.resolve({...result,avisos:['antigo']});await sleep(0);assert.deepEqual(instance.result.avisos,[]);
    instance.render=async()=>{throw Error('Erro de prévia');};instance.schedule();await sleep(280);await instance.sequence;assert.equal(instance.get('tentar').hidden,false);assert.equal(instance.get('salvar').disabled,true);
    instance.render=async()=>{};instance.schedule();await sleep(280);await instance.sequence;assert.equal(instance.get('tentar').hidden,true);assert.equal(instance.get('salvar').disabled,false);assert.equal(instance.get('imprimir').disabled,false);delete globalThis.__impositionTestGenerate;
});
test('UI mudanças rápidas coalescem; cancelar e reabrir aguardam encerramento',async()=>{
    const instance=ui();let calls=0;globalThis.__impositionTestGenerate=async()=>{calls++;return result;};for(let i=0;i<30;i++)instance.schedule();await sleep(280);await instance.sequence;assert.equal(calls,1);
    const gate=pending();instance.sequence=gate.promise;instance.dialog.close=()=>{instance.dialog.open=false;};const closing=instance.close();assert.ok(instance.closing);gate.resolve();await closing;assert.equal(instance.result,null);assert.equal(instance.bytes,null);delete globalThis.__impositionTestGenerate;
});
test('alternar guias durante geração invalidada também regenera e libera ações',async()=>{
    const instance=ui();instance.result=null;let calls=0;globalThis.__impositionTestGenerate=async()=>{calls++;return result;};instance.schedule(false);await sleep(280);await instance.sequence;assert.equal(calls,1);assert.equal(instance.get('salvar').disabled,false);delete globalThis.__impositionTestGenerate;
});
test('excesso exige aceite; cancelar preserva resultado, nova geometria revoga e zoom mantém aceite',async()=>{
    const instance=ui(),r={...result,geometria:{requerConfirmacao:true}};instance.result=r;instance.geometryState=()=>{};
    const old=globalThis.confirm;try{
        globalThis.confirm=()=>false;instance.buttons(false);assert.equal(instance.get('salvar').disabled,true);assert.equal(instance.acceptClipping(),false);assert.equal(instance.result,r);assert.equal(instance.canOutput(),false);
        globalThis.confirm=()=>true;assert.equal(instance.acceptClipping(),true);assert.equal(instance.canOutput(),true);assert.equal(instance.get('imprimir').disabled,false);
        instance.schedule(false);await sleep(280);await instance.sequence;assert.equal(instance.approved,r);assert.equal(instance.canOutput(),true);
        globalThis.__impositionTestGenerate=async()=>({...r});instance.schedule();assert.equal(instance.approved,null);assert.equal(instance.canOutput(),false);await sleep(280);await instance.sequence;assert.notEqual(instance.result,r);assert.equal(instance.get('salvar').disabled,true);
    }finally{globalThis.confirm=old;delete globalThis.__impositionTestGenerate;}
});
