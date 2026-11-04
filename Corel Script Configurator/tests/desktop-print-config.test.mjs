import test from 'node:test';
import assert from 'node:assert/strict';
import {mkdtemp,writeFile,readFile} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import path from 'node:path';
import {createHash} from 'node:crypto';
import {EventEmitter} from 'node:events';
import {PassThrough} from 'node:stream';
import {PrintConfigClient,validatePrintRequest} from '../desktop/print-config.mjs';
import {calcularPreviaImpressao,PrintConfigPresentation} from '../public/js/app/desktop-print-layout.js';
import {JanelaConfiguracaoImpressao} from '../public/js/app/desktop-print-ui.js';

test('canal de configuração aceita apenas comandos e campos necessários',()=>{
    assert.deepEqual(validatePrintRequest({action:'select',printer:'KONICA MINOLTA C6100PS'}),{action:'select',printer:'KONICA MINOLTA C6100PS'});
    for(const request of [null,{},[],{action:'print'},{action:'list',path:'C:\\x.exe'},{action:'select',printer:'bad\0name'},{action:'properties',session:'a'.repeat(32),revision:0},{action:'properties',session:'a'.repeat(32),revision:1,owner:'1'}])assert.throws(()=>validatePrintRequest(request));
});
test('A4 em A3 a 100% preserva tamanho e calcula margens reais',()=>{
    const layout=calcularPreviaImpressao({width:297,height:420},100);
    assert.deepEqual(layout.content,{left:43.5,top:61.5,width:210,height:297});assert.equal(layout.outsidePaper,false);
});
test('A4 em A3 a 120% mede 252 por 356,4 mm e fica centralizado',()=>{
    const layout=calcularPreviaImpressao({width:297,height:420},120);
    assert.equal(layout.content.width,252);assert.equal(layout.content.height,356.4);assert.equal(layout.margins.left,22.5);assert.ok(Math.abs(layout.margins.top-31.8)<1e-9);assert.equal(layout.outsidePaper,false);
});
test('prévia identifica corte e limite imprimível sem presumir área válida',()=>{
    const paper={width:210,height:297},area={...paper,left:5,top:5,printableWidth:200,printableHeight:287};
    assert.equal(calcularPreviaImpressao(paper,120,area).outsidePaper,true);
    assert.equal(calcularPreviaImpressao(paper,100,area).outsidePrintable,true);
    const mismatch=calcularPreviaImpressao({width:297,height:420},100,area);assert.equal(mismatch.geometryMismatch,true);assert.equal(mismatch.printable,null);
    assert.equal(calcularPreviaImpressao(paper,100,{...area,left:-1}).printable,null);
    for(const p of [{width:NaN,height:300},{width:0,height:297},{width:Infinity,height:297}])assert.throws(()=>calcularPreviaImpressao(p));
    assert.throws(()=>calcularPreviaImpressao(paper,144));
});
test('revisão e geração da interface rejeitam resultados antigos',()=>{
    const model=new PrintConfigPresentation(),old=model.begin(),current=model.begin();
    assert.equal(model.accept(old,{session:'old',revision:1}),false);
    assert.equal(model.accept(current,{session:'new',revision:2}),true);
    assert.equal(model.accept(current,{session:'new',revision:1}),false);assert.equal(model.state.revision,2);
    model.clear();assert.equal(model.accept(current,{session:'new',revision:3}),false);assert.equal(model.state,null);
});
test('reabrir aguarda a liberação nativa da sessão anterior',async()=>{
    const ui=Object.create(JanelaConfiguracaoImpressao.prototype);let release,listed=0;
    ui.busy=false;ui.model=new PrintConfigPresentation();ui.dialog={open:true,close(){this.open=false;},showModal(){this.open=true;}};
    ui.render=ui.controls=()=>{};ui.list=async()=>listed++;
    ui.bridge={configuracaoImpressao:()=>new Promise(resolve=>{release=resolve;})};
    const closing=ui.close(),opening=ui.open();await Promise.resolve();assert.equal(listed,0);release(true);await Promise.all([closing,opening]);assert.equal(listed,1);assert.equal(ui.busy,false);
});

async function fixture(t,{timeout=1000,tamper=false}={}){
    const folder=await mkdtemp(path.join(tmpdir(),'4printshops-print-config-unit-')),helper=path.join(folder,'4PrintShops.PrintConfig.exe'),manifest=path.join(folder,'manifest.json'),bytes=Buffer.from('test fixture, not executable');
    await writeFile(helper,bytes);await writeFile(manifest,JSON.stringify({schema:1,file:'4PrintShops.PrintConfig.exe',sha256:createHash('sha256').update(tamper?Buffer.from('modified'):bytes).digest('hex')}));
    const children=[];
    const client=new PrintConfigClient({helper,manifest,platform:'win32',timeout,propertiesTimeout:timeout,launch:()=>{
        const child=new EventEmitter();child.stdin=new PassThrough();child.stdout=new PassThrough();child.kill=()=>{};children.push(child);return child;
    }});
    t.after(()=>client.destroy());
    const command=async promise=>{for(let n=0;n<100;n++){const child=children.at(-1),chunk=child?.stdin.read();if(chunk)return {child,message:JSON.parse(chunk)};await new Promise(r=>setTimeout(r,5));}await promise;throw Error('No native command');};
    return {client,children,command};
}
test('reserva o auxiliar antes da inicialização e ignora respostas com id antigo',async t=>{
    const {client,children,command}=await fixture(t),pending=client.request({action:'list'});
    await assert.rejects(client.request({action:'list'}),/Aguarde/);const {child,message}=await command(pending);
    child.stdout.write(JSON.stringify({id:message.id-1,ok:true,result:['old']})+'\n');assert.equal(client.busy,true);
    child.stdout.write(JSON.stringify({id:message.id,ok:true,result:[{name:'virtual'}]})+'\n');assert.deepEqual(await pending,[{name:'virtual'}]);assert.equal(children.length,1);assert.equal(client.busy,false);
});
test('binário adulterado não inicia o auxiliar',async t=>{
    const {client,children}=await fixture(t,{tamper:true});await assert.rejects(client.request({action:'list'}),/incompatível/);assert.equal(children.length,0);assert.equal(client.busy,false);
});
test('falha do auxiliar libera controles e respostas de processo anterior não interferem',async t=>{
    const {client,command}=await fixture(t),first=client.request({action:'list'}),firstRejected=assert.rejects(first,/encerrado/),{child:old}=await command(first);
    old.emit('exit',1);await firstRejected;assert.equal(client.busy,false);
    const next=client.request({action:'list'}),{child,message}=await command(next);
    old.stdout.emit('data',JSON.stringify({id:message.id,ok:true,result:['obsolete']})+'\n');assert.equal(client.busy,true);
    child.stdout.write(JSON.stringify({id:message.id,ok:true,result:[]})+'\n');assert.deepEqual(await next,[]);
});
test('timeout cancela a sessão sem manter operação pendente',async t=>{
    const {client}=await fixture(t,{timeout:30});await assert.rejects(client.request({action:'list'}),/tempo/);assert.equal(client.busy,false);assert.equal(client.child,null);
});
test('auxiliar nativo tem apenas APIs de consulta/propriedades e nunca persistência ou envio',async()=>{
    const source=await readFile(new URL('../desktop/native/PrintConfigNative.cs',import.meta.url),'utf8');
    assert.doesNotMatch(source,/\b(StartDoc\w*|StartPage\w*|WritePrinter|AddJob\w*|SetPrinter\w*|SetDefaultPrinter\w*|ResetDC\w*)\s*\(/);
    assert.match(source,/CreateICW/);assert.match(source,/DocumentPropertiesW\(owner,printer,Name,output,input,2\|4\|8\)/);
    assert.match(source,/ReadMode\(output,capacity\)/);
});
