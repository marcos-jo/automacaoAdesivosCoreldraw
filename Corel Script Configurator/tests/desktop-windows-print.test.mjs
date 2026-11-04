import test from 'node:test';
import assert from 'node:assert/strict';
import {readFile,access} from 'node:fs/promises';
import {WindowsPrintCoordinator,validateDocumentPrint} from '../desktop/windows-print.mjs';
import {validateNativeDocumentRequest,validatePrintRequest} from '../desktop/print-config.mjs';
import {calcularPreviaImpressao} from '../public/js/app/desktop-print-layout.js';

const input={session:'a'.repeat(32),revision:1,scale:120,name:'Trabalho sintético com acentos.pdf',operation:'b'.repeat(32)};
const bytes=Buffer.from('%PDF-1.7 synthetic test data');
function fixture({printer='EPSON teste',fail=false,choose=true}={}){
    const calls=[];let file;
    const state={printer,queue:{driver:printer==='Microsoft Print to PDF'?'Microsoft Print To PDF':'EPSON padrão'}};
    const client={request:async request=>{
        calls.push(request);
        if(request.action==='summary')return state;
        if(request.action==='prepare-document'){file=request.file;assert.deepEqual(await readFile(file),bytes);return {...input,printer,pages:3,token:'c'.repeat(32)};}
        if(request.action==='print-document'){if(fail)throw Error('timeout');return {submitted:true,pages:3,printer,jobId:42};}
        return true;
    }};
    return {calls,client,coordinator:new WindowsPrintCoordinator(client,async()=>choose?'C:\\lab\\virtual.pdf':null),file:()=>file};
}
test('impressão comum não exige Hold, sintético, papel específico, cópia única ou nova confirmação',async()=>{
    const f=fixture();assert.equal((await f.coordinator.run(input,bytes)).submitted,true);
    assert.deepEqual(f.calls.map(c=>c.action),['summary','prepare-document','print-document','cancel-document']);
    assert.equal(f.calls[1].scale,120);assert.equal(f.calls[2].name,input.name);
    await assert.rejects(access(f.file()));
});
test('mesmo clique não duplica; novo clique deliberado cria outro trabalho',async()=>{
    const f=fixture();await f.coordinator.run(input,bytes);
    await assert.rejects(f.coordinator.run(input,bytes),/utilizada/);
    await f.coordinator.run({...input,operation:'d'.repeat(32)},bytes);
    assert.equal(f.calls.filter(c=>c.action==='print-document').length,2);
    assert.notEqual(f.calls[1].file,f.calls[5].file);
});
test('timeout nunca repete nem presume que o trabalho falhou antes do envio',async()=>{
    const f=fixture({fail:true});const result=await f.coordinator.run(input,bytes);
    assert.equal(result.uncertain,true);assert.match(result.message,/Confira a fila/);
    assert.equal(f.calls.filter(c=>c.action==='print-document').length,1);
    await assert.rejects(f.coordinator.run(input,bytes),/utilizada/);
});
test('destino Microsoft é escolhido pelo operador; cancelamento não cria trabalho',async()=>{
    const f=fixture({printer:'Microsoft Print to PDF',choose:false});
    assert.equal((await f.coordinator.run(input,bytes)).canceled,true);
    assert.equal(f.calls.some(c=>c.action==='print-document'),false);
    const next=fixture({printer:'Microsoft Print to PDF'});await next.coordinator.run(input,bytes);
    assert.equal(next.calls.find(c=>c.action==='print-document').output,'C:\\lab\\virtual.pdf');
});
test('sessão ou token inválido no preparo impede o envio',async()=>{
    const f=fixture();const previous=f.client.request;
    f.client.request=async request=>request.action==='prepare-document'?{printer:'outra fila',pages:1,token:'c'.repeat(32)}:previous(request);
    await assert.rejects(f.coordinator.run(input,bytes),/comprovados/);assert.equal(f.calls.some(c=>c.action==='print-document'),false);
});
test('validação IPC não aceita caminhos do renderer nem comandos de envio pelo canal de configuração',()=>{
    assert.deepEqual(validateDocumentPrint(input),input);
    for(const request of [{...input,file:'C:\\cliente.pdf'},{...input,scale:401},{...input,scale:0},{...input,operation:'id'},{...input,name:'bad\0name'}])assert.throws(()=>validateDocumentPrint(request));
    assert.throws(()=>validatePrintRequest({action:'print-document',...input}));
    assert.throws(()=>validateNativeDocumentRequest({action:'print-document',session:input.session,revision:1,token:'c'.repeat(32),name:'PDF',output:'relative.pdf'}));
});
test('prévia usa página real, paisagem e corte sem alterar a decisão do operador',()=>{
    const source={width:400,height:200};
    const layout=calcularPreviaImpressao({width:210,height:297},120,null,source);
    assert.equal(layout.content.width,480);assert.equal(layout.content.height,240);assert.equal(layout.outsidePaper,true);
    assert.throws(()=>calcularPreviaImpressao({width:210,height:297},100,null,{width:0,height:297}));
});
test('caminho Windows não contém comandos proprietários, reconstrução privada ou laço de cópias',async()=>{
    const source=await readFile(new URL('../desktop/native/WindowsPrintJob.cs',import.meta.url),'utf8');
    assert.doesNotMatch(source,/SHKM|DocumentKO|WritePrinter|VerifyHold|SetPrinter|SetDefaultPrinter/);
    assert.match(source,/CreateDCW\("WINSPOOL",session.Name,null,memory\)/);
    assert.match(source,/mode.SequenceEqual\(confirmed\)/);assert.match(source,/FPDFPage_Flatten\(page,1\)/);
    assert.match(source,/current=null; \/\/ Consumir/);assert.match(source,/FPDF_SetPrintMode\(0\)/);
});
