import test from 'node:test';
import assert from 'node:assert/strict';
import {mkdtemp,readdir,readFile,writeFile} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import path from 'node:path';
import {openExternalCopy} from '../desktop/external-reader.mjs';
import {validatePrintRequest} from '../desktop/print-config.mjs';
import {JanelaConfiguracaoImpressao} from '../public/js/app/desktop-print-ui.js';

test('cancelar abertura externa preserva o diretório sem gravar a cópia',async()=>{
    const folder=await mkdtemp(path.join(tmpdir(),'4printshops-reader-'));
    const result=await openExternalCopy({}, {showMessageBox:async()=>({response:2})}, Buffer.from('%PDF-1.7\ntest'),'Teste.pdf',folder,{findEdge:async()=> 'C:\\Edge\\msedge.exe'});
    assert.deepEqual(result,{opened:false});assert.deepEqual(await readdir(folder),[]);
});
test('leitor externo recebe cópia exclusiva com bytes atuais e nome validado',async()=>{
    const folder=await mkdtemp(path.join(tmpdir(),'4printshops-reader-')),bytes=Buffer.from('%PDF-1.7\natual');
    let delivered;
    const result=await openExternalCopy({}, {showMessageBox:async()=>({response:0})}, bytes,'Cópia São José.pdf',folder,{
        findEdge:async()=> 'C:\\Edge\\msedge.exe',launch:async(executable,file)=>{delivered={executable,file};}
    });
    assert.equal(result.opened,true);assert.equal(result.reader,'Microsoft Edge');assert.deepEqual(await readFile(delivered.file),bytes);assert.equal(path.basename(delivered.file),'Cópia São José.pdf');
    await assert.rejects(openExternalCopy({}, {},bytes,'../arquivo.pdf',folder),/inválido/);
});
test('falha do leitor escolhido é relatada e conserva a cópia',async()=>{
    const folder=await mkdtemp(path.join(tmpdir(),'4printshops-reader-'));let copy;
    const executable=path.join(folder,'Reader.exe');await writeFile(executable,'fixture');
    const dialog={showMessageBox:async()=>({response:0}),showOpenDialog:async()=>({canceled:false,filePaths:[executable]})};
    await assert.rejects(openExternalCopy({},dialog,Buffer.from('%PDF-1.7\ntest'),'Teste.pdf',folder,{findEdge:async()=>null,launch:async(exe,file)=>{copy=file;throw new Error('Leitor indisponível');}}),/indisponível/);
    assert.ok((await readFile(copy)).length>0);
});
test('o canal de configuração recusa comandos de envio e exige revisão para Hold',()=>{
    for(const action of ['send','send-hold','print','virtual-file'])assert.throws(()=>validatePrintRequest({action}));
    assert.throws(()=>validatePrintRequest({action:'verify-hold',session:'a'.repeat(32),revision:0}));
    assert.deepEqual(validatePrintRequest({action:'verify-hold',session:'a'.repeat(32),revision:1}),{action:'verify-hold',session:'a'.repeat(32),revision:1});
});
test('o botão de envio permanece bloqueado mesmo com uma resposta favorável forjada',()=>{
    const ui=Object.create(JanelaConfiguracaoImpressao.prototype),buttons=new Map();
    ui.busy=false;ui.model={state:{hold:{canSend:true,holdConfirmed:true}}};ui.get=id=>{if(!buttons.has(id))buttons.set(id,{});return buttons.get(id);};
    ui.controls();assert.equal(buttons.get('send-hold').disabled,true);
});
