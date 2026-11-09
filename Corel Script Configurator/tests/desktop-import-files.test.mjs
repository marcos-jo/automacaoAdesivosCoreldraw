import test from 'node:test';
import assert from 'node:assert/strict';
import {mkdtemp,writeFile,readFile,rm,open} from 'node:fs/promises';
import path from 'node:path';
import {tmpdir} from 'node:os';
import {readImportedFiles,openImportedFiles,IMPORT_TOTAL} from '../desktop/import-files.mjs';
const temp=async run=>{const folder=await mkdtemp(path.join(tmpdir(),'4printshops-import-test-'));try{return await run(folder);}finally{await rm(folder,{recursive:true,force:true});}};
const bytes=Buffer.from('conteudo original seguro');
test('seleção mista preserva ordem/bytes PDF e imagem; converte somente Word e devolve relatório',()=>temp(async folder=>{
    const files=['um.pdf','dois.PNG','três.docx','quatro.DOC'].map(name=>path.join(folder,name));for(const file of files)await writeFile(file,bytes);
    const converted=[];
    const result=await readImportedFiles(files,{convert:async(file,{signal})=>{converted.push(file);assert.equal(signal,undefined);return {name:path.basename(file)+'.pdf',bytes,report:{pages:2,fonts:[]}};}});
    assert.deepEqual(result.map(f=>f.name),['um.pdf','dois.PNG','três.docx.pdf','quatro.DOC.pdf']);
    assert.deepEqual(converted,files.slice(2));assert.deepEqual(Buffer.from(result[0].bytes),bytes);assert.equal(result[2].wordReport.originalName,'três.docx');
    for(const file of files)assert.deepEqual(await readFile(file),bytes);
}));
test('extensão incompatível e excesso de arquivos recusam todo o lote antes de converter',()=>temp(async folder=>{
    const file=path.join(folder,'a.docx'),bad=path.join(folder,'b.exe');await writeFile(file,bytes);await writeFile(bad,bytes);
    let count=0;const convert=async()=>{count++;};
    await assert.rejects(readImportedFiles([file,bad],{convert}),/PDF/);assert.equal(count,0);
    await assert.rejects(readImportedFiles(Array(129).fill(file),{convert}),/128/);assert.equal(count,0);
    await assert.rejects(readImportedFiles([]),/128/);
}));
test('tamanho agregado de entrada é conferido antes de converter e saída também tem limite',()=>temp(async folder=>{
    const file=path.join(folder,'a.docx'),big=path.join(folder,'b.pdf');await writeFile(file,bytes);const handle=await open(big,'w');await handle.truncate(IMPORT_TOTAL);await handle.close();
    let count=0;await assert.rejects(readImportedFiles([file,big],{convert:async()=>{count++;}}),/512/);assert.equal(count,0);
    await assert.rejects(readImportedFiles([file],{convert:async()=>({name:'a.pdf',bytes:{length:IMPORT_TOTAL+1},report:{}})}),/512/);
}));
test('falha em um Word identifica a origem, não retorna lote parcial e permite nova seleção',()=>temp(async folder=>{
    const a=path.join(folder,'ok.pdf'),b=path.join(folder,'erro.docx');await writeFile(a,bytes);await writeFile(b,bytes);
    await assert.rejects(readImportedFiles([a,b],{convert:async()=>{throw Error('LibreOffice indisponível');}}),/erro.docx.*LibreOffice/);
    assert.equal((await readImportedFiles([a])).length,1);assert.deepEqual(await readFile(a),bytes);
}));
test('cancelamento antes/depois de Word impede resultado parcial ou tardio',()=>temp(async folder=>{
    const file=path.join(folder,'a.docx');await writeFile(file,bytes);const controller=new AbortController();
    await assert.rejects(readImportedFiles([file],{signal:controller.signal,convert:async(_file,{signal})=>{assert.equal(signal,controller.signal);controller.abort();return {name:'a.pdf',bytes,report:{}};}}),/cancelada/);
    let called=false;await assert.rejects(readImportedFiles([file],{signal:controller.signal,convert:async()=>{called=true;}}),/cancelada/);assert.equal(called,false);
}));
test('diálogo unificado oferece todos os tipos, admite múltiplos e preserva cancelamento',async()=>{
    let options;
    assert.deepEqual(await openImportedFiles(null,{showOpenDialog:async(_window,value)=>{options=value;return {canceled:true};}},true),[]);
    assert.equal(options.title,'Importar arquivos');assert.deepEqual(options.filters[0].extensions,['pdf','jpg','jpeg','png','docx','doc']);assert.ok(options.properties.includes('multiSelections'));
    await assert.rejects(openImportedFiles(null,{showOpenDialog:async()=>({filePaths:['a.pdf','b.pdf']})},false),/somente um/);
});
test('eventos antecipam seleção sem caminhos e entregam progresso antes de terminar Word lento',()=>temp(async folder=>{
    const files=['a.pdf','lento.docx'].map(n=>path.join(folder,n));for(const f of files)await writeFile(f,bytes);
    const events=[];let release;const gate=new Promise(r=>{release=r;});
    const task=readImportedFiles(files,{onSelection:items=>events.push(['selection',items]),onProgress:(...args)=>events.push(['stage',...args]),onEntry:index=>events.push(['entry',index]),convert:async()=>{await gate;return{name:'lento.pdf',bytes,report:{}};}});
    while(!events.some(e=>e[0]==='stage'&&e[2]==='word'))await new Promise(r=>setImmediate(r));
    assert.deepEqual(events[0][1],[{index:0,name:'a.pdf',type:'PDF'},{index:1,name:'lento.docx',type:'DOCX'}]);assert.ok(events.some(e=>e[0]==='entry'&&e[1]===0));assert.ok(!events.some(e=>e[0]==='entry'&&e[1]===1));release();assert.equal((await task).length,2);
}));
test('cancelar conversão impede publicação tardia de arquivo e devolução parcial',()=>temp(async folder=>{
    const file=path.join(folder,'a.docx');await writeFile(file,bytes);const controller=new AbortController(),entries=[];
    await assert.rejects(readImportedFiles([file],{signal:controller.signal,onEntry:(...args)=>entries.push(args),convert:async()=>{controller.abort();return{name:'a.pdf',bytes,report:{}};}}),/cancelada/);assert.deepEqual(entries,[]);
}));
