import test from 'node:test';import assert from 'node:assert/strict';
import {mkdtemp,writeFile,readFile,rename} from 'node:fs/promises';import {tmpdir} from 'node:os';import path from 'node:path';
import {save} from '../desktop/update-transaction.mjs';
async function fixture(){const root=await mkdtemp(path.join(tmpdir(),'4ps-journal-retry-')),file=path.join(root,'journal.json');await writeFile(file,'{"phase":"prepared"}');return file;}
test('duas recusas EPERM: grava uma vez, aguarda 50/100 ms e renomeia na terceira tentativa',async()=>{
 const file=await fixture(),delays=[];let writes=0,calls=0;
 await save(file,{phase:'copying'},{writeTemp:async(...args)=>{writes++;await writeFile(...args);},renameFile:async(a,b)=>{calls++;assert.equal(await readFile(file,'utf8'),'{"phase":"prepared"}');if(calls<=2)throw Object.assign(Error('ocupado'),{code:'EPERM'});await rename(a,b);},wait:async ms=>{delays.push(ms);}});
 assert.equal(writes,1);assert.equal(calls,3);assert.deepEqual(delays,[50,100]);assert.deepEqual(JSON.parse(await readFile(file,'utf8')),{phase:'copying'});
});
test('EPERM persistente: cinco tentativas, erro original propagado e diário anterior/temporário preservados',async()=>{
 const file=await fixture(),error=Object.assign(Error('ocupado'),{code:'EPERM'}),delays=[];let calls=0,writes=0;
 await assert.rejects(save(file,{phase:'copying'},{writeTemp:async(...args)=>{writes++;await writeFile(...args);},renameFile:async()=>{calls++;throw error;},wait:async ms=>{delays.push(ms);}}),e=>e===error);
 assert.equal(writes,1);assert.equal(calls,5);assert.deepEqual(delays,[50,100,200,400]);assert.equal(await readFile(file,'utf8'),'{"phase":"prepared"}');assert.deepEqual(JSON.parse(await readFile(file+'.tmp','utf8')),{phase:'copying'});
});
test('erros diferentes de EPERM não são repetidos',async()=>{
 for(const code of ['EACCES','ENOENT','EIO','EBUSY']){const file=await fixture(),error=Object.assign(Error(code),{code});let calls=0;
 await assert.rejects(save(file,{phase:'copying'},{renameFile:async()=>{calls++;throw error;},wait:async()=>assert.fail('Espera indevida')}),e=>e===error);assert.equal(calls,1);assert.equal(await readFile(file,'utf8'),'{"phase":"prepared"}');}
});
test('falha ao gravar temporário não inicia rename nem espera',async()=>{
 const file=await fixture(),error=Object.assign(Error('disco'),{code:'ENOSPC'});
 await assert.rejects(save(file,{}, {writeTemp:async()=>{throw error;},renameFile:async()=>assert.fail('Rename indevido'),wait:async()=>assert.fail('Espera indevida')}),e=>e===error);assert.equal(await readFile(file,'utf8'),'{"phase":"prepared"}');
});
test('caminho padrão sem injeção persiste JSON normalmente',async()=>{const file=await fixture();await save(file,{phase:'installed'});assert.deepEqual(JSON.parse(await readFile(file,'utf8')),{phase:'installed'});});
