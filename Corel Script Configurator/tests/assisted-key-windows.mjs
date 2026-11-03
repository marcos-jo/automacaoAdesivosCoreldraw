// Teste local de DPAPI/ACL em pasta NOVA e temporária; não cria certificados.
import assert from 'node:assert/strict';
import {randomUUID,createPublicKey} from 'node:crypto';
import {readFile,writeFile,unlink,rmdir} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import path from 'node:path';
import {generatePublishingKey,loadPublishingKey} from '../scripts/assisted/keys.mjs';
import {keyId} from '../desktop/assisted-manifest.mjs';

const folder=path.join(tmpdir(),'4printshops-key-test-'+randomUUID()),file=path.join(folder,'publication.dpapi');
try{
    const info=await generatePublishingKey(folder,{laboratory:true}),key=await loadPublishingKey(file);
    assert.equal(keyId(createPublicKey(key).export({format:'pem',type:'spki'})),info.keyId);
    const encrypted=await readFile(file);assert.equal(encrypted.includes(Buffer.from('PRIVATE KEY')),false);
    await writeFile(file,Buffer.alloc(encrypted.length));await assert.rejects(loadPublishingKey(file),/Proteção local/);
    console.log('PASSOU: chave local em disco fixo fora do Git, DPAPI real ida/volta e recusa de chave protegida adulterada; nenhum certificado criado.');
}finally{
    // Sem exclusão recursiva: somente dois arquivos criados por este teste.
    await unlink(file).catch(()=>{});await unlink(path.join(folder,'public-trust.json')).catch(()=>{});await rmdir(folder).catch(()=>{});
}
