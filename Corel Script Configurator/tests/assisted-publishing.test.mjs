import test from 'node:test';
import assert from 'node:assert/strict';
import {generateKeyPairSync,createHash} from 'node:crypto';
import {mkdtemp,writeFile,readFile,rm,mkdir,symlink} from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import {publishRelease,auditRepository,packagedIdentity} from '../scripts/assisted/publish.mjs';
import {keyId,LAB_PRODUCT,authenticateEnvelope} from '../desktop/assisted-manifest.mjs';
import {outsideWorkspace} from '../scripts/assisted/keys.mjs';
import {createRequire} from 'node:module';
const require=createRequire(import.meta.url),asar=require('@electron/asar');
async function fixture(t){
    const folder=await mkdtemp(path.join(os.tmpdir(),'4ps-publishing-unit-'));t.after(()=>rm(folder,{recursive:true,force:true}));
    const {privateKey,publicKey}=generateKeyPairSync('ed25519'),pem=publicKey.export({type:'spki',format:'pem'});
    const policy={product:LAB_PRODUCT,channel:'internal-lab',keys:[{id:keyId(pem),pem}]};
    const installer=path.join(folder,'Setup.exe');await writeFile(installer,'MZ bytes sintéticos de laboratório');
    return {folder,installer,privateKey,policy,root:path.join(folder,'repository'),version:'1.0.0-experimental.6',notes:['Primeira publicação']};
}
test('publicação copia arquivo real e publica manifesto assinado por último; sequência é crescente',async t=>{
    const f=await fixture(t),first=await publishRelease(f),r=authenticateEnvelope(await readFile(first.manifest),f.policy);
    assert.equal(r.data.sequence,1);assert.deepEqual(await readFile(first.installer),await readFile(f.installer));
    assert.equal(createHash('sha512').update(await readFile(first.installer)).digest('hex'),r.data.installer.sha512);
    assert.equal((await publishRelease(f)).sequence,2);
    const before=await readFile(first.manifest);
    await assert.rejects(publishRelease({...f,sequence:1}),/antiga|sequência/i);
    await assert.rejects(publishRelease({...f,version:'1.0.0-experimental.5'}),/downgrade/);
    await writeFile(f.installer,'MZ arquivo diferente');await assert.rejects(publishRelease(f),/bytes diferentes/);
    assert.deepEqual(await readFile(first.manifest),before,'Falhas não modificam a publicação vigente');
});
test('chave incorreta, notas inválidas e segredos no repositório impedem publicar',async t=>{
    const f=await fixture(t);
    await assert.rejects(publishRelease({...f,privateKey:generateKeyPairSync('ed25519').privateKey}),/Chave privada/);
    await assert.rejects(publishRelease({...f,notes:[{}]}),/Notas/);
    await writeFile(path.join(f.root,'publication.dpapi'),'não é uma chave real');
    await assert.rejects(auditRepository(f.root),/não publicável/);
});
test('repositório rejeita junctions; chave não pode ser armazenada dentro do Git',async t=>{
    const f=await fixture(t);await mkdir(f.root);const target=path.join(f.folder,'elsewhere');await mkdir(target);
    await symlink(target,path.join(f.root,'desvio'),process.platform==='win32'?'junction':'dir');
    await assert.rejects(auditRepository(f.root),/Link/);
    await assert.rejects(outsideWorkspace(path.resolve('.')),/fora do repositório/);
    await assert.rejects(outsideWorkspace(path.resolve('..')),/fora do repositório/);
});
test('publicador recusa pacotes dos laboratórios anteriores e exige a raiz incorporada',async t=>{
    const f=await fixture(t),source=path.join(f.folder,'source'),resources=path.join(f.folder,'build','win-unpacked','resources');await mkdir(source);await mkdir(resources,{recursive:true});
    const create=async pkg=>{await writeFile(path.join(source,'package.json'),JSON.stringify(pkg));await asar.createPackage(source,path.join(resources,'app.asar'));asar.uncache(path.join(resources,'app.asar'));};
    await create({name:'4printshops-open-test',version:f.version,assistedUpdates:f.policy});
    await assert.rejects(packagedIdentity(path.join(f.folder,'build'),f.policy),/Identidade/);
    await create({name:'4printshops-assisted-lab',desktopAssistedLab:true,version:f.version,assistedUpdates:f.policy});
    assert.equal((await packagedIdentity(path.join(f.folder,'build'),f.policy)).version,f.version);
    await create({name:'4printshops-assisted-lab',desktopAssistedLab:true,version:f.version});
    await assert.rejects(packagedIdentity(path.join(f.folder,'build'),f.policy),/raiz pública/);
});
test('configuração principal não aceita confiança do laboratório, mesmo com variável de laboratório ativa',async t=>{
    const f=await fixture(t),file=path.join(f.folder,'public-trust.json');await writeFile(file,JSON.stringify(f.policy));
    const original={trust:process.env.FOURPRINT_PUBLIC_TRUST,lab:process.env.FOURPRINT_ASSISTED_LAB,version:process.env.FOURPRINT_LAB_VERSION};
    try{
        process.env.FOURPRINT_PUBLIC_TRUST=file;process.env.FOURPRINT_ASSISTED_LAB='1';process.env.FOURPRINT_LAB_VERSION=f.version;
        const main=require.resolve('../electron-builder.assisted.cjs'),lab=require.resolve('../electron-builder.assisted-lab.cjs');
        delete require.cache[main];delete require.cache[lab];
        assert.throws(()=>require(main),/identidade/);
        const config=require(lab);assert.equal(config.appId,LAB_PRODUCT);assert.equal(config.nsis.include,null);
        assert.equal(config.extraMetadata.assistedUpdates.product,LAB_PRODUCT);
        await writeFile(file,JSON.stringify({...f.policy,product:'br.com.4printshops.desktop',channel:'internal'}));delete require.cache[main];
        const production=require(main);assert.equal(production.appId,'br.com.4printshops.desktop');assert.equal(production.productName,'4PrintShops');assert.equal(production.nsis.include,'build/desktop-installer.nsh');
    }finally{
        for(const [key,value]of [['FOURPRINT_PUBLIC_TRUST',original.trust],['FOURPRINT_ASSISTED_LAB',original.lab],['FOURPRINT_LAB_VERSION',original.version]])if(value===undefined)delete process.env[key];else process.env[key]=value;
    }
});
