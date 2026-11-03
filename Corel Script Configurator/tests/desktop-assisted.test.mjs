import test from 'node:test';
import assert from 'node:assert/strict';
import {generateKeyPairSync,sign,createHash} from 'node:crypto';
import {mkdtemp,readFile,writeFile,rm,readdir} from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import http from 'node:http';
import {keyId,authenticateEnvelope,validateRelease,sourceURL,LAB_PRODUCT} from '../desktop/assisted-manifest.mjs';
import {AssistedUpdates,configuredAssisted} from '../desktop/assisted-updates.mjs';
import {assistedAction} from '../desktop/assisted-ui.mjs';
import {getManifest} from '../desktop/assisted-http.mjs';

function fixture(){
    const {privateKey,publicKey}=generateKeyPairSync('ed25519');
    const pem=publicKey.export({format:'pem',type:'spki'}),policy={product:LAB_PRODUCT,channel:'internal-lab',keys:[{id:keyId(pem),pem}]};
    const installer=Buffer.from('MZ instalador sintético, nunca executado');
    const data={schema:1,product:policy.product,channel:policy.channel,platform:'win32',arch:'x64',version:'1.0.0-experimental.6',sequence:2,
        issuedAt:new Date(Date.now()-60000).toISOString(),expiresAt:new Date(Date.now()+3600000).toISOString(),notes:['Teste de atualização'],
        installer:{path:'releases/1.0.0-experimental.6/Setup.exe',size:installer.length,sha512:createHash('sha512').update(installer).digest('hex')}};
    const envelope=(change={})=>{const bytes=Buffer.from(JSON.stringify({...data,...change}));return Buffer.from(JSON.stringify({keyId:policy.keys[0].id,payload:bytes.toString('base64url'),signature:sign(null,bytes,privateKey).toString('base64url')}));};
    return {policy,data,envelope,installer};
}
async function lab(t){
    const f=fixture(),folder=await mkdtemp(path.join(os.tmpdir(),'4ps-assisted-unit-'));
    let manifest=f.envelope(),installer=f.installer,status=200,mode='normal';
    const server=http.createServer((req,res)=>{
        if(mode==='timeout'){res.writeHead(200);res.write(' ');return;}
        if(mode==='redirect'){res.writeHead(302,{Location:'/elsewhere'});res.end();return;}
        if(mode==='chunked'&&!req.url.startsWith('/channels/')){res.writeHead(200);const half=Math.floor(installer.length/2);res.write(installer.subarray(0,half));setTimeout(()=>res.end(installer.subarray(half)),20);return;}
        res.writeHead(status);res.end(req.url.startsWith('/channels/')?manifest:installer);
    });
    await new Promise(r=>server.listen(0,'127.0.0.1',r));
    const source={url:`http://127.0.0.1:${server.address().port}/`,allowHttp:true};
    const controller=new AssistedUpdates({policy:f.policy,source,currentVersion:'1.0.0-experimental.5',folder,networkOptions:{timeout:200}});
    t.after(async()=>{server.closeAllConnections();await new Promise(r=>server.close(r));await rm(folder,{recursive:true,force:true});});
    return {...f,controller,source,folder,set manifest(v){manifest=v;},get installer(){return installer;},set installer(v){installer=v;},set status(v){status=v;},set mode(v){mode=v;}};
}
test('manifesto autentica bytes exatos, produto, canal, plataforma e caminho autorizado',()=>{
    const f=fixture(),release=authenticateEnvelope(f.envelope(),f.policy);
    assert.equal(validateRelease(release,{currentVersion:'1.0.0-experimental.5'}),true);
    for(const change of [{product:'outro'},{channel:'experimental'},{arch:'arm64'},{platform:'linux'},{version:'lixo'},{sequence:0},
        {installer:{...f.data.installer,path:'../Setup.exe'}},{installer:{...f.data.installer,path:'releases/1.0.0-experimental.6/%2e%2e.exe'}},
        {installer:{...f.data.installer,path:'https://outro/Setup.exe'}},{installer:{...f.data.installer,sha512:'0'}},{notes:[{}]}])assert.throws(()=>authenticateEnvelope(f.envelope(change),f.policy));
    const bad=JSON.parse(f.envelope());bad.payload=Buffer.from(JSON.stringify({...f.data,notes:['alterado']})).toString('base64url');
    assert.throws(()=>authenticateEnvelope(Buffer.from(JSON.stringify(bad)),f.policy),/Assinatura/);
    assert.throws(()=>authenticateEnvelope(f.envelope(),fixture().policy),/Chave/);
});
test('recusa downgrade, replay, sequência reutilizada, validade e relógio adiantado',()=>{
    const f=fixture(),r=authenticateEnvelope(f.envelope(),f.policy);
    assert.throws(()=>validateRelease(r,{currentVersion:'1.0.0-experimental.7'}),/downgrade/);
    assert.equal(validateRelease(r,{currentVersion:r.data.version}),false);
    for(const floor of [{sequence:3},{sequence:2,digest:'outro'}])assert.throws(()=>validateRelease(r,{currentVersion:'1.0.0',floor}),/antiga/);
    assert.throws(()=>validateRelease(r,{currentVersion:'0.1.0',now:Date.now()+7200000}),/expirada/);
    assert.throws(()=>validateRelease(r,{currentVersion:'0.1.0',now:Date.now()-7200000}),/relógio/);
    assert.throws(()=>authenticateEnvelope(f.envelope({expiresAt:new Date(Date.now()+32*86400000).toISOString()}),f.policy),/Validade/);
});
test('origem não aceita credenciais, protocolos de disco, consulta ou HTTP implícito',()=>{
    for(const url of ['file:///C:/repo/','https://a/?x=1','https://u:p@a/','https://a/#x','https://a/sem-barra','http://a/'])assert.throws(()=>sourceURL(url));
    assert.equal(sourceURL('http://servidor/repo/',{allowHttp:true}).hostname,'servidor');
});
test('publicação real HTTP, adiamento, download, cache e nova conferência antes de revelar',async t=>{
    const s=await lab(t),c=s.controller;
    await c.check();assert.equal(c.state.phase,'available');assert.equal(c.file,null);
    const floor=JSON.parse(await readFile(path.join(s.folder,'publication-state.json')));assert.equal(floor.sequence,2);
    await c.download();assert.equal(c.state.phase,'ready');assert.deepEqual(await readFile(await c.preparedFile()),s.installer);
    await writeFile(c.file,Buffer.alloc(s.installer.length));
    await assert.rejects(c.preparedFile(),/modificado/);assert.equal(c.state.phase,'error');
    await c.check();await c.download();assert.equal(c.state.phase,'ready');
});
test('manifesto adulterado, tamanho/hash divergentes, rede e estado local inválido bloqueiam oferta',async t=>{
    const s=await lab(t),c=s.controller;s.manifest=Buffer.from('{}');await c.check();assert.equal(c.state.phase,'error');
    await assert.rejects(c.download(),/Nenhuma/);
    s.manifest=s.envelope();await c.check();s.installer=Buffer.alloc(s.installer.length);await c.download();assert.equal(c.state.phase,'error');
    assert.deepEqual(await readdir(path.join(s.folder,'installers')),[]);
    s.installer=Buffer.from('curto');await c.check();await c.download();assert.equal(c.state.phase,'error');
    s.status=503;await c.check();assert.equal(c.state.phase,'error');
    s.status=200;await writeFile(path.join(s.folder,'publication-state.json'),'{}');await c.check();assert.match(c.state.message,/Estado local/);
});
test('timeout absoluto inclui corpo lento; redirecionamento não é seguido',async t=>{
    const s=await lab(t);s.mode='timeout';
    await assert.rejects(getManifest(new URL(s.source.url),{timeout:60}));
    s.mode='redirect';await s.controller.check();assert.equal(s.controller.state.phase,'error');
});
test('sequência persistida recusa manifesto antigo mesmo com assinatura válida',async t=>{
    const s=await lab(t);await s.controller.check();s.manifest=s.envelope({sequence:1});await s.controller.check();assert.match(s.controller.state.message,/antiga/);
});
test('progresso intermediário e publicação expirada após download são tratados antes do preparo',async t=>{
    const s=await lab(t),events=[];s.controller.on('state',state=>events.push(state));s.mode='chunked';
    await s.controller.check();await s.controller.download();
    assert.ok(events.some(s=>s.phase==='downloading'&&s.percent>0&&s.percent<100));
    s.controller.now=()=>Date.now()+7200000;await assert.rejects(s.controller.preparedFile(),/expirada/);assert.equal(s.controller.state.phase,'error');
});
test('PDFs pendentes, cancelar e alteração durante confirmação impedem revelar; nunca executa',async t=>{
    const s=await lab(t);await s.controller.check();await s.controller.download();
    let pending=true,answer=1,reveals=0,locks=[],change=false;
    const ui=assistedAction({}, {},{getWindow:()=>null,configure:async()=>s.controller,publish:()=>{},canPrepare:async()=>!pending,lock:v=>locks.push(v),
        shell:{showItemInFolder:()=>reveals++},dialog:{showMessageBox:async(_w,options)=>{if(change&&options.type==='question')pending=true;return {response:answer};}}});
    await ui('prepare');assert.equal(reveals,0);pending=false;await ui('prepare');assert.equal(reveals,0);
    answer=0;change=true;await ui('prepare');assert.equal(reveals,0);change=false;pending=false;await ui('prepare');assert.equal(reveals,1);
    assert.deepEqual(locks,[true,false,true,false,true,false,true,false]);
    await assert.rejects(ui('execute'),/inválida/);ui.destroy();
});
test('inicialização silenciosa não abre diálogos ou baixa e base sem confiança não consulta rede',async t=>{
    const s=await lab(t);let dialogs=0;
    const ui=assistedAction({}, {},{configure:async()=>s.controller,dialog:{showMessageBox:()=>dialogs++}});
    await ui.start();assert.equal(s.controller.state.phase,'available');assert.equal(dialogs,0);assert.equal(s.controller.file,null);ui.destroy();
    assert.equal(await configuredAssisted({},{}),null);
});
