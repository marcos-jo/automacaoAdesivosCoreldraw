import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import {execFileSync,spawn} from 'node:child_process';
import {fileURLToPath} from 'node:url';
import {inspectStorage,main} from '../scripts/storage-cleanup.mjs';

const app=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
const testBase=path.join(app,'.desktop-test','storage-cleanup-tests');
fs.mkdirSync(testBase,{recursive:true});
// Não há teardown destrutivo: fixtures pequenas ficam disponíveis para auditoria.
function fixture(){
 const root=fs.mkdtempSync(path.join(testBase,'case-'));
 const put=(p,data='synthetic')=>{const full=path.join(root,p);fs.mkdirSync(path.dirname(full),{recursive:true});fs.writeFileSync(full,data);return full;};
 put('package.json',JSON.stringify({name:'corel-script-configurator'}));put('AGENTS.md','synthetic');
 put('.gitignore','.desktop-test/\ndist-desktop/\n');
 execFileSync('git',['init','--quiet',root]);
 const git=(...args)=>execFileSync('git',['-C',root,...args],{stdio:'pipe'});
 const setup=put('dist-desktop/hold-native-test/4PrintShops-Hold-Test-1.0.0-experimental.33-Windows-x64-Setup.exe','MZsynthetic-setup');
 const payload='.desktop-test/release33/payload';
 put(payload+'/4PrintShops Hold Test.exe','MZsynthetic-exe');put(payload+'/resources/app.asar','synthetic-asar');
 const run='.desktop-test/autoscroll/chrome-approved/profile-123';
 put(run+'/Local State','{}');put(run+'/Default/Preferences','user-settings');put(run+'/Default/Cache/data_0','cache');
 return {root,put,git,setup,payload,run,inspect:io=>inspectStorage({appRoot:root,io:io||fs})};
}
function snapshot(root){
 const out={};function walk(p){const s=fs.lstatSync(p);if(s.isSymbolicLink()){out[p]='link:'+fs.readlinkSync(p);return;}if(s.isDirectory()){out[p]='dir';for(const n of fs.readdirSync(p))walk(path.join(p,n));}else out[p]=crypto.createHash('sha256').update(fs.readFileSync(p)).digest('hex');}walk(root);return out;
}
test('simulação seleciona apenas grupos conhecidos e não altera nenhum arquivo/diretório',()=>{
 const f=fixture();f.put('public/client.pdf','customer');f.put('.desktop-test/unknown/payload/file.bin');
 const before=snapshot(f.root),r=f.inspect();assert.equal(r.mode,'simulation-only');assert.ok(r.recoverableBytes>0);
 assert.ok(r.candidates.some(p=>p.path.endsWith('app.asar')));assert.ok(r.candidates.some(p=>p.path.endsWith('data_0')));
 assert.ok(!r.candidates.some(p=>/unknown|client.pdf|Preferences/.test(p.path)));
 assert.deepEqual(snapshot(f.root),before);assert.ok(r.candidates.every(p=>p.sha256.length===64));
});
test('código, Git, deps, runtime, cache, Setups e win-unpacked nunca entram na seleção',()=>{
 const f=fixture();for(const p of ['src/a.js','tests/test.mjs','docs/a.md','node_modules/a.bin','build/libreoffice/a.bin','build/libreoffice-cache/a.bin','dist-desktop/hold-native-test/win-unpacked/a.bin','dist-desktop/hold-native-test/4PrintShops-Hold-Test-1.0.0-experimental.32-Windows-x64-Setup.exe'])f.put(p);
 const r=f.inspect();assert.ok(r.candidates.every(p=>p.path.startsWith(path.join(f.root,'.desktop-test'))));
 assert.ok(!r.candidates.some(p=>p.path===f.setup));
});
test('backups misturados, manifestos, evidências e documentos de clientes são protegidos',()=>{
 const f=fixture();for(const p of ['profile-backup/data.bin','before/data.bin','manifest.bin','results.json','photo.png','customer.pdf','evidence.bin','fix.patch'])f.put(f.payload+'/'+p);
 const before=snapshot(f.root),r=f.inspect();assert.ok(r.protected.length>=8);assert.ok(r.candidates.every(p=>!r.protected.some(x=>x.path===p.path)));assert.deepEqual(snapshot(f.root),before);
});
test('arquivos rastreados continuam protegidos mesmo dentro de um payload ignorado',()=>{
 const f=fixture(),p=f.put(f.payload+'/tracked.bin');f.git('add','--force',p);const r=f.inspect();
 assert.ok(r.protected.some(x=>x.path===p&&x.reason.includes('rastreado')));assert.ok(!r.candidates.some(x=>x.path===p));
});
test('Setup ausente ou sem cabeçalho de executável recusa o grupo',()=>{
 const f=fixture();f.put(path.relative(f.root,f.setup),'invalid');const r=f.inspect();assert.ok(!r.candidates.some(x=>x.category==='payloads'));
 assert.ok(r.groups.find(x=>x.category==='payloads'&&x.path.endsWith('payload')).status==='excluded');
});
test('raízes inexistentes são registradas, não criadas',()=>{
 const f=fixture(),before=snapshot(f.root),r=f.inspect();assert.ok(r.excluded.some(e=>e.reason.includes('inexistente')));assert.deepEqual(snapshot(f.root),before);
});
test('perfil com nome parecido fora do cenário permitido permanece protegido',()=>{
 const f=fixture();f.put('.desktop-test/private-user/profile-999/Default/Cache/data');f.put('.desktop-test/autoscroll/chrome-approved/profile-backup/Cache/data');
 const r=f.inspect();assert.ok(!r.candidates.some(x=>/private-user|profile-backup/.test(x.path)));
});
test('links/junctions para fora do repositório não são percorridos',()=>{
 const f=fixture(),other=fixture();other.put('external.bin','external');const link=path.join(f.root,f.payload,'outside');
 fs.symlinkSync(other.root,link,process.platform==='win32'?'junction':'dir');const before=snapshot(other.root),r=f.inspect();
 assert.ok(r.excluded.some(x=>x.reason.includes('Link')||x.reason.includes('link')));assert.ok(!r.candidates.some(x=>x.path.startsWith(path.join(f.root,f.payload))));assert.deepEqual(snapshot(other.root),before);
});
test('junction no ancestral de perfil permitido bloqueia a descoberta',()=>{
 const f=fixture(),other=fixture();const parent=path.join(f.root,'.desktop-test/nup25');fs.mkdirSync(parent,{recursive:true});
 fs.symlinkSync(other.root,path.join(parent,'chrome'),process.platform==='win32'?'junction':'dir');const r=f.inspect();assert.ok(!r.candidates.some(x=>x.path.includes('nup25')));
});
test('hardlinks são preservados',()=>{
 const f=fixture(),original=f.put('protected.bin');fs.linkSync(original,path.join(f.root,f.payload,'hardlink.bin'));const r=f.inspect();assert.ok(r.protected.some(x=>x.reason.includes('hardlink')));
});
test('arquivo bloqueado/inacessível recusa o grupo e mantém os arquivos',()=>{
 const f=fixture(),blocked=path.join(f.root,f.payload,'resources/app.asar');const before=snapshot(f.root);
 const io={...fs,openSync(p,...args){if(p===blocked)throw Object.assign(Error('synthetic EACCES'),{code:'EACCES'});return fs.openSync(p,...args);}};
 const r=f.inspect(io);assert.ok(!r.candidates.some(x=>x.path.startsWith(path.join(f.root,f.payload))));assert.deepEqual(snapshot(f.root),before);
});
test('arquivo criado durante inspeção invalida todo o grupo',()=>{
 const f=fixture();let changed=false;const target=path.join(f.root,f.payload,'resources/app.asar');
 const io={...fs,openSync(p,...args){if(p===target&&!changed){changed=true;f.put(f.payload+'/new.bin','new');}return fs.openSync(p,...args);}};
 const r=f.inspect(io);assert.ok(changed);assert.ok(!r.candidates.some(x=>x.path.startsWith(path.join(f.root,f.payload))));assert.ok(r.excluded.some(x=>x.reason.includes('mudou')));
});
test('alteração de arquivo durante leitura invalida o grupo',()=>{
 const f=fixture();let changed=false;const target=path.join(f.root,f.payload,'resources/app.asar');
 const io={...fs,openSync(p,...args){if(p===target&&!changed){changed=true;f.put(f.payload+'/4PrintShops Hold Test.exe','MZchanged');}return fs.openSync(p,...args);}};
 const r=f.inspect(io);assert.ok(!r.candidates.some(x=>x.path.startsWith(path.join(f.root,f.payload))));
});
test('parâmetros destrutivos e caminhos livres são recusados antes de inspecionar',()=>{
 for(const arg of ['--apply','--delete','--force','../outside','C:\\outside','--manifest=anything'])assert.throws(()=>main([arg]),/Somente simulação/);
});
test('Setup alterado após conferir hash invalida os candidatos dependentes',()=>{
 const f=fixture();let changed=false;const target=path.join(f.root,f.payload,'resources/app.asar');
 const io={...fs,openSync(p,...args){if(p===target&&!changed){changed=true;f.put(path.relative(f.root,f.setup),'MZmodified-setup');}return fs.openSync(p,...args);}};
 const r=f.inspect(io);assert.ok(changed);assert.ok(!r.candidates.some(x=>x.path.startsWith(path.join(f.root,f.payload))));assert.ok(r.excluded.some(x=>x.reason.includes('Setup mudou')));
});
test('arquivo adicionado ao índice durante a inspeção é protegido na conferência final',()=>{
 const f=fixture(),target=f.put(f.payload+'/new-tracked.bin');let added=false;
 const io={...fs,openSync(p,...args){if(p===target&&!added){added=true;f.git('add','--force',target);}return fs.openSync(p,...args);}};
 const r=f.inspect(io);assert.ok(added);assert.ok(!r.candidates.some(x=>x.path===target));assert.ok(r.protected.some(x=>x.path===target));
});
test('Windows: bloqueio real FileShare.None é tratado sem alteração', {skip:process.platform!=='win32'},async()=>{
 const f=fixture(),blocked=path.join(f.root,f.payload,'resources/app.asar');
 const child=spawn('powershell.exe',['-NoProfile','-NonInteractive','-Command',"$f=[IO.File]::Open($env:CLEANUP_LOCK_TEST,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::None);Write-Output 'READY';[Console]::ReadLine() | Out-Null;$f.Dispose()"],{env:{...process.env,CLEANUP_LOCK_TEST:blocked},windowsHide:true,stdio:['pipe','pipe','pipe']});
 const exited=new Promise((resolve,reject)=>{child.once('error',reject);child.once('exit',code=>code===0?resolve():reject(Error('lock helper exit '+code)));});
 try{await new Promise((resolve,reject)=>{child.stdout.on('data',b=>{if(b.toString().includes('READY'))resolve();});child.once('error',reject);child.once('exit',code=>reject(Error('helper exited before lock '+code)));});
  const r=f.inspect();assert.ok(!r.candidates.some(x=>x.path.startsWith(path.join(f.root,f.payload))));
 }finally{child.stdin.end('\n');await exited;}
 assert.equal(fs.readFileSync(blocked,'utf8'),'synthetic-asar');
});
