// Somente leitura. Não há implementação de exclusão, movimento ou aplicação.
import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import {execFileSync} from 'node:child_process';
import {fileURLToPath} from 'node:url';

const slash=p=>p.split(path.sep).join('/');
const inside=(base,p)=>p===base||(!path.relative(base,p).startsWith('..'+path.sep)&&path.relative(base,p)!=='..'&&!path.isAbsolute(path.relative(base,p)));
const setup=v=>`dist-desktop/hold-native-test/4PrintShops-Hold-Test-1.0.0-experimental.${v}-Windows-x64-Setup.exe`;
const payloads=[['home24/installer-payload',24],['nup25/installer-payload',25],['nup26/installer-payload',26],['word30/setup-payload',30],['import31/setup-payload',31],['release33/payload',33]];
const extractions=[['word30/setup-extracted',30],['import31/setup-extracted',31],['release33/setup-extracted',33]];
// Pais exatos revisados na auditoria. Novos cenários exigem revisão desta lista.
const profileRuns=[
 'actualtext-audit/chrome-1794149408266','actualtext-audit/chrome-approved','actualtext-audit/isolated-before',
 'autoscroll/after-chrome','autoscroll/before-chrome','autoscroll/before-electron','autoscroll/chrome-approved',
 'autoscroll/chrome-final','autoscroll/chrome-lifecycle','autoscroll/chrome-lifecycle-final','autoscroll/chrome-lifecycle2',
 'autoscroll/chrome-native','autoscroll/chrome-native-final','autoscroll/electron-approved',
 'checkpoint26/policy-chrome','checkpoint26/policy-electron','home24/chrome-integrated','home24/electron-booklet',
 'nup25/chrome','nup25/chrome-integrated','nup25/electron','nup26/chrome','nup26/chrome-integrated',
 'nup26/electron','nup26/selection-chrome','release33/imposition-chrome','release33/thumbnails-electron'
];
const protectedName=/(?:backup|before|patch|manifest|report|relatorio|evidence|evidencia|proof|hash|sha256|results|screenshot|capture|licen[cs]e|notice|source-code)/i;
const protectedExtension=/\.(?:pdf|docx?|jpe?g|png|tiff?|svg|psd|cdr|ai|patch|diff|log|md|txt|json|ya?ml|csv|xml)$/i;
const profileCaches=new Set(['Cache','Code Cache','GPUCache','GPUPersistentCache','GrShaderCache','ShaderCache','DawnGraphiteCache','DawnWebGPUCache','optimization_guide_model_store','component_crx_cache','extensions_crx_cache']);
const signature=s=>`${s.size}:${s.mtimeMs}:${s.ctimeMs}:${s.ino}:${s.mode}`;

export function inspectStorage({appRoot=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..'),io=fs}={}){
 appRoot=path.resolve(appRoot);
 const git=args=>execFileSync('git',['--no-optional-locks','-C',appRoot,...args],{encoding:'utf8',stdio:['ignore','pipe','pipe']});
 const repo=path.resolve(git(['rev-parse','--show-toplevel']).trim());
 if(!inside(repo,appRoot))throw Error('Aplicação fora do repositório.');
 // Verificar também os ancestrais: realpath sozinho esconderia junctions.
 function resolveSafe(relative,base=appRoot){
  if(path.isAbsolute(relative)||relative.split(/[\\/]/).includes('..'))throw Error('Caminho absoluto/ascendente proibido.');
  const target=path.resolve(base,relative);if(!inside(repo,target))throw Error('Caminho externo proibido.');
  const parsed=path.parse(target);let cursor=parsed.root;
  for(const part of target.slice(parsed.root.length).split(path.sep).filter(Boolean)){
   cursor=path.join(cursor,part);const st=io.lstatSync(cursor);
   if(st.isSymbolicLink())throw Error('Link/junction proibido: '+cursor);
  }
  if(path.resolve(io.realpathSync(target))!==target)throw Error('Caminho real divergente.');
  return target;
 }
 resolveSafe('');
 if(JSON.parse(io.readFileSync(resolveSafe('package.json'),'utf8')).name!=='corel-script-configurator')throw Error('Identidade da aplicação inválida.');
 resolveSafe('AGENTS.md');
 const tracked=new Set(git(['ls-files','-z']).split('\0').filter(Boolean).map(p=>path.resolve(appRoot,p).toLowerCase()));
 const pending=git(['status','--porcelain=v1','-z','--untracked-files=all']);
 let head=null;try{head=git(['rev-parse','--verify','HEAD']).trim();}catch{/* Fixture/repositório ainda sem commit. */}
 const report={schema:1,mode:'simulation-only',createdAt:new Date().toISOString(),repo,appRoot,
  head,gitStatus:pending,totalBytes:0,totalComplete:true,
  candidates:[],protected:[],excluded:[],skippedLinks:[],groups:[],categories:{},recoverableBytes:0,
  limitations:['Snapshot, não autorização de exclusão. Mudanças posteriores invalidam o manifesto.',
   'Não detecta todos os processos ativos: perfil selecionável não significa perfil livre para exclusão.',
   'Clientes/evidências não podem ser identificados universalmente por conteúdo; arquivos desconhecidos são preservados fora da lista.',
   'Setup preservado recebe SHA256; sua existência/layout não substitui auditoria de correspondência do pacote.',
   'Somente caches explicitamente permitidos dos perfis são candidatos; preferências, histórico e armazenamento local permanecem protegidos.']};
 const files=new Map(),dirs=new Map(),issues=[];
 function inventory(base){
  const out=new Map(),dirList=new Map(),faults=[];
  function walk(p){
   let st;try{st=io.lstatSync(p);}catch(e){faults.push({path:p,reason:e.code||e.message});return;}
   if(st.isSymbolicLink()){faults.push({path:p,reason:'link/junction não seguido'});return;}
   if(st.isDirectory()){
    try{const names=io.readdirSync(p).sort();dirList.set(p,JSON.stringify(names));for(const name of names)walk(path.join(p,name));}
    catch(e){faults.push({path:p,reason:e.code||e.message});}
   }else if(st.isFile())out.set(p,{bytes:st.size,signature:signature(st),links:st.nlink});
   else faults.push({path:p,reason:'tipo especial não seguido'});
  }
  walk(base);return {out,dirList,faults};
 }
 const initial=inventory(repo);
 for(const [p,s]of initial.out){files.set(p,s);report.totalBytes+=s.bytes;}
 for(const [p,s]of initial.dirList)dirs.set(p,s);
 issues.push(...initial.faults);
 report.skippedLinks=issues.filter(e=>e.reason==='link/junction não seguido');
 report.totalComplete=issues.every(e=>e.reason==='link/junction não seguido');
 report.excluded.push(...issues);
 const rel=p=>slash(path.relative(appRoot,p));
 function protectedReason(p){
  if(tracked.has(p.toLowerCase()))return 'arquivo rastreado pelo Git';
  const parts=rel(p).split('/');
  if(parts.some(n=>protectedName.test(n)))return 'backup/manifesto/relatório/evidência ou aviso protegido';
  if(protectedExtension.test(p))return 'documento, imagem, configuração ou evidência preservada';
  if(files.get(p)?.links>1)return 'hardlink preservado';
  return null;
 }
 function hash(p){
  const fd=io.openSync(p,'r'),h=crypto.createHash('sha256'),buf=Buffer.alloc(1024*1024);
  try{let n;while((n=io.readSync(fd,buf,0,buf.length,null)))h.update(buf.subarray(0,n));}
  finally{io.closeSync(fd);}
  return h.digest('hex');
 }
 function regular(p){const s=io.lstatSync(p);if(!s.isFile())throw Error('Arquivo regular obrigatório.');return s;}
 function mz(p){regular(p);const fd=io.openSync(p,'r');try{const b=Buffer.alloc(2);if(io.readSync(fd,b,0,2,0)!==2||b.toString()!=='MZ')throw Error('Executável sem cabeçalho MZ.');}finally{io.closeSync(fd);}}
 const setupProof=new Map();
 function proof(version){
  if(setupProof.has(version))return setupProof.get(version);
  const p=resolveSafe(setup(version));mz(p);const before=signature(regular(p)),sha256=hash(p);
  if(before!==signature(regular(p)))throw Error('Setup mudou durante leitura.');
  const result={path:p,sha256,bytes:regular(p).size,stat:before};setupProof.set(version,result);return result;
 }
 function evaluate(relative,category,version){
  const group={path:path.resolve(appRoot,relative),category,reason:category==='profiles'?'Perfil sintético em cenário explicitamente autorizado':'Extração conhecida com Setup original preservado',bytes:0,status:'excluded'};
  report.groups.push(group);
  try{
   const root=resolveSafe(relative);if(!io.lstatSync(root).isDirectory())throw Error('Diretório obrigatório.');
   if(version)group.setup=proof(version);
   if(category==='payloads'){mz(resolveSafe(relative+'/4PrintShops Hold Test.exe'));regular(resolveSafe(relative+'/resources/app.asar'));}
   if(category==='extractions')regular(resolveSafe(relative+'/$PLUGINSDIR/app-64.7z'));
   if(category==='profiles'){
    const chrome=files.has(path.join(root,'Local State'))&&dirs.has(path.join(root,'Default'));
    const electron=files.has(path.join(root,'DevToolsActivePort'))&&dirs.has(path.join(root,'Cache'))&&dirs.has(path.join(root,'Local Storage'));
    if(!chrome&&!electron)throw Error('Estrutura de perfil sintético não reconhecida.');
   }
   if(issues.some(e=>inside(root,e.path)))throw Error('Link, erro ou objeto especial no candidato.');
   const selected=[];
   for(const [p,s]of files){if(!inside(root,p))continue;
    let reason=protectedReason(p);
    if(category==='profiles'){
     const parts=path.relative(root,p).split(path.sep);
     const cache=parts[0]==='Default'?parts[1]:parts[0];
     if(!profileCaches.has(cache))reason='Dados de perfil fora dos caches sintéticos explicitamente permitidos';
    }
    if(reason){report.protected.push({path:p,bytes:s.bytes,reason});continue;}
    resolveSafe(rel(p));if(signature(regular(p))!==s.signature)throw Error('Arquivo mudou antes da leitura: '+p);
    const sha256=hash(p);
    if(signature(regular(p))!==s.signature)throw Error('Arquivo mudou durante leitura: '+p);
    selected.push({path:p,bytes:s.bytes,sha256,stat:s.signature,category,group:root});
   }
   // Uma segunda fotografia recusa o grupo inteiro se houver criação, remoção ou alteração.
   const after=inventory(root),beforeFiles=[...files].filter(([p])=>inside(root,p));
   if(after.faults.length||after.out.size!==beforeFiles.length||beforeFiles.some(([p,s])=>after.out.get(p)?.signature!==s.signature)||
      [...dirs].filter(([p])=>inside(root,p)).some(([p,names])=>after.dirList.get(p)!==names))throw Error('Conteúdo mudou durante inspeção; grupo recusado.');
   report.candidates.push(...selected);group.bytes=selected.reduce((n,f)=>n+f.bytes,0);group.status=selected.length?'potential':'protected';
   if(!selected.length)group.reason='Nenhum arquivo elegível; conteúdo protegido integralmente.';
  }catch(e){group.reason=e.code==='ENOENT'?'Caminho ou pré-requisito inexistente':e.message;report.excluded.push({path:group.path,reason:group.reason});}
 }
 for(const [p,v]of payloads)evaluate('.desktop-test/'+p,'payloads',v);
 for(const [p,v]of extractions)evaluate('.desktop-test/'+p,'extractions',v);
 for(const run of profileRuns){
  const parent='.desktop-test/'+run;
  try{const p=resolveSafe(parent);for(const name of io.readdirSync(p).sort())if(/^profile(?:-\d+)?$/.test(name))evaluate(parent+'/'+name,'profiles');}
  catch(e){report.excluded.push({path:path.resolve(appRoot,parent),reason:e.code||e.message});}
 }
 // Conferir novamente os pré-requisitos: nunca manter candidatos cujo Setup mudou/sumiu.
 for(const g of report.groups.filter(g=>g.setup&&g.status==='potential')){
  try{resolveSafe(rel(g.setup.path));if(signature(regular(g.setup.path))!==g.setup.stat)throw Error('Setup mudou após inspeção.');}
  catch(e){report.candidates=report.candidates.filter(f=>f.group!==g.path);g.status='excluded';g.bytes=0;g.reason=e.message;report.excluded.push({path:g.path,reason:g.reason});}
 }
 const finalTracked=new Set(git(['ls-files','-z']).split('\0').filter(Boolean).map(p=>path.resolve(appRoot,p).toLowerCase()));
 report.candidates=report.candidates.filter(f=>{
  if(!finalTracked.has(f.path.toLowerCase()))return true;
  report.protected.push({path:f.path,bytes:f.bytes,reason:'Arquivo rastreado ao final da inspeção'});return false;
 });
 for(const g of report.groups.filter(g=>g.status==='potential'))g.bytes=report.candidates.filter(f=>f.group===g.path).reduce((n,f)=>n+f.bytes,0);
 // Não selecionar tudo o que está dentro de uma pasta: o manifesto contém somente arquivos individuais.
 for(const f of report.candidates){report.recoverableBytes+=f.bytes;const c=report.categories[f.category]??={files:0,bytes:0};c.files++;c.bytes+=f.bytes;}
 report.protectedRoots=['.git','AGENTS.md','node_modules','build/libreoffice','build/libreoffice-cache','public','desktop','src','scripts','tests','docs',setup(32),setup(33),'dist-desktop/hold-native-test/win-unpacked'].map(p=>path.resolve(p==='.git'?repo:appRoot,p));
 report.protectedPolicy='Tudo fora dos grupos explicitamente avaliados é protegido por padrão. Dentro deles, rastreados, documentos, imagens, backups, avisos e evidências permanecem protegidos.';
 return report;
}

export function main(args=process.argv.slice(2)){
 // Sem --apply, --delete, caminhos livres ou manifesto de entrada executável.
 if(args.some(a=>a!=='--json'))throw Error('Somente simulação: aceita apenas --json. Nenhuma exclusão disponível.');
 const r=inspectStorage();
 if(args.includes('--json'))console.log(JSON.stringify(r,null,2));
 else{
  const gb=n=>(n/1e9).toFixed(3)+' GB';
  console.log(`SIMULAÇÃO — nenhum arquivo alterado\nTotal: ${gb(r.totalBytes)}${r.totalComplete?'':' (parcial: erros de leitura)'}; links/junctions não contados: ${r.skippedLinks.length}\nPotencial: ${gb(r.recoverableBytes)}`);
  for(const g of r.groups)console.log(`${g.status} | ${gb(g.bytes)} | ${g.path} | ${g.reason}`);
  console.log('Categorias:',JSON.stringify(r.categories));
  console.log(`Protegidos dentro dos grupos: ${r.protected.length}. Excluídos/alertas: ${r.excluded.length}.`);
  console.log(r.protectedPolicy);for(const p of r.protectedRoots)console.log('PROTEGIDO | '+p);
  console.log('Manifesto detalhado em stdout: npm run --silent cleanup -- --json (redirecione para um arquivo NOVO se desejar guardar).');
 }
 return r;
}
if(process.argv[1]&&path.resolve(process.argv[1])===fileURLToPath(import.meta.url)){
 try{main();}catch(e){console.error('Inspeção recusada: '+e.message);process.exitCode=1;}
}
