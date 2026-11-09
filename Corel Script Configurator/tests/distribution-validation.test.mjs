import test from'node:test';import assert from'node:assert/strict';import{readFile}from'node:fs/promises';
import {validateDistribution} from'../desktop/distribution-validation.mjs';
function mock(packaged=true){const calls=[];return{calls,app:{isPackaged:packaged,setPath:(...args)=>calls.push(args),exit:code=>calls.push(['exit',code])}};}
test('validação interna verifica somente o pacote e usa perfil temporário antes de verificar',async()=>{
    const {app,calls}=mock();let path,probed=false;await validateDistribution({app,resourcesPath:'package',trustFile:'trusted.json',temporary:()=> 'temporary',verify:async(root,trust)=>{assert.deepEqual(calls,[['userData','temporary']]);path=[root,trust];return{files:19410,version:'test'};},probe:async(result,temporary)=>{assert.equal(result.files,19410);assert.equal(temporary,'temporary');probed=true;},log:()=>{},error:()=>{}});assert.match(path[0],/[/\\]lo$/);assert.equal(path[1],'trusted.json');assert.ok(probed);assert.deepEqual(calls.at(-1),['exit',0]);
});
test('runtime íntegro que não inicializa também impede sucesso da instalação',async()=>{
    const{app,calls}=mock();await validateDistribution({app,resourcesPath:'package',temporary:()=> 'temporary',verify:async()=>({files:19410}),probe:async()=>{throw Error('startup failed');},log:()=>assert.fail('Não pode aprovar'),error:()=>{}});assert.deepEqual(calls.at(-1),['exit',2]);
});
test('ausência, alteração e timeout recusam sucesso e não abrem o perfil do operador',async()=>{
    for(const mode of ['missing','changed','timeout']){const{app,calls}=mock();await validateDistribution({app,resourcesPath:'package',temporary:()=> 'temporary',verify:()=>mode==='timeout'?new Promise(()=>{}):Promise.reject(Error(mode)),timeout:5,log:()=>assert.fail('Não pode aprovar'),error:()=>{}});assert.deepEqual(calls,[['userData','temporary'],['exit',2]]);}
    const{app,calls}=mock(false);await validateDistribution({app,temporary:()=>assert.fail(),verify:()=>assert.fail(),error:()=>{}});assert.deepEqual(calls,[['exit',2]]);
});
test('adaptador mantém template original, proíbe cópia legada e falha fechada',async()=>{
    const header=await readFile('build/desktop-installer.nsh','utf8');assert.match(header,/!addincludedir.*templates\\nsis"/);assert.ok(header.indexOf('!addincludedir')<header.indexOf('!cd'));
    const nsis=await readFile('build/nsis/extractAppPackage.nsh','utf8'),entry=await readFile('build/desktop-installer.nsh','utf8');assert.match(nsis,/!include.*node_modules.*extractAppPackage/);assert.match(nsis,/!macroundef extractUsing7za/);assert.match(nsis,/robocopy\.exe/);assert.doesNotMatch(nsis,/^\s*CopyFiles\b/m);assert.deepEqual(nsis.split(/\r?\n/).filter(line=>/^\s*RMDir\b/.test(line)),['  RMDir /r "$PLUGINSDIR\\4ps-extract"']);assert.doesNotMatch(nsis,/ \/(?:256|MIR|PURGE)(?: |')/);assert.match(nsis,/IntCmp \$R2 8/);assert.match(nsis,/SetErrorLevel 2/);assert.match(entry,/--verify-distribution/);assert.match(entry,/SetErrorLevel 2/);assert.match(entry,/pdf-association\.nsh/);
});
