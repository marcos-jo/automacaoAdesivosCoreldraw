import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import {createRequire} from 'node:module';
import {fileURLToPath} from 'node:url';

const root=fileURLToPath(new URL('../',import.meta.url));
const require=createRequire(new URL('../package.json',import.meta.url));
// Sem instalação: nos ensaios locais, apontar DESKTOP_TEST_MODULES ao preparo isolado.
const builderRequire=process.env.DESKTOP_TEST_MODULES
    ? createRequire(path.join(process.env.DESKTOP_TEST_MODULES,'../package.json')) : require;
const {FileMatcher}=builderRequire('app-builder-lib/out/fileMatcher');
const configs=[require('./electron-builder.config.cjs'),require('./electron-builder.hold-native-test.cjs')];
function seleciona(config,relative){
    const filter=new FileMatcher(root,path.join(root,'saida-nao-criada'),s=>s,config.files).createFilter();
    return filter(path.join(root,relative),{isDirectory:()=>false});
}

test('pacote: módulo legítimo de backup incluído nas configurações base e Hold Test',()=>{
    for(const config of configs)assert.equal(seleciona(config,'public/js/tools/pdf-format-backup.js'),true);
});
test('pacote: exceção não libera legado, outros backups, testes ou dados privados',()=>{
    for(const config of configs)for(const file of [
        'public/js/tools/pdfs backup.js','public/js/tools/outro-backup.js',
        'public/js/tools/pdf-format-backup-old.js','public/js/tools/pdf-format-backup.js.bak',
        'public/js/tools/subpasta/pdf-format-backup.js',
        'tests/package-format-backup.test.mjs','docs/ESTADO_PROJETO.md',
        '.desktop-test/casos-reais/uniao/CASO-030.pdf','profiles/user/Preferences',
        'public/vendor/exemplo.pdf','public/js/tools/credencial.key'
    ])assert.equal(seleciona(config,file),false,file);
});
test('pacote: grafo local do catálogo não contém dependências de runtime excluídas',()=>{
    const visited=new Set();
    function conferir(relative){
        if(visited.has(relative))return;visited.add(relative);
        assert.ok(fs.statSync(path.join(root,relative)).isFile());
        for(const config of configs)assert.equal(seleciona(config,relative),true,relative);
        const source=fs.readFileSync(path.join(root,relative),'utf8');
        for(const [,specifier]of source.matchAll(/\bfrom\s*['"]([^'"]+)['"]/g)){
            assert.ok(specifier.startsWith('.'),'Dependência não local: '+specifier);
            conferir(path.posix.normalize(path.posix.join(path.posix.dirname(relative),specifier)));
        }
    }
    conferir('public/js/tools/pdf-format-catalog.js');
    assert.deepEqual([...visited].sort(),['pdf-format-backup.js','pdf-format-catalog.js','pdf-format-storage.js','pdf-formats.js'].map(f=>'public/js/tools/'+f).sort());
});
test('pacote: identidade Hold Test e configurações de segurança permanecem preservadas',()=>{
    const hold=configs[1];assert.equal(hold.appId,'br.com.4printshops.holdtest');assert.equal(hold.productName,'4PrintShops Hold Test');
    for(const config of configs){
        assert.equal(config.asar,true);assert.equal(config.npmRebuild,false);assert.equal(config.publish,null);
        assert.equal(config.nsis.deleteAppDataOnUninstall,false);assert.equal(config.nsis.allowElevation,false);assert.equal(config.nsis.runAfterFinish,false);
        assert.deepEqual(config.electronFuses,{runAsNode:false,enableNodeOptionsEnvironmentVariable:false,enableNodeCliInspectArguments:false,onlyLoadAppFromAsar:true,enableEmbeddedAsarIntegrityValidation:true});
    }
});
