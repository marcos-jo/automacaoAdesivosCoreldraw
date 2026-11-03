import test from 'node:test';
import assert from 'node:assert/strict';
import {readFileSync} from 'node:fs';
import {createHash} from 'node:crypto';
import {createRequire} from 'node:module';
const require=createRequire(import.meta.url);

test('artes originais preservadas por hash e PNGs finais RGBA de 1024 pixels',()=>{
    for(const [source,hash,name]of [
        ['icone4ps.png','2001def4aae94153a077f9548be1676185390f8c83142c243212dbb8cfd6eeca','app'],
        ['iconepdf.png','5fd6f58cc1a567faa0d5fa5e2ca7c35ac04b74ad66b77e7a5ad0c4e5168c07cc','pdf']]){
        assert.equal(createHash('sha256').update(readFileSync('build/icons/sources/'+source)).digest('hex'),hash);
        const png=readFileSync('build/icons/'+name+'.png');assert.equal(png.subarray(0,8).toString('hex'),'89504e470d0a1a0a');
        assert.equal(png.readUInt32BE(16),1024);assert.equal(png.readUInt32BE(20),1024);assert.equal(png[24],8);assert.equal(png[25],6,'RGBA sem fundo adicionado');
    }
});
test('ICOs têm sete resoluções reais com alpha e máscara de transparência coerentes',()=>{
    for(const name of ['app','pdf']){
        const ico=readFileSync('build/icons/'+name+'.ico'),sizes=[16,24,32,48,64,128,256];
        assert.equal(ico.readUInt16LE(0),0);assert.equal(ico.readUInt16LE(2),1);assert.equal(ico.readUInt16LE(4),7);
        for(let n=0;n<sizes.length;n++){
            const entry=6+n*16,size=sizes[n],bytes=ico.readUInt32LE(entry+8),offset=ico.readUInt32LE(entry+12),frame=ico.subarray(offset,offset+bytes);
            assert.equal(ico[entry]||256,size);assert.equal(ico[entry+1]||256,size);assert.equal(frame.length,bytes);assert.equal(ico.readUInt16LE(entry+6),32);
            if(size===256){assert.equal(frame.subarray(0,8).toString('hex'),'89504e470d0a1a0a');assert.equal(frame.readUInt32BE(16),256);assert.equal(frame[25],6);continue;}
            assert.equal(frame.readUInt32LE(0),40);assert.equal(frame.readUInt32LE(4),size);assert.equal(frame.readUInt32LE(8),size*2);
            let visible=0,transparent=0;const stride=Math.ceil(size/32)*4,mask=40+size*size*4;
            for(let y=0;y<size;y++)for(let x=0;x<size;x++){
                const alpha=frame[40+(y*size+x)*4+3],masked=!!(frame[mask+y*stride+(x>>3)]&(0x80>>(x%8)));
                assert.equal(masked,alpha===0);if(alpha===0)transparent++;else visible++;
            }
            assert.ok(visible>0&&transparent>0,'Conteúdo e cantos transparentes em '+size+' px');
        }
    }
});
test('instalador e ProgID usam ícones distintos sem escrita no leitor padrão',()=>{
    const config=require('../electron-builder.config.cjs'),lab=require('../electron-builder.icon-test.cjs');
    assert.equal(config.win.icon,'build/icons/app.ico');assert.equal(config.nsis.installerIcon,config.win.icon);assert.equal(config.nsis.uninstallerIcon,config.win.icon);
    assert.ok(config.extraResources.some(r=>r.from==='build/icons/pdf.ico'&&r.to==='icons/pdf.ico'));
    assert.notEqual(lab.appId,config.appId);assert.equal(lab.extraMetadata.desktopIconTest,true);
    const nsis=readFileSync('build/pdf-association.nsh','utf8');assert.match(nsis,/DefaultIcon[^\n]*resources\\icons\\pdf\.ico/);
    assert.doesNotMatch(nsis,/WriteReg\w+[^\n]*UserChoice/);
    assert.doesNotMatch(nsis,/WriteReg\w+[^\n]*"Software\\Classes\\\.pdf"\s+""/);
    assert.match(nsis,/--open-pdf/);
});
