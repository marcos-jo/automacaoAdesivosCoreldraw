// Compile fault variants from the same verified archive/uninstaller as the lab.
// Instrument this build process only; never edit dependencies or production IDs.
import assert from 'node:assert/strict';
import {createRequire} from 'node:module';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
const root=fileURLToPath(new URL('../',import.meta.url)),require=createRequire(import.meta.url);
const builder=require('electron-builder'),{NsisTarget}=require('app-builder-lib/out/targets/nsis/NsisTarget.js');
assert.equal(require('app-builder-lib/package.json').version,'26.18.0');
const original=NsisTarget.prototype.executeMakensis;assert.equal(typeof original,'function');
let variants=0;
NsisTarget.prototype.executeMakensis=async function(defines,commands,script,options){
 assert.equal(defines.APP_ID,'br.com.4printshops.upgradelab.transaction36','Fault builds must never target a real installation');
 await original.call(this,defines,commands,script,options);
 if('BUILD_UNINSTALLER' in defines)return;
 assert.equal(defines.VERSION,'1.0.0-experimental.36');
 for(const [label,flag] of [['FailCopy','FOURPS_UPGRADE_FAIL_AFTER_COPY'],['FailRegistry','FOURPS_UPGRADE_FAIL_AFTER_REGISTRY']]){
  const output=path.join(root,'dist-desktop/upgrade-transaction-lab','4PrintShops-Upgrade-Lab-'+label+'-1.0.0-experimental.36-Setup.exe');
  await original.call(this,{...defines,[flag]:null},{...commands,OutFile:'"'+output+'"'},script,options);
  variants++;console.log('Isolated NSIS fault variant: '+output);
 }
};
try{
 process.env.PRINT_GHOSTSCRIPT='none';
 await builder.build({projectDir:root,config:path.join(root,'electron-builder.upgrade-lab.cjs'),targets:builder.Platform.WINDOWS.createTarget(['nsis'],builder.Arch.x64),publish:'never'});
 assert.equal(variants,2);
}finally{NsisTarget.prototype.executeMakensis=original;}
