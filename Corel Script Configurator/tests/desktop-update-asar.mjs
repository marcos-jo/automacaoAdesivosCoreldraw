// Real Electron regression: physical installation I/O must bypass virtual ASAR.
import * as electron from 'electron';
import {createRequire} from 'node:module';import path from 'node:path';import assert from 'node:assert/strict';
import {randomUUID} from 'node:crypto';import {readFile as virtualReadFile} from 'node:fs/promises';
import {UpdateTransaction,readAsarJSON} from '../desktop/update-transaction.mjs';import {registryScope} from '../desktop/update-registry.mjs';
const require=createRequire(import.meta.url),raw=require('original-fs'),fs=raw.promises;
const lab=raw.mkdtempSync(path.join(electron.app.getPath('temp'),'4ps-asar-regression-'));electron.app.setPath('userData',path.join(lab,'profile'));
try{
 const source=path.resolve(process.argv[2]||'dist-desktop/upgrade-transaction-lab/win-unpacked/resources/app.asar');
 const pkg=await readAsarJSON(source,'package.json');assert.equal(pkg.version,'1.0.0-experimental.36');
 await assert.rejects(virtualReadFile(source),error=>['EISDIR','ENOENT'].includes(error.code));assert.equal((await fs.stat(source)).isFile(),true);
 const id={appId:'br.com.4printshops.upgradelab.'+randomUUID(),guid:randomUUID(),product:'4PS ASAR Lab',exe:'4PS ASAR Lab.exe',version:pkg.version,lab:true};
 const stage=path.join(lab,'stage'),target=path.join(lab,id.product);await fs.mkdir(path.join(stage,'resources'),{recursive:true});await fs.copyFile(source,path.join(stage,'resources/app.asar'));await fs.writeFile(path.join(stage,id.exe),'INERT LAB MARKER');
 const scope=registryScope(id),snapshot={trees:scope.allowedTrees.map(p=>({path:p,exists:false})),values:scope.allowedValues.map(v=>({...v,exists:false})),environment:{}};
 const tx=new UpdateTransaction({identity:id,stage,target,configFile:path.join(lab,'fourps-upgrade.ini'),verify:async()=>{},registry:async(_id,r)=>r.op==='busy'?{processes:[]}:r.op==='snapshot'?snapshot:{restored:true},desktopLink:path.join(lab,'lab-desktop',id.product+'.lnk'),startLink:path.join(lab,'lab-programs',id.product+'.lnk')});
 await tx.prepare();await tx.install();assert.deepEqual(await fs.readFile(path.join(target,'resources/app.asar')),await fs.readFile(source));await tx.rollback();
 console.log(JSON.stringify({passed:true,realElectron:true,virtualAsarFailureReproduced:true,physicalReadAndMigration:true,rollback:true,lab}));electron.app.exit(0);
}catch(e){console.error(e);electron.app.exit(1);}
