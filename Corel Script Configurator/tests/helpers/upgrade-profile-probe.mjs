// Probe the real ASAR before opening a window or accessing a real user profile.
import * as electron from 'electron';
import assert from 'node:assert/strict';
import path from 'node:path';
import {pathToFileURL} from 'node:url';
import {mkdtempSync} from 'node:fs';
electron.app.setPath('userData',mkdtempSync(path.join(electron.app.getPath('temp'),'4ps-profile-probe-')));
const root=path.resolve(process.argv[2]);
const stopped=Error('PROFILE_PROBE_STOP');
let product,profile;
try {
 const {startDesktop}=await import(pathToFileURL(path.join(root,'resources/app.asar/desktop/app.mjs')).href);
 const app={isPackaged:true,setName:name=>{product=name;},getPath:key=>key==='exe'?path.join(root,'4PrintShops Upgrade Lab.exe'):'C:\\PROFILE_MUST_NOT_BE_USED',setPath:(key,value)=>{assert.equal(key,'userData');profile=value;throw stopped;}};
 await assert.rejects(startDesktop({app}),error=>error===stopped);
 assert.equal(product,'4PrintShops Upgrade Lab');
 assert.equal(profile,path.resolve(root,'../lab-profile'));
 console.log(JSON.stringify({passed:true,product,profile,realProfileNeverOpened:true}));
 electron.app.exit(0);
}catch(error){console.error(error.stack);electron.app.exit(1);}
