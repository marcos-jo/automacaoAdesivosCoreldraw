// Somente harness: entrada separada, userData obrigatório e saída app.quit normal.
const fs=require('node:fs'),path=require('node:path'),{pathToFileURL}=require('node:url'),electron=require('electron');
const folder=process.env.FORMATS_TEST_FOLDER,source=process.env.FORMATS_TEST_SOURCE;
if(!folder||!source||!path.basename(folder).startsWith('4ps-size-manager-'))throw Error('Ambiente isolado obrigatório');
const profile=path.join(folder,'profile');electron.app.setPath('userData',profile);
let desktop,timer,requested=false;
const trace=(event,data={})=>fs.appendFileSync(path.join(folder,'electron-events.ndjson'),JSON.stringify({event,pid:process.pid,...data})+'\n');
for(const event of ['before-quit','will-quit','quit'])electron.app.on(event,()=>{trace(event);if(event==='will-quit')clearInterval(timer);});
const ipcMain=new Proxy(electron.ipcMain,{get(target,key){if(key==='handle')return(channel,handler)=>target.handle(channel,/print|updates:.*(download|install|check)/i.test(channel)?async()=>{throw Error('Operação bloqueada pelo teste');}:handler);const v=Reflect.get(target,key);return typeof v==='function'?v.bind(target):v;}});
import(pathToFileURL(path.join(source,'desktop/app.mjs')).href).then(async({startDesktop})=>{
    desktop=await startDesktop({...electron,ipcMain,dialog:{...electron.dialog,showMessageBoxSync:()=>1}},{profile});
    if(!desktop)throw Error('Aplicativo isolado não iniciou');
    fs.writeFileSync(path.join(folder,'electron-ready.json'),JSON.stringify({pid:process.pid,userData:electron.app.getPath('userData'),versions:process.versions,preferences:desktop.window.webContents.getLastWebPreferences()}));
    timer=setInterval(()=>{if(!requested&&fs.existsSync(path.join(folder,'close-request'))){requested=true;trace('app.quit');electron.session.defaultSession.flushStorageData();electron.app.quit();}},100);
}).catch(e=>{trace('error',{stack:e.stack});process.exitCode=1;});
