// Exclusivo dos testes; excluído do instalador. Não existe modo de teste no app distribuído.
import * as electron from 'electron';
import { readFile, writeFile, rename } from 'node:fs/promises';
import path from 'node:path';
import { startDesktop } from '../../desktop/app.mjs';

const folder=process.env.PRINTSHOPS_TEST_FOLDER;
const control=path.join(folder,'command.json'), status=path.join(folder,'status.json');
let options={}, serial=0, dialogs=0, workState;
startDesktop({...electron,dialog:{...electron.dialog,
    showSaveDialog:async()=>{dialogs++;return options.save || {canceled:true};},
    showMessageBoxSync:()=>{dialogs++;return options.answer ?? 0;}
}}, {profile:path.join(folder,'profile')}).then(desktop => {
if (desktop) {
    let reading=false;
    const timer=setInterval(async()=>{
        if(reading)return;reading=true;
        try {
            const command=JSON.parse(await readFile(control,'utf8'));if(command.serial<serial)return;options=command;
            if(command.serial>serial){serial=command.serial;
                if(command.action==='probe-work'){workState=await new Promise(resolve=>{const token='smoke-'+Date.now(),timer=setTimeout(()=>resolve({pending:true,busy:true,timeout:true}),3000);const receive=(event,id,state)=>{if(id!==token||event.sender!==desktop.window.webContents)return;clearTimeout(timer);electron.ipcMain.removeListener('updates:work-response',receive);resolve(state);};electron.ipcMain.on('updates:work-response',receive);desktop.window.webContents.send('updates:work-request',token);});}
                if(command.action==='close')desktop.window?.close();
                if(command.action==='minimize')desktop.window?.minimize();
                if(command.action==='maximize')desktop.window?.maximize();
                if(command.action==='restore')desktop.window?.restore();
                if(command.action==='size')desktop.window?.setSize(1024,650);
            }
            const w=desktop.window;
            await writeFile(status+'.tmp',JSON.stringify({serial,dialogs,workState,state:desktop.documentState,busy:desktop.busy,
                preferences:w?.webContents.getLastWebPreferences(),maximized:w?.isMaximized(),minimized:w?.isMinimized(),pid:process.pid}));
            await rename(status+'.tmp',status); // O teste nunca deve ler JSON parcialmente escrito.
        } catch {} finally {reading=false;}
    },50);
    electron.app.once('will-quit',()=>clearInterval(timer));
}
}).catch(error=>{console.error(error);electron.app.exit(1);});
