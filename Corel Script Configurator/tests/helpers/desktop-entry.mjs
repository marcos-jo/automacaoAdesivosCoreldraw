// Exclusivo dos testes; excluído do instalador. Não existe modo de teste no app distribuído.
import * as electron from 'electron';
import { readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { startDesktop } from '../../desktop/app.mjs';

const folder=process.env.PRINTSHOPS_TEST_FOLDER;
const control=path.join(folder,'command.json'), status=path.join(folder,'status.json');
let options={}, serial=0, dialogs=0;
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
                if(command.action==='close')desktop.window?.close();
                if(command.action==='minimize')desktop.window?.minimize();
                if(command.action==='maximize')desktop.window?.maximize();
                if(command.action==='restore')desktop.window?.restore();
                if(command.action==='size')desktop.window?.setSize(1024,650);
            }
            const w=desktop.window;
            await writeFile(status,JSON.stringify({serial,dialogs,state:desktop.documentState,busy:desktop.busy,
                preferences:w?.webContents.getLastWebPreferences(),maximized:w?.isMaximized(),minimized:w?.isMinimized(),pid:process.pid}));
        } catch {} finally {reading=false;}
    },50);
    electron.app.once('will-quit',()=>clearInterval(timer));
}
}).catch(error=>{console.error(error);electron.app.exit(1);});
