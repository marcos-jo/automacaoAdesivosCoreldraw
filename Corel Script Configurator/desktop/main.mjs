import * as electron from 'electron';
import {startDesktop} from './app.mjs';
import {validateDistribution} from './distribution-validation.mjs';
import {runUpgradeCommand} from './update-transaction.mjs';
const upgrade=process.argv.find(v=>v.startsWith('--4ps-upgrade='));
if(upgrade){
    await runUpgradeCommand({...electron,action:upgrade.slice('--4ps-upgrade='.length),configFile:process.argv.find(v=>v.startsWith('--4ps-config='))?.slice('--4ps-config='.length)});
}else if(process.argv.includes('--verify-distribution')){
    await validateDistribution(electron);
}else{
    startDesktop(electron).catch(error => {
        electron.dialog.showErrorBox('Não foi possível iniciar o 4PrintShops', error.message);
        electron.app.quit();
    });
}
