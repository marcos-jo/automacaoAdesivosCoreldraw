import * as electron from 'electron';
import { startDesktop } from './app.mjs';

startDesktop(electron).catch(error => {
    electron.dialog.showErrorBox('Não foi possível iniciar o 4PrintShops', error.message);
    electron.app.quit();
});
