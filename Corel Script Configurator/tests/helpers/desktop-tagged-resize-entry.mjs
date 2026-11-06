// Somente testes; não incluído no pacote. Seleciona pasta isolada sem diálogo humano.
import * as electron from 'electron';
import path from 'node:path';
import {startDesktop} from '../../desktop/app.mjs';
const folder=process.env.PRINTSHOPS_TEST_FOLDER;
if(!folder||!path.isAbsolute(folder))throw Error('Pasta de teste absoluta obrigatória.');
startDesktop({...electron,dialog:{...electron.dialog,showOpenDialog:async()=>({canceled:false,filePaths:[folder]}),showMessageBoxSync:()=>1}}, {profile:path.join(folder,'profile')}).catch(error=>{console.error(error);electron.app.exit(1);});
