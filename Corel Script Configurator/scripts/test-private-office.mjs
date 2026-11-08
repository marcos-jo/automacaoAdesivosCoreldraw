import {spawn} from 'node:child_process';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
const root=fileURLToPath(new URL('../',import.meta.url));
const env={...process.env};delete env.ELECTRON_RUN_AS_NODE;
const executable=path.join(root,'node_modules/electron/dist/electron.exe');
const child=spawn(executable,[path.join(root,'tests/desktop-private-office-package.mjs'),process.argv[2]||'dist-desktop/hold-native-test/win-unpacked'],{cwd:root,env,windowsHide:true,stdio:'inherit'});
process.exitCode=await new Promise((resolve,reject)=>{child.once('error',reject);child.once('close',code=>resolve(code??1));});
