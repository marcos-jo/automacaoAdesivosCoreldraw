import {spawn} from 'node:child_process';
import electron from 'electron';
import {fileURLToPath} from 'node:url';
const env={...process.env}; delete env.ELECTRON_RUN_AS_NODE;
const child=spawn(electron,['.',...process.argv.slice(2)],{cwd:fileURLToPath(new URL('../',import.meta.url)),env,stdio:'inherit'});
child.on('error',error=>{console.error(error.message);process.exitCode=1;});
child.on('exit',code=>{process.exitCode=code??1;});
for(const signal of ['SIGINT','SIGTERM'])process.on(signal,()=>child.kill(signal));
