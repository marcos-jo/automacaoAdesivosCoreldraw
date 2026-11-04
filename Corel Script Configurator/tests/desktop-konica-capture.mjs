// Captura somente na fila FILE: autorizada. O programa nativo não aceita destino físico.
import assert from 'node:assert/strict';
import {spawn,execFile} from 'node:child_process';
import {promisify} from 'node:util';
import {readFile,mkdir,writeFile} from 'node:fs/promises';
import {randomUUID} from 'node:crypto';
import path from 'node:path';
import {setTimeout as pause} from 'node:timers/promises';

const exec=promisify(execFile);
const powershell=path.join(process.env.SystemRoot,'System32/WindowsPowerShell/v1.0/powershell.exe');
const source=await readFile(new URL('./helpers/windows-print-properties.ps1',import.meta.url),'utf8');
const folder=path.resolve('.desktop-test/hold-native');
await mkdir(folder,{recursive:true});
const action=process.argv[2]||'hold';
const scale=process.argv[3]||'100';
if(!['100','120'].includes(scale))throw new Error('Escala de ensaio inválida.');
if(!['hold','print','holdprint'].includes(action))throw new Error('Método de ensaio inválido.');
const output=path.join(folder,action+'-'+randomUUID()+'.ps');
const child=spawn(path.resolve('build/hold-capture/4PrintShops.KonicaCapture.exe'),[
    path.resolve('build/hold-capture/synthetic.pdf'),output,scale
],{windowsHide:false,stdio:['ignore','pipe','pipe']});
let stdout='',stderr='';
child.stdout.on('data',chunk=>stdout+=chunk);child.stderr.on('data',chunk=>stderr+=chunk);
const completed=new Promise((resolve,reject)=>{child.once('error',reject);child.once('exit',code=>resolve(code));});
const native=async command=>exec(powershell,['-NoProfile','-NonInteractive','-Command',`& {${source}} -DriverProcess ${child.pid} -Action '${command}'`],{windowsHide:true,timeout:30000});
try {
    const deadline=Date.now()+30000;
    while(!stdout.includes('PROPERTIES_READY')){if(child.exitCode!==null||Date.now()>deadline)throw new Error('Diálogo nativo indisponível. '+stderr);await pause(100);}
    if(scale==='120')await native('a3-sizes');
    await native('workflow');await pause(250);await native(action);await native('accept');
    const code=await Promise.race([completed,pause(60000).then(()=>{throw new Error('Captura sem resposta. Não reenviar automaticamente.');})]);
    assert.equal(code,0,stderr);
    const report=JSON.parse(stdout.trim().split('\n').at(-1));
    assert.equal(report.captureQueue,'4PrintShops Konica PS Capture Lab');
    const expected={hold:'Hold',print:'Print',holdprint:'HoldPrint'}[action];
    assert.equal(report.hold.outputMethod,expected);
    const ps=await readFile(output);
    assert.ok(ps.length>1000);
    await writeFile(path.join(folder,action+'-latest.json'),JSON.stringify(report,null,2)+'\n');
    console.log(JSON.stringify({output,bytes:ps.length,method:report.hold.outputMethod,scale:report.scale,paper:[report.contextWidthMM,report.contextHeightMM]},null,2));
} finally {if(child.exitCode===null)child.kill();}
