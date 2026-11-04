// Impressão automatizada SOMENTE na Microsoft Print to PDF, com arquivo explícito.
// Epson/Konica são consultadas e têm documentos preparados/cancelados, sem StartDoc.
import assert from 'node:assert/strict';
import {mkdir,readFile,writeFile,unlink} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import path from 'node:path';
import {randomUUID} from 'node:crypto';
import {spawn,execFile} from 'node:child_process';
import {promisify} from 'node:util';
import {PDFDocument} from '../public/vendor/pdf-lib/pdf-lib.js';
import {createMixedPrintFixture,createPrintFormFixture,printScales} from './helpers/windows-print-fixtures.mjs';
import {PrintConfigClient} from '../desktop/print-config.mjs';
import {WindowsPrintCoordinator} from '../desktop/windows-print.mjs';
const exec=promisify(execFile),folder=path.resolve('.desktop-test/print12/native');await mkdir(folder,{recursive:true});
const ps=path.join(process.env.SystemRoot,'System32/WindowsPowerShell/v1.0/powershell.exe');
const powershell=async command=>(await exec(ps,['-NoProfile','-NonInteractive','-Command',command],{windowsHide:true,timeout:30000})).stdout.trim();
assert.equal(await powershell("(Get-ItemProperty 'HKCU:/Software/Microsoft/Windows NT/CurrentVersion/Windows').LegacyDefaultPrinterMode"),'1','Recusar ensaio virtual com gerência automática do padrão');
const defaultBefore=await powershell('(Get-CimInstance Win32_Printer | Where-Object Default).Name');
const client=new PrintConfigClient({helper:path.resolve('build/hold-native/4PrintShops.PrintConfig.exe'),manifest:path.resolve('build/hold-native/manifest.json'),experimental:true,windowsPrint:true});
const bytes=await createMixedPrintFixture(await readFile('build/hold-native/synthetic.pdf'));
await writeFile(path.join(folder,'source.pdf'),bytes);
const work=path.join(tmpdir(),'4PrintShops-print-jobs');await mkdir(work,{recursive:true});
const file=path.join(work,randomUUID().replaceAll('-','')+'.pdf');await writeFile(file,bytes);
const observations=[],watchers=[];
async function monitor(printer){
    const script=await readFile(new URL('./helpers/windows-print-defaults.ps1',import.meta.url),'utf8');
    const child=spawn(ps,['-NoProfile','-NonInteractive','-Command',`& {${script}} -PrinterName '${printer.replaceAll("'","''")}'`],{windowsHide:true,stdio:['pipe','pipe','pipe']});watchers.push(child);let text='';child.stdout.on('data',data=>text+=data);
    const wait=async match=>{for(let n=0;n<300;n++){if(text.includes(match))return;await new Promise(r=>setTimeout(r,50));}throw Error('Monitor de preferências sem resposta');};
    await wait('READY');return async()=>{text='';child.stdin.write('check\n');await wait('CHANGED');assert.equal(text.trim(),'UNCHANGED');};
}
try{
    const list=await client.request({action:'list'});
    const queues=process.env.PRINT_VIRTUAL_ONLY==='1'?['Microsoft Print to PDF']:['WF-C21000 Series(Rede)','KONICA MINOLTA C6100PS','Microsoft Print to PDF'];
    for(const printer of queues){
        assert.ok(list.some(item=>item.name===printer));
        const unchanged=await monitor(printer),state=await client.request({action:'select',printer});
        const base={session:state.session,revision:state.revision};
        const proof=await client.request({action:'prepare-document',...base,file,scale:120});
        assert.equal(proof.pages,3);assert.ok(Math.abs(proof.dimensions[1].width-300)<0.01);assert.equal(proof.appScale,120);
        await assert.rejects(client.request({action:'prepare-document',...base,revision:base.revision+1,file,scale:120}),/desatualizada/);
        await client.request({action:'cancel-document',...base});
        await assert.rejects(client.request({action:'print-document',...base,token:proof.token,name:'invalid consumed token'}),/desatualizada/);
        // O token cancelado deve falhar antes de CreateDC/StartDoc, inclusive nas filas físicas.
        observations.push({printer,driver:state.queue,geometry:state.geometry,preparedOnly:true,pages:proof.pages});await unchanged();
    }
    let state=await client.request({action:'select',printer:'Microsoft Print to PDF'});
    const coordinator=new WindowsPrintCoordinator(client,async()=>path.join(folder,'virtual-'+Date.now()+'.pdf'));
    const results=[];
    for(const scale of printScales){
        const output=path.join(folder,'virtual-'+scale+'.pdf');await unlink(output).catch(()=>{});
        coordinator.chooseOutput=async()=>output;
        const result=await coordinator.run({session:state.session,revision:state.revision,scale,name:'4PS VIRTUAL PRINT12 '+scale,operation:randomUUID().replaceAll('-','')},bytes);
        assert.equal(result.submitted,true);assert.equal(result.pages,3);
        const pdf=await PDFDocument.load(await readFile(output));assert.equal(pdf.getPageCount(),3);
        results.push({scale,pages:3,output,jobId:result.jobId});
    }
    coordinator.chooseOutput=async()=>path.join(folder,'form.pdf');await unlink(path.join(folder,'form.pdf')).catch(()=>{});
    const formResult=await coordinator.run({session:state.session,revision:state.revision,scale:100,name:'4PS VIRTUAL PRINT12 FORM',operation:randomUUID().replaceAll('-','')},await createPrintFormFixture());assert.equal(formResult.submitted,true);
    const bad=path.join(work,randomUUID().replaceAll('-','')+'.pdf');await writeFile(bad,'%PDF-1.7 invalid');
    try{await assert.rejects(client.request({action:'prepare-document',session:state.session,revision:state.revision,file:bad,scale:100}),/não conseguiu abrir|páginas/);}finally{await unlink(bad);}
    assert.equal(await powershell('(Get-CimInstance Win32_Printer | Where-Object Default).Name'),defaultBefore);
    await writeFile(path.join(folder,'result.json'),JSON.stringify({passed:true,virtualOnly:process.env.PRINT_VIRTUAL_ONLY==='1',physicalSendingTested:false,physicalQueuesPreparedOnly:true,defaultBefore,defaultPreserved:true,observations,virtual:results,formAppearanceGenerated:true,corruptRejected:true},null,2));
    console.log('PASSOU: filas consultadas '+queues.join(', ')+'. Três páginas mistas, escalas 100/120/87,5, formulário, PDF inválido, token cancelado e preferências. Saídas exclusivamente Microsoft Print to PDF.');
}finally{for(const child of watchers){child.stdin.end('quit\n');child.kill();}client.destroy();await unlink(file).catch(()=>{});}
