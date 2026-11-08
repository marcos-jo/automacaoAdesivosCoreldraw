import {open, lstat, stat, mkdtemp, mkdir, writeFile, readFile, rm} from 'node:fs/promises';
import path from 'node:path';
import {tmpdir} from 'node:os';
import {pathToFileURL} from 'node:url';
import {spawn} from 'node:child_process';
import {createHash} from 'node:crypto';
import {inflateRawSync} from 'node:zlib';
import {PDFDocument, PDFDict, PDFName, PDFArray} from '../public/vendor/pdf-lib/pdf-lib.js';

export const WORD_LIMIT = 64 * 1024 * 1024;
export const PDF_LIMIT = 128 * 1024 * 1024;
const hash = bytes => createHash('sha256').update(bytes).digest('hex');
const canceled = () => new Error('Importação de Word cancelada.');
const xmlText = text => text.replace(/&(#x[0-9a-f]+|#\d+|quot|apos|lt|gt|amp);/gi, (_all,code)=>{
    if(code[0]==='#') {const value=code[1].toLowerCase()==='x'?parseInt(code.slice(2),16):Number(code.slice(1));return value>0&&value<=0x10ffff?String.fromCodePoint(value):'\uFFFD';}
    return {quot:'"',apos:"'",lt:'<',gt:'>',amp:'&'}[code]||_all;
});

// Somente caminhos escolhidos pelo diálogo nativo. Sem URLs, ADS ou dispositivos.
export function isWordPath(file) {
    if (typeof file !== 'string' || !path.isAbsolute(file) || /[\x00-\x1f]/.test(file)) return false;
    if (process.platform === 'win32' && (!/^[a-z]:\\/i.test(file) || /[:]/.test(file.slice(2)))) return false;
    return /\.(docx|doc)$/i.test(file);
}

async function boundedRead(file, limit) {
    const info = await lstat(file);
    if (!info.isFile() || info.isSymbolicLink() || info.size < 8 || info.size > limit) throw Error('Arquivo inválido ou maior que o limite permitido.');
    const handle = await open(file, 'r');
    try {
        const before = await handle.stat();
        if (!before.isFile() || before.size !== info.size || before.size > limit) throw Error('O arquivo mudou durante a leitura.');
        const bytes = Buffer.alloc(before.size + 1);
        let length = 0;
        while (length < bytes.length) {
            const result = await handle.read(bytes, length, bytes.length - length, length);
            if (!result.bytesRead) break;
            length += result.bytesRead;
        }
        const after = await handle.stat();
        if (length !== before.size || before.size !== after.size || before.mtimeMs !== after.mtimeMs) throw Error('O arquivo mudou durante a leitura.');
        return bytes.subarray(0, length);
    } finally { await handle.close(); }
}

// Leitura limitada do diretório ZIP; nunca extrai caminhos do documento.
// Apenas XMLs de metadados/fontes e relações são descompactados.
export function inspectWord(bytes, extension) {
    if (extension === '.doc') {
        if (!bytes.subarray(0,8).equals(Buffer.from('d0cf11e0a1b11ae1','hex')) || !bytes.includes(Buffer.from('WordDocument','utf16le'))) {
            throw Error('O arquivo .doc não é um documento Word binário reconhecido.');
        }
        return {sourcePages:null, requestedFonts:[], legacy:true};
    }
    if (bytes.readUInt32LE(0) !== 0x04034b50) throw Error('O arquivo .docx não é um pacote Word válido.');
    let end = -1;
    for (let i=bytes.length-22; i>=Math.max(0,bytes.length-65557); i--) {
        if (bytes.readUInt32LE(i)===0x06054b50 && i+22+bytes.readUInt16LE(i+20)===bytes.length) {end=i;break;}
    }
    if (end<0) throw Error('Pacote Word incompleto ou corrompido.');
    const count=bytes.readUInt16LE(end+10), size=bytes.readUInt32LE(end+12), offset=bytes.readUInt32LE(end+16);
    if (bytes.readUInt16LE(end+4) || bytes.readUInt16LE(end+6) || bytes.readUInt16LE(end+8)!==count || count>4096 || count===0 || offset+size!==end) throw Error('Pacote Word dividido, ZIP64 ou excessivamente complexo não suportado.');
    const entries=new Map(); let cursor=offset,total=0;
    for(let n=0;n<count;n++) {
        if(cursor+46>end || bytes.readUInt32LE(cursor)!==0x02014b50) throw Error('Diretório do Word inválido.');
        const flags=bytes.readUInt16LE(cursor+8),method=bytes.readUInt16LE(cursor+10),compressed=bytes.readUInt32LE(cursor+20),plain=bytes.readUInt32LE(cursor+24);
        const len=bytes.readUInt16LE(cursor+28),extra=bytes.readUInt16LE(cursor+30),comment=bytes.readUInt16LE(cursor+32),local=bytes.readUInt32LE(cursor+42);
        const next=cursor+46+len+extra+comment;
        if(next>end || flags&1 || ![0,8].includes(method) || local+30>offset) throw Error('Pacote Word criptografado ou inválido.');
        const name=bytes.subarray(cursor+46,cursor+46+len).toString('utf8');
        if(entries.has(name) || /(^\/|\\|(^|\/)\.\.(\/|$))/.test(name)) throw Error('Caminho interno inválido no Word.');
        total+=plain;
        if(total>256*1024*1024 || /vbaproject|activex|embeddings\//i.test(name)) throw Error('Word com macros, objetos ativos ou tamanho descompactado excessivo não suportado nesta etapa.');
        entries.set(name,{flags,method,compressed,plain,local}); cursor=next;
    }
    if(cursor!==end || !entries.has('word/document.xml') || !entries.has('[Content_Types].xml')) throw Error('O pacote não contém um documento Word.');
    function xml(name) {
        const e=entries.get(name); if(!e)return '';
        const p=e.local;
        if(bytes.readUInt32LE(p)!==0x04034b50 || bytes.readUInt16LE(p+8)!==e.method || bytes.readUInt16LE(p+6)!==e.flags) throw Error('Entrada interna do Word inválida.');
        const nameLength=bytes.readUInt16LE(p+26),start=p+30+nameLength+bytes.readUInt16LE(p+28);
        if(bytes.subarray(p+30,p+30+nameLength).toString('utf8')!==name || start+e.compressed>offset || e.plain>8*1024*1024) throw Error('XML do Word inválido ou excessivo.');
        const compressed=bytes.subarray(start,start+e.compressed);
        const result=e.method===0?compressed:inflateRawSync(compressed,{maxOutputLength:8*1024*1024});
        if(result.length!==e.plain) throw Error('XML do Word incompleto.');
        const text=result.toString('utf8');
        if(/<!DOCTYPE|<!ENTITY/i.test(text) || text.includes('\0')) throw Error('XML do Word não suportado.');
        return text;
    }
    for(const name of entries.keys()) if(/\.rels$/i.test(name) && /TargetMode\s*=\s*["']External["']/i.test(xmlText(xml(name)))) {
        throw Error('Este Word contém vínculos externos. Exporte um PDF pelo Word ou remova os vínculos em uma cópia antes de importar.');
    }
    if(/macroEnabled/i.test(xml('[Content_Types].xml'))) throw Error('Pacote Word com macros não suportado.');
    const app=xml('docProps/app.xml'),pages=Number(app.match(/<(?:\w+:)?Pages\b[^>]*>(\d+)<\//)?.[1]);
    const fonts=xml('word/fontTable.xml');
    const requestedFonts=[...new Set([...fonts.matchAll(/<(?:\w+:)?font\s[^>]*?(?:\w+:)?name\s*=\s*["']([^"']+)["']/g)].map(m=>xmlText(m[1]).slice(0,256)))].slice(0,256);
    return {sourcePages:pages>0&&pages<=2000?pages:null, requestedFonts, legacy:false};
}

export async function findLibreOffice(env=process.env) {
    const candidates=env.PRINTSHOPS_LIBREOFFICE?[env.PRINTSHOPS_LIBREOFFICE]:process.platform==='win32'
        ?[env.ProgramFiles,env['ProgramFiles(x86)']].filter(Boolean).map(p=>path.join(p,'LibreOffice','program','soffice.com'))
        :['/usr/bin/libreoffice','/usr/bin/soffice','/Applications/LibreOffice.app/Contents/MacOS/soffice'];
    for(const candidate of candidates) {
        if(!path.isAbsolute(candidate) || (process.platform==='win32'&&!/soffice\.(com|exe)$/i.test(path.basename(candidate))))continue;
        try{if((await stat(candidate)).isFile())return candidate;}catch{}
    }
    throw Error('LibreOffice não encontrado. Instale-o pelo site oficial e reabra o aplicativo. Uma instalação em outra pasta pode ser indicada por PRINTSHOPS_LIBREOFFICE (caminho completo do soffice.com).');
}

export const safeProfile = `<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry">
<item oor:path="/org.openoffice.Office.Common/Security/Scripting">
<prop oor:name="MacroSecurityLevel" oor:finalized="true"><value>3</value></prop>
<prop oor:name="DisableMacrosExecution" oor:finalized="true"><value>true</value></prop>
<prop oor:name="DisableActiveContent" oor:finalized="true"><value>true</value></prop>
<prop oor:name="BlockUntrustedRefererLinks" oor:finalized="true"><value>true</value></prop>
<prop oor:name="SecureURL" oor:finalized="true"><value/></prop>
</item>
<item oor:path="/org.openoffice.Office.Writer/Content/Update">
<prop oor:name="Link" oor:finalized="true"><value>2</value></prop>
<prop oor:name="Field" oor:finalized="true"><value>false</value></prop>
<prop oor:name="Chart" oor:finalized="true"><value>false</value></prop>
</item></oor:items>`;

// Não usa shell e termina somente a árvore do processo criado por esta operação.
export function runOffice(executable,args,{signal,timeoutMs=120000,cwd}={}) {
    return new Promise((resolve,reject)=>{
        if(signal?.aborted)return reject(canceled());
        const child=spawn(executable,args,{cwd,windowsHide:true,shell:false,stdio:['ignore','ignore','ignore']});
        let failure,timer;
        const stop=error=>{
            if(failure)return;failure=error;
            if(process.platform==='win32'&&child.pid) {
                const killer=spawn(path.join(process.env.SystemRoot||'C:\\Windows','System32','taskkill.exe'),['/PID',String(child.pid),'/T','/F'],{windowsHide:true,shell:false,stdio:'ignore'});
                killer.on('error',()=>child.kill());
            }else child.kill('SIGKILL');
        };
        const abort=()=>stop(canceled());
        signal?.addEventListener('abort',abort,{once:true});
        timer=setTimeout(()=>stop(Error('A conversão excedeu 120 segundos. O Word pode estar protegido, danificado ou ser complexo demais.')),timeoutMs);
        const done=()=>{clearTimeout(timer);signal?.removeEventListener('abort',abort);};
        child.once('error',()=>{done();reject(Error('Não foi possível iniciar o LibreOffice. Confira sua instalação e permissões.'));});
        child.once('close',code=>{done();if(failure)reject(failure);else if(code!==0)reject(Error('LibreOffice não concluiu a conversão. Verifique se o Word está protegido por senha ou danificado.'));else resolve();});
    });
}

export async function inspectPDF(bytes) {
    if(bytes.length>PDF_LIMIT || bytes.subarray(0,5).toString()!=='%PDF-')throw Error('O LibreOffice não gerou um PDF válido dentro do limite de 128 MiB.');
    let pdf,pages;try{pdf=await PDFDocument.load(bytes,{updateMetadata:false});pages=pdf.getPages();}catch{throw Error('O PDF convertido está inválido, incompleto ou protegido.');}
    if(!pages.length||pages.length>2000)throw Error('O PDF convertido deve ter entre 1 e 2.000 páginas.');
    const dimensions=pages.map(p=>{const {width,height}=p.getSize();if(!Number.isFinite(width)||!Number.isFinite(height)||width<=0||height<=0)throw Error('O PDF convertido possui dimensões inválidas.');return {widthMm:width*25.4/72,heightMm:height*25.4/72};});
    const fonts=new Map(),name=key=>PDFName.of(key);
    for(const [,object]of pdf.context.enumerateIndirectObjects()) {
        if(!(object instanceof PDFDict)||object.get(name('Type'))?.toString()!=='/Font')continue;
        const base=object.get(name('BaseFont'));if(!(base instanceof PDFName))continue;
        let dict=object;
        if(object.get(name('Subtype'))?.toString()==='/Type0') {
            const descendants=object.lookup(name('DescendantFonts'));
            if(descendants instanceof PDFArray)dict=descendants.lookup(0);
        }
        const descriptor=dict instanceof PDFDict?dict.lookup(name('FontDescriptor')):null;
        const embedded=descriptor instanceof PDFDict&&['FontFile','FontFile2','FontFile3'].some(k=>!!descriptor.get(name(k)));
        const font=base.decodeText().replace(/^[A-Z]{6}\+/,'');fonts.set(font,(fonts.get(font)??true)&&embedded);
    }
    return {pages:pages.length,dimensions,fonts:[...fonts].map(([name,embedded])=>({name,embedded}))};
}

export async function convertWord(file,{signal,executable,runner=runOffice,tempRoot=tmpdir(),timeoutMs}={}) {
    if(!isWordPath(file))throw Error('Selecione um arquivo Word .docx ou .doc em um disco local.');
    let folder;
    try {
        if(signal?.aborted)throw canceled();
        const original=await boundedRead(file,WORD_LIMIT);
        const source=inspectWord(original,path.extname(file).toLowerCase());
        executable??=await findLibreOffice();
        if(signal?.aborted)throw canceled();
        folder=await mkdtemp(path.join(tempRoot,'4printshops-word-'));
        const profile=path.join(folder,'profile'),out=path.join(folder,'out'),input=path.join(folder,'documento'+path.extname(file).toLowerCase());
        await mkdir(path.join(profile,'user'),{recursive:true});await mkdir(out);
        await writeFile(path.join(profile,'user','registrymodifications.xcu'),safeProfile,{flag:'wx'});
        await writeFile(input,original,{flag:'wx'});
        const options=Object.fromEntries(Object.entries({UseLosslessCompression:true,ReduceImageResolution:false,IsSkipEmptyPages:false,ExportFormFields:false,IsAddStream:false}).map(([k,v])=>[k,{type:'boolean',value:String(v)}]));
        const args=['-env:UserInstallation='+pathToFileURL(profile).href,'--headless','--nologo','--nodefault','--nofirststartwizard','--norestore','--infilter='+(source.legacy?'MS Word 97':'Office Open XML Text'),'--convert-to','pdf:writer_pdf_Export:'+JSON.stringify(options),'--outdir',out,input];
        await runner(executable,args,{signal,timeoutMs,cwd:folder});
        if(signal?.aborted)throw canceled();
        let bytes;try{bytes=await boundedRead(path.join(out,'documento.pdf'),PDF_LIMIT);}catch{throw Error('Nenhum PDF válido foi produzido. O Word pode estar protegido por senha, danificado ou não ser compatível.');}
        const report=await inspectPDF(bytes);
        if(hash(await boundedRead(file,WORD_LIMIT))!==hash(original))throw Error('O Word original mudou durante a conversão. Importe novamente.');
        if(signal?.aborted)throw canceled();
        return {name:path.basename(file).replace(/\.(docx|doc)$/i,'')+'_WORD.pdf',bytes:new Uint8Array(bytes),report:{...source,...report}};
    } catch(error) {
        if(error.code==='ENOENT'||error.code==='EACCES'||error.code==='EPERM')throw Error('Não foi possível ler o Word. Confira se o arquivo existe e se você tem permissão.');
        throw error;
    } finally {
        if(folder)try{await rm(folder,{recursive:true,force:true,maxRetries:4,retryDelay:200});}catch{throw Error('Não foi possível limpar a pasta temporária da importação: '+folder);}
    }
}

export async function importWord(window,dialog,options={}) {
    const choice=await dialog.showOpenDialog(window,{title:'Importar Word (conversão local)',filters:[{name:'Documentos Word',extensions:['docx','doc']}],properties:['openFile']});
    if(choice.canceled)return null;
    if(choice.filePaths?.length!==1)throw Error('Selecione somente um Word.');
    return convertWord(choice.filePaths[0],options);
}

export {boundedRead as readImportFile};
