import {createHash,createPublicKey,verify} from 'node:crypto';
import {createRequire} from 'node:module';

// Reutiliza a implementação SemVer que acompanha o updater instalado.
const require=createRequire(import.meta.url),semver=createRequire(require.resolve('electron-updater/package.json'))('semver');
export const MAIN_PRODUCT='br.com.4printshops.desktop';
export const LAB_PRODUCT='br.com.4printshops.assistedlab';
export const MAX_INSTALLER=512*1024*1024;
export function keyId(pem){return createHash('sha256').update(createPublicKey(pem).export({type:'spki',format:'der'})).digest('hex');}
export function validatePolicy(policy){
    if(!policy||![MAIN_PRODUCT,LAB_PRODUCT].includes(policy.product)||!/^internal(?:-lab)?$/.test(policy.channel)||!Array.isArray(policy.keys)||policy.keys.length<1||policy.keys.length>4)throw new Error('Confiança de publicação não configurada no aplicativo.');
    for(const key of policy.keys){if(typeof key.pem!=='string'||!key.pem.startsWith('-----BEGIN PUBLIC KEY-----')||key.pem.includes('PRIVATE KEY')||createPublicKey(key.pem).asymmetricKeyType!=='ed25519'||key.id!==keyId(key.pem))throw new Error('Chave pública incorporada inválida.');}
    if(new Set(policy.keys.map(k=>k.id)).size!==policy.keys.length)throw new Error('Chaves públicas duplicadas.');
    return policy;
}
function decode(value){if(typeof value!=='string'||!value.length||!/^[A-Za-z0-9_-]+$/.test(value))throw new Error('Envelope de assinatura inválido.');const bytes=Buffer.from(value,'base64url');if(bytes.toString('base64url')!==value)throw new Error('Codificação de assinatura inválida.');return bytes;}
export function authenticateEnvelope(bytes,policy){
    validatePolicy(policy);if(bytes.length>65536)throw new Error('Manifesto excede o limite permitido.');
    const envelope=JSON.parse(Buffer.from(bytes).toString('utf8'));
    const key=policy.keys.find(k=>k.id===envelope.keyId);if(!key)throw new Error('Chave de publicação não reconhecida.');
    const payload=decode(envelope.payload),signature=decode(envelope.signature);
    if(signature.length!==64||!verify(null,payload,key.pem,signature))throw new Error('Assinatura do manifesto inválida.');
    const data=JSON.parse(payload.toString('utf8')),artifact=data.installer;
    if(data.schema!==1||data.product!==policy.product||data.channel!==policy.channel||data.platform!=='win32'||data.arch!=='x64')throw new Error('Publicação de produto, plataforma ou canal incompatível.');
    if(!semver.valid(data.version)||!Number.isSafeInteger(data.sequence)||data.sequence<1)throw new Error('Versão ou sequência de publicação inválida.');
    if(!artifact||!Number.isSafeInteger(artifact.size)||artifact.size<8||artifact.size>MAX_INSTALLER||!/^[a-f0-9]{128}$/.test(artifact.sha512||''))throw new Error('Integridade declarada do instalador inválida.');
    if(typeof artifact.path!=='string'||!new RegExp('^releases/'+data.version.replace(/[.*+?^${}()|[\]\\]/g,'\\$&')+'/[A-Za-z0-9_-][A-Za-z0-9._-]{0,170}\\.exe$').test(artifact.path)||artifact.path.includes('..'))throw new Error('Caminho de instalador não autorizado.');
    if(!Array.isArray(data.notes)||data.notes.length>30||data.notes.some(n=>typeof n!=='string'||n.length>1000))throw new Error('Notas de versão inválidas.');
    const issued=Date.parse(data.issuedAt),expires=Date.parse(data.expiresAt);
    if(!Number.isFinite(issued)||!Number.isFinite(expires)||expires<=issued||expires-issued>31*86400000)throw new Error('Validade da publicação inválida.');
    return {data,digest:createHash('sha256').update(payload).digest('hex')};
}
export function validateRelease(release,{currentVersion,floor={sequence:0},now=Date.now()}){
    const {data,digest}=release;
    if(Date.parse(data.issuedAt)>now+300000||Date.parse(data.expiresAt)<=now)throw new Error('Publicação expirada ou relógio incorreto.');
    if(data.sequence<floor.sequence||data.sequence===floor.sequence&&floor.digest!==digest)throw new Error('Publicação antiga ou sequência reutilizada.');
    if(!semver.valid(currentVersion)||semver.lt(data.version,currentVersion))throw new Error('Versão antiga: downgrade recusado.');
    return semver.gt(data.version,currentVersion);
}
export function sourceURL(value,{allowHttp=false}={}){
    const url=new URL(value);
    if((url.protocol!=='https:'&&!(url.protocol==='http:'&&allowHttp))||url.username||url.password||url.search||url.hash||!url.pathname.endsWith('/'))throw new Error('Origem deve usar HTTPS ou HTTP explicitamente autorizado para manifestos assinados, sem credenciais e terminando em /.');
    return url;
}
