const fs=require('node:fs');
const crypto=require('node:crypto');
module.exports=function loadTrust(product,channel){
    const file=process.env.FOURPRINT_PUBLIC_TRUST;
    if(!file)throw new Error('Informe FOURPRINT_PUBLIC_TRUST com a raiz pública aprovada.');
    const source=JSON.parse(fs.readFileSync(file,'utf8'));
    if(source.product!==product||source.channel!==channel||!Array.isArray(source.keys)||!source.keys.length||source.keys.length>4)throw new Error('Raiz de confiança incompatível com a identidade da instalação.');
    const keys=source.keys.map(({id,pem})=>{
        if(typeof pem!=='string'||!pem.startsWith('-----BEGIN PUBLIC KEY-----')||pem.includes('PRIVATE'))throw new Error('Somente chaves públicas são permitidas.');
        const key=crypto.createPublicKey(pem),digest=crypto.createHash('sha256').update(key.export({format:'der',type:'spki'})).digest('hex');
        if(key.asymmetricKeyType!=='ed25519'||digest!==id)throw new Error('Chave pública inválida.');return {id,pem};
    });
    if(new Set(keys.map(k=>k.id)).size!==keys.length)throw new Error('Chaves duplicadas.');
    return {product,channel,keys};
};
