const {readFileSync}=require('node:fs');
const {assertGuest}=require('./scripts/authenticode-lab/guest-guard.cjs');
assertGuest(process.env.PRINTSHOPS_VM_MANIFEST,process.env.PRINTSHOPS_DISPOSABLE_GUEST==='YES');
if(process.env.ELECTRON_BUILDER_OFFLINE!=='true')throw new Error('Laboratório deve assinar sem TSA externo: ELECTRON_BUILDER_OFFLINE=true, com caches já preparados.');
const base=require('./electron-builder.update-lab.cjs');
const kind=process.env.PRINTSHOPS_LAB_SIGNER||'primary';
if(!['primary','foreign'].includes(kind))throw new Error('Assinante de laboratório inválido.');
const identities=JSON.parse(readFileSync(process.env.PRINTSHOPS_LAB_IDENTITIES,'utf8'));
const identity=identities.find(item=>item.id===kind);
const expected=kind==='primary'?'4PrintShops Update Lab (TEST ONLY)':'4PrintShops Other Publisher (TEST ONLY)';
if(identity?.publisher!==expected||!/^[A-Fa-f0-9]{40}$/.test(identity?.thumbprint||''))throw new Error('Identidade de teste inválida.');
module.exports={...base,forceCodeSigning:true,
    directories:{output:`dist-desktop/authenticode-lab/${kind}/${process.env.PRINTSHOPS_LAB_VERSION}`},
    win:{...base.win,verifyUpdateCodeSignature:true,signtoolOptions:{publisherName:expected,certificateSha1:identity.thumbprint,signingHashAlgorithms:['sha256']}}
};
