// Build de distribuição com atualização. Sem identidade/provedor/certificado,
// deve falhar, nunca produzir silenciosamente um atualizador sem assinatura.
const base=require('./electron-builder.config.cjs');
const url=new URL(process.env.PRINTSHOPS_UPDATE_URL||'https://invalid.invalid/');
const publisher=process.env.PRINTSHOPS_PUBLISHER_NAME;
if(!process.env.PRINTSHOPS_UPDATE_URL||url.protocol!=='https:'||url.username||url.password||url.search||url.hash||!publisher?.trim())throw new Error('Informe PRINTSHOPS_UPDATE_URL HTTPS e PRINTSHOPS_PUBLISHER_NAME da assinatura real.');
module.exports={...base,forceCodeSigning:true,
    win:{...base.win,verifyUpdateCodeSignature:true,signtoolOptions:{publisherName:publisher}},
    publish:{provider:'generic',url:url.href,channel:'experimental'},
    releaseInfo:{releaseNotes:'- Notificação opcional de atualizações e lista de novidades.\n- Download com progresso e validação de integridade/assinatura.\n- Verificação das abas antes da instalação e reinício.'}
};
