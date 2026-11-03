// Laboratório isolado. Nunca usar este instalador como distribuição comercial.
const base=require('./electron-builder.config.cjs');
const version=process.env.PRINTSHOPS_LAB_VERSION;
if(!['1.0.0-experimental.2','1.0.0-experimental.3'].includes(version))throw new Error('Informe PRINTSHOPS_LAB_VERSION experimental.2 ou experimental.3.');
module.exports={...base,appId:'br.com.4printshops.updatelab',productName:'4PrintShops Update Lab',executableName:'4PrintShops Update Lab',
    directories:{output:'dist-desktop/update-lab/'+version},
    extraMetadata:{name:'4printshops-update-lab',version,desktopUpdateLab:true},
    win:{...base.win,artifactName:'4PrintShops-Update-Lab-${version}-Setup.${ext}',signtoolOptions:{publisherName:'4PrintShops Update Lab (TEST ONLY)'},verifyUpdateCodeSignature:true},
    nsis:{...base.nsis,shortcutName:'4PrintShops Update Lab',createDesktopShortcut:false,createStartMenuShortcut:false},
    publish:{provider:'generic',url:'http://127.0.0.1:18083/',channel:'experimental'},
    releaseInfo:{releaseNotes:version.endsWith('.3')?'- Notificação discreta de nova versão e novidades.\n- Download opcional com progresso.\n- Consulta de todas as abas antes de instalar e reiniciar.':'- Versão inicial do laboratório de atualizações.'}
};
