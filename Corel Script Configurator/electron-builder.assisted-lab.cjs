if(process.env.FOURPRINT_ASSISTED_LAB!=='1')throw new Error('Ative FOURPRINT_ASSISTED_LAB=1 somente para os ensaios isolados.');
const base=require('./electron-builder.config.cjs');
const loadTrust=require('./build/assisted-trust.cjs');
const version=process.env.FOURPRINT_LAB_VERSION;
if(!['1.0.0-experimental.5','1.0.0-experimental.6'].includes(version))throw new Error('Versão de laboratório não autorizada.');
module.exports={...base,appId:'br.com.4printshops.assistedlab',productName:'4PrintShops Assisted Lab',executableName:'4PrintShops Assisted Lab',
    directories:{output:'dist-desktop/assisted-lab/'+version},
    extraMetadata:{assistedUpdates:loadTrust('br.com.4printshops.assistedlab','internal-lab'),name:'4printshops-assisted-lab',desktopAssistedLab:true,version},
    win:{...base.win,artifactName:'4PrintShops-Assisted-Lab-${version}-Windows-${arch}-Setup.${ext}'},
    nsis:{...base.nsis,include:null,shortcutName:'4PrintShops Assisted Lab',createDesktopShortcut:false,createStartMenuShortcut:false}};
