// Validação visual isolada: não substitui a instalação/perfil nem os labs anteriores.
const base=require('./electron-builder.config.cjs');
module.exports={...base,appId:'br.com.4printshops.icontest',productName:'4PrintShops Icon Test',executableName:'4PrintShops Icon Test',
    directories:{output:'dist-desktop/icon-test'},
    extraMetadata:{name:'4printshops-icon-test',desktopIconTest:true,version:'1.0.0-experimental.7'},
    win:{...base.win,artifactName:'4PrintShops-Icon-Test-${version}-Windows-${arch}-Setup.${ext}'},
    nsis:{...base.nsis,shortcutName:'4PrintShops Icon Test',createDesktopShortcut:true,createStartMenuShortcut:true}};
