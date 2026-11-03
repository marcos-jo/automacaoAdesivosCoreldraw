// Identidade separada para ensaio de registro/abertura/desinstalação reais.
const base=require('./electron-builder.config.cjs');
module.exports={...base,appId:'br.com.4printshops.opentest',productName:'4PrintShops PDF Test',executableName:'4PrintShops PDF Test',
    directories:{output:'dist-desktop/pdf-open-test'},extraMetadata:{name:'4printshops-open-test',desktopOpenTest:true},
    nsis:{...base.nsis,shortcutName:'4PrintShops PDF Test',createDesktopShortcut:false,createStartMenuShortcut:false}};
