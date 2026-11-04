// Identidade de homologação separada: não substitui o produto nem laboratórios anteriores.
const base=require('./electron-builder.config.cjs');
module.exports={...base,appId:'br.com.4printshops.printconfigtest',productName:'4PrintShops Print Config Test',executableName:'4PrintShops Print Config Test',
    directories:{output:'dist-desktop/print-config-test'},
    extraMetadata:{name:'4printshops-print-config-test',desktopPrintConfigTest:true,version:'1.0.0-experimental.8'},
    win:{...base.win,artifactName:'4PrintShops-Print-Config-Test-${version}-Windows-${arch}-Setup.${ext}'},
    nsis:{...base.nsis,shortcutName:'4PrintShops Print Config Test'}};
