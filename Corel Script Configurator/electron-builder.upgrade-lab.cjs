// Isolated upgrade tests: registry and shortcuts never use the real product.
const base=require('./electron-builder.hold-native-test.cjs');
module.exports={...base,
    appId:'br.com.4printshops.upgradelab.transaction36',
    productName:'4PrintShops Upgrade Lab',executableName:'4PrintShops Upgrade Lab',
    directories:{output:'dist-desktop/upgrade-transaction-lab'},
    extraMetadata:{...base.extraMetadata,name:'4printshops-upgrade-lab'},
    nsis:{...base.nsis,include:'build/nsis/upgrade-lab.nsh',shortcutName:'4PrintShops Upgrade Lab'},
    win:{...base.win,artifactName:'4PrintShops-Upgrade-Lab-${version}-Setup.${ext}'}
};
