const base=require('./electron-builder.config.cjs');
module.exports={
    ...base,
    appId:'br.com.4printshops.holdtest',productName:'4PrintShops Hold Test',executableName:'4PrintShops Hold Test',
    directories:{output:'dist-desktop/hold-native-test'},
    extraMetadata:{name:'4printshops-hold-test',desktopHoldTest:true,desktopHoldNativeTest:true,desktopWindowsPrint:true,version:'1.0.0-experimental.14'},
    files:[...base.files,'build/hold-native/manifest.json'],
    beforePack:async()=>{
        await base.beforePack();
        await (await import('./scripts/build-print-lab.mjs')).buildPrintLab();
        await (await import('./scripts/build-hold-native.mjs')).buildHoldNative();
    },
    extraResources:[...base.extraResources,{from:'build/hold-native',to:'hold-native',filter:['**/*','!*.cs']}],
    win:{...base.win,artifactName:'4PrintShops-Hold-Test-${version}-Windows-${arch}-Setup.${ext}'},
    nsis:{...base.nsis,shortcutName:'4PrintShops Hold Test'}
};
