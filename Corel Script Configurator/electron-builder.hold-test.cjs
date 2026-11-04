const base=require('./electron-builder.config.cjs');
module.exports={
    ...base,
    appId:'br.com.4printshops.holdtest',
    productName:'4PrintShops Hold Test',
    executableName:'4PrintShops Hold Test',
    directories:{output:'dist-desktop/hold-test'},
    extraMetadata:{name:'4printshops-hold-test',desktopHoldTest:true,version:'1.0.0-experimental.9'},
    beforePack:async()=>{
        await base.beforePack();
        await (await import('./scripts/build-print-lab.mjs')).buildPrintLab();
    },
    extraResources:[...base.extraResources,{from:'build/print-lab',to:'print-lab',filter:['**/*']}],
    win:{...base.win,artifactName:'4PrintShops-Hold-Test-${version}-Windows-${arch}-Setup.${ext}'},
    nsis:{...base.nsis,shortcutName:'4PrintShops Hold Test'}
};
