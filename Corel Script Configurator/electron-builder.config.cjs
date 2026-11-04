module.exports = {
    appId: 'br.com.4printshops.desktop', productName: '4PrintShops', executableName: '4PrintShops',
    directories: { output: 'dist-desktop' }, asar: true, npmRebuild: false,
    beforePack:async()=>{await (await import('./scripts/build-print-helper.mjs')).buildPrintHelper();},
    electronFuses: { runAsNode: false, enableNodeOptionsEnvironmentVariable: false, enableNodeCliInspectArguments: false,
        onlyLoadAppFromAsar: true, enableEmbeddedAsarIntegrityValidation: true },
    files: [ 'desktop/**/*.mjs', 'desktop/preload.cjs', 'public/index.html', 'public/css/*.css', 'public/img/logo.png',
        'build/icons/app.png','build/native/manifest.json', 'public/js/main.js', 'public/js/app/{shell,tabs,menus,export-dialog,save-file,desktop-client,desktop-updates,desktop-open,desktop-assisted,desktop-print-ui,desktop-print-layout}.js', 'public/js/tools/*.js',
        '!public/js/tools/*backup*', 'public/vendor/**/*.js', 'package.json',
        '!node_modules/pdf-lib/**', '!node_modules/pdfjs-dist/**', '!node_modules/@pdf-lib/**', '!node_modules/@napi-rs/**',
        '!**/*.{pdf,pem,key,pfx,p12,log,tmp}', '!tests/**', '!docs/**', '!modelos/**', '!src/**' ],
    extraResources: [{from:'build/icons/pdf.ico',to:'icons/pdf.ico'},{from:'build/native/4PrintShops.PrintConfig.exe',to:'native/4PrintShops.PrintConfig.exe'}],
    win: { target: [ { target: 'nsis', arch: ['x64'] } ], icon: 'build/icons/app.ico',
        artifactName: '4PrintShops-${version}-Windows-${arch}-Setup.${ext}', requestedExecutionLevel: 'asInvoker' },
    nsis: { include:'build/pdf-association.nsh', installerIcon:'build/icons/app.ico',uninstallerIcon:'build/icons/app.ico',oneClick: false, perMachine: false, allowElevation: false, allowToChangeInstallationDirectory: true,
        createDesktopShortcut: true, createStartMenuShortcut: true, shortcutName: '4PrintShops',
        deleteAppDataOnUninstall: false, runAfterFinish: false, installerLanguages: ['pt_BR'], language: '1046' },
    // Sem publish/provider: nada é publicado ou consultado nesta versão experimental.
    publish: null
};
