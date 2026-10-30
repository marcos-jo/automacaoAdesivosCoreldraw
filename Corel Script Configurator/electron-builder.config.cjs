module.exports = {
    appId: 'br.com.4printshops.desktop', productName: '4PrintShops', executableName: '4PrintShops',
    directories: { output: 'dist-desktop' }, asar: true, npmRebuild: false,
    electronFuses: { runAsNode: false, enableNodeOptionsEnvironmentVariable: false, enableNodeCliInspectArguments: false,
        onlyLoadAppFromAsar: true, enableEmbeddedAsarIntegrityValidation: true },
    files: [ 'desktop/**/*.mjs', 'desktop/preload.cjs', 'public/index.html', 'public/css/*.css', 'public/img/logo.png',
        'public/js/main.js', 'public/js/app/{shell,tabs,menus,export-dialog,save-file,desktop-client}.js', 'public/js/tools/*.js',
        '!public/js/tools/*backup*', 'public/vendor/**/*.js', 'package.json',
        '!node_modules/pdf-lib/**', '!node_modules/pdfjs-dist/**', '!node_modules/@pdf-lib/**', '!node_modules/@napi-rs/**',
        '!**/*.{pdf,pem,key,pfx,p12,log,tmp}', '!tests/**', '!docs/**', '!modelos/**', '!src/**' ],
    win: { target: [ { target: 'nsis', arch: ['x64'] } ], icon: 'public/img/logo.png',
        artifactName: '4PrintShops-${version}-Windows-${arch}-Setup.${ext}', requestedExecutionLevel: 'asInvoker' },
    nsis: { oneClick: false, perMachine: false, allowElevation: false, allowToChangeInstallationDirectory: true,
        createDesktopShortcut: true, createStartMenuShortcut: true, shortcutName: '4PrintShops',
        deleteAppDataOnUninstall: false, runAfterFinish: false, installerLanguages: ['pt_BR'], language: '1046' },
    // Sem publish/provider: nada é publicado ou consultado nesta versão experimental.
    publish: null
};
