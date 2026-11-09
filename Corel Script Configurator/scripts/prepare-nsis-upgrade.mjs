import {readFile,writeFile,mkdir} from 'node:fs/promises';
import {createHash} from 'node:crypto';
import {createRequire} from 'node:module';
const require=createRequire(import.meta.url),{UUID}=require('builder-util-runtime');
const pins={
    'installSection.nsh':'74c4dbb4163cdd90ca4850b75d304e916235c882d6af7b5cdad63bffbc977072',
    'include/installer.nsh':'0e319437dd01dcbf911f3f48f664fde0cefbaef704f1cdb1739f63d563f5d4a0',
    'include/installUtil.nsh':'97bd546b5cd2aaf16b77bc9e2be8a18962dd74ab5c4d23b35b163ca89bf4dd2a',
    'uninstaller.nsh':'9ee2dac4593478083e8aa6f8487287ce9401006ccd50ecc538871d133ea4a42c'
};
export async function prepareNsisUpgrade(context){
    const root=context.packager.projectDir,info=context.packager.appInfo;
    const id={schema:1,appId:info.id,guid:context.packager.config.nsis?.guid||UUID.v5(info.id,UUID.parse('50e065bc-3134-11e6-9bab-38c9862bdaf3')),product:info.productFilename,exe:info.productFilename+'.exe',version:info.version,shortcut:context.packager.config.nsis?.shortcutName||info.productFilename,cacheDir:info.updaterCacheDirName,lab:info.id.startsWith('br.com.4printshops.upgradelab.')};
    await mkdir(root+'/build/distribution',{recursive:true});await writeFile(root+'/build/distribution/identity.json',JSON.stringify(id,null,2));
    const templates={};for(const[f,pin]of Object.entries(pins)){const text=(await readFile(root+'/node_modules/app-builder-lib/templates/nsis/'+f,'utf8')).replaceAll('\r\n','\n');if(createHash('sha256').update(text).digest('hex')!==pin)throw Error('Revalidar transação NSIS: '+f);templates[f]=text;}
    let section=templates['installSection.nsh'];
    const from='!insertmacro uninstallOldVersion SHELL_CONTEXT\n!insertmacro handleUninstallResult SHELL_CONTEXT\n\n${if} $installMode == "all"\n  !insertmacro uninstallOldVersion HKEY_CURRENT_USER\n  !insertmacro handleUninstallResult HKEY_CURRENT_USER\n${endIf}';
    if(!section.includes(from))throw Error('Hook de migração não encontrado.');
    section=section.replace(from,'!insertmacro fourpsPrepareUpgrade');
    section=section.replace('!ifdef UNINSTALLER_ICON\n  File /oname=uninstallerIcon.ico "${UNINSTALLER_ICON}"\n!endif','; Uninstaller and icon are already staged and included in the transaction.');
    section=section.replace('!insertmacro installApplicationFiles','!insertmacro fourpsInstallUpgrade');
    section=section.replace('!insertmacro addStartMenuLink $keepShortcuts\n!insertmacro addDesktopLink $keepShortcuts','!ifdef FOURPS_UPGRADE_LAB\n  !insertmacro fourpsLabShortcuts\n!else\n  !insertmacro addStartMenuLink $keepShortcuts\n  !insertmacro addDesktopLink $keepShortcuts\n!endif');
    section=section.replace('!insertmacro registryAddInstallInfo','ClearErrors\n!insertmacro registryAddInstallInfo\nIfErrors fourps_metadata_failed fourps_metadata_ok\nfourps_metadata_failed:\n  !insertmacro fourpsAbortUpgrade\nfourps_metadata_ok:');
    await mkdir(root+'/build/nsis-generated',{recursive:true});await writeFile(root+'/build/nsis-generated/installSection.nsh',section);
    await writeFile(root+'/build/nsis-generated/extractAppPackage.nsh','!include "${BUILD_RESOURCES_DIR}\\nsis\\extractAppPackage.nsh"\n');
    // Omit only unused legacy caller functions; keep compiler warnings as errors.
    const util=templates['include/installUtil.nsh'].replace(/^Function (?:GetInQuotes|GetFileParent|handleUninstallResult|uninstallOldVersion)\n[\s\S]*?^FunctionEnd\n/gm,'');
    await writeFile(root+'/build/nsis-generated/installUtil.nsh',util);
    return id;
}
