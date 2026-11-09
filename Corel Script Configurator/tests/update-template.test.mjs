import test from 'node:test';import assert from 'node:assert/strict';
import {mkdtemp,mkdir,copyFile,readFile} from 'node:fs/promises';import path from 'node:path';import {tmpdir} from 'node:os';
import {prepareNsisUpgrade} from '../scripts/prepare-nsis-upgrade.mjs';
test('pinned NSIS section uses transaction instead of the old uninstaller',async()=>{
 const root=await mkdtemp(path.join(tmpdir(),'4ps-template-test-'));for(const f of ['installSection.nsh','include/installer.nsh','include/installUtil.nsh','uninstaller.nsh']){const dest=path.join(root,'node_modules/app-builder-lib/templates/nsis',f);await mkdir(path.dirname(dest),{recursive:true});await copyFile(path.join('node_modules/app-builder-lib/templates/nsis',f),dest);}
 const identity=await prepareNsisUpgrade({packager:{projectDir:root,appInfo:{id:'br.com.4printshops.holdtest',productFilename:'4PrintShops Hold Test',version:'1.0.0-experimental.36',updaterCacheDirName:'4printshops-hold-test-updater'},config:{nsis:{shortcutName:'4PrintShops Hold Test'}}}});
 assert.equal(identity.guid,'90ff2ff0-f013-571c-b109-7edb265dacc8');
 const section=await readFile(path.join(root,'build/nsis-generated/installSection.nsh'),'utf8');assert.doesNotMatch(section,/!insertmacro (?:uninstallOldVersion|handleUninstallResult|installApplicationFiles)/);assert.match(section,/!insertmacro fourpsPrepareUpgrade/);assert.match(section,/!insertmacro fourpsInstallUpgrade/);assert.ok(section.indexOf('fourpsInstallUpgrade')<section.indexOf('registryAddInstallInfo'));assert.match(section,/fourpsAbortUpgrade/);
});
