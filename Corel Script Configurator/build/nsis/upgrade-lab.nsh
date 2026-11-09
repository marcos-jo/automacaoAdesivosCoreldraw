; All registration is redirected to a disposable namespace, never Windows Open With.
!define INSTALL_REGISTRY_KEY "Software\4PrintShops\UpgradeLab\${APP_GUID}\Install"
!define UNINSTALL_REGISTRY_KEY "Software\4PrintShops\UpgradeLab\${APP_GUID}\Uninstall"
!define FOURPS_CLASSES "Software\4PrintShops\UpgradeLab\${APP_GUID}\Classes"
!define FOURPS_REGISTERED "Software\4PrintShops\UpgradeLab\${APP_GUID}\RegisteredApplications"
!define PDF_CAPABILITIES "Software\4PrintShops\UpgradeLab\${APP_GUID}\Capabilities"
!define FOURPS_UPGRADE_LAB
!include "${BUILD_RESOURCES_DIR}\desktop-installer.nsh"
