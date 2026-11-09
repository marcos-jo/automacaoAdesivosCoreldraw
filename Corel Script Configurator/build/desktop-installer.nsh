; Somente o compilador NSIS muda de pasta; nada muda no Windows de destino.
; Isso seleciona nosso adaptador extractAppPackage, preservando o template padrão.
!macro customHeader
  !insertmacro fourpsDeclareFunctions
  !insertmacro fourpsUnregisterCallback
  ; O template vem pelo stdin: preservar também os includes relativos da raiz,
  ; inclusive uninstaller.nsh, antes de selecionar o adaptador local.
  !addincludedir "${PROJECT_DIR}\node_modules\app-builder-lib\templates\nsis"
  !addincludedir "${BUILD_RESOURCES_DIR}\nsis-generated"
  !cd "${BUILD_RESOURCES_DIR}\nsis-generated"
!macroend

!macro verifyDistribution
  DetailPrint "Verificando integralmente o conversor privado..."
  nsExec::ExecToStack '"$INSTDIR\${APP_EXECUTABLE_FILENAME}" --verify-distribution'
  Pop $R2
  Pop $R3
  StrCmp $R2 "0" distribution_valid
    DetailPrint "$R3"
    MessageBox MB_OK|MB_ICONSTOP "A instalação não passou na verificação integral do conversor. Nenhuma conversão foi liberada. Preserve o perfil e procure suporte." /SD IDOK
    SetErrorLevel 2
    Abort
  distribution_valid:
!macroend

!include "${BUILD_RESOURCES_DIR}\pdf-association.nsh"
!include "${BUILD_RESOURCES_DIR}\nsis\upgrade-transaction.nsh"
!include "${BUILD_RESOURCES_DIR}\nsis\uninstall-owned.nsh"
