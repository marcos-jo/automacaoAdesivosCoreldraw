; All changes to the old installation happen in the isolated staged executable.
!ifndef BUILD_UNINSTALLER
  !define MUI_CUSTOMFUNCTION_ABORT fourpsUserAbort
  Var /GLOBAL fourpsStage
  Var /GLOBAL fourpsConfig
  Var /GLOBAL fourpsStarted
  Var /GLOBAL fourpsCommitted
  Var /GLOBAL fourpsSequence
!endif

!macro customCheckAppRunning
  ${nsProcess::FindProcess} "${APP_EXECUTABLE_FILENAME}" $R0
  StrCmp $R0 "603" fourps_app_closed
    MessageBox MB_OK|MB_ICONSTOP "Feche normalmente o ${PRODUCT_NAME} e seus auxiliares antes de instalar. Nenhum processo será encerrado automaticamente." /SD IDOK
    SetErrorLevel 2
    Abort
  fourps_app_closed:
!macroend

!macro fourpsCallUpgrade ACTION
  ; Do not interrupt the child between moving files and recording its journal.
  GetDlgItem $R5 $HWNDPARENT 2
  EnableWindow $R5 0
  IntOp $fourpsSequence $fourpsSequence + 1
  ClearErrors
  WriteINIStr $fourpsConfig "Update" "Request" "$fourpsSequence"
  ${if} ${Errors}
    StrCpy $R2 "2"
    StrCpy $R3 "Não foi possível preparar o resultado do componente de atualização. Preserve os registros para conferência."
  ${else}
  nsExec::ExecToStack '"$fourpsStage\${APP_EXECUTABLE_FILENAME}" --4ps-upgrade=${ACTION} --4ps-config="$fourpsConfig"'
  Pop $R2
  Pop $R3
  ; UTF-16 result file is written atomically before the staged GUI process exits.
  ReadINIStr $R4 "$PLUGINSDIR\fourps-result-${ACTION}-$fourpsSequence.ini" "Result" "Action"
  ReadINIStr $R6 "$PLUGINSDIR\fourps-result-${ACTION}-$fourpsSequence.ini" "Result" "Exit"
  ReadINIStr $R7 "$PLUGINSDIR\fourps-result-${ACTION}-$fourpsSequence.ini" "Result" "Message"
  ${if} $R4 != "${ACTION}"
  ${orIf} $R6 != $R2
  ${orIf} $R7 == ""
    ${if} $R2 == "0"
      StrCpy $R2 "2"
    ${endif}
    StrCpy $R3 "O componente de atualização não entregou um relatório válido. Ação ${ACTION}; saída $R2. Preserve os registros para conferência."
  ${else}
    StrCpy $R3 $R7
  ${endif}
  ${endif}
  DetailPrint "$R3"
  EnableWindow $R5 1
!macroend

!macro fourpsRollbackUpgrade
  Call fourpsRollback
!macroend

!macro fourpsDeclareFunctions
!ifndef BUILD_UNINSTALLER
Function fourpsRollback
  StrCmp $fourpsStarted "true" 0 fourps_rollback_end
  StrCmp $fourpsCommitted "true" fourps_rollback_end
    !insertmacro fourpsCallUpgrade rollback
    StrCmp $R2 "0" fourps_rollback_end
      MessageBox MB_OK|MB_ICONSTOP "A recuperação precisa de conferência. Preserve o diário .4ps-upgrade ao lado da instalação e procure suporte. Não repita o instalador." /SD IDOK
  fourps_rollback_end:
FunctionEnd
!endif
!macroend

!macro fourpsAbortUpgrade
  WriteINIStr "$INSTDIR\..\.4ps-upgrade-last-error.ini" "Installer" "Detail" "$R3"
  WriteINIStr "$INSTDIR\..\.4ps-upgrade-last-error.ini" "Installer" "Exit" "$R2"
  !insertmacro fourpsRollbackUpgrade
  MessageBox MB_OK|MB_ICONSTOP "A atualização não foi concluída. Confira o registro e o diário de recuperação antes de tentar novamente. Seu perfil não foi alterado." /SD IDOK
  SetErrorLevel 2
  Abort
!macroend

!macro fourpsPrepareUpgrade
  StrCmp $installMode "all" 0 fourps_user_install
    MessageBox MB_OK|MB_ICONSTOP "Esta migração experimental é somente para a instalação do usuário atual." /SD IDOK
    SetErrorLevel 2
    Abort
  fourps_user_install:
  StrCpy $fourpsStage "$PLUGINSDIR\fourps-stage"
  StrCpy $fourpsConfig "$PLUGINSDIR\fourps-upgrade.ini"
  FileOpen $R4 $fourpsConfig w
  FileWriteByte $R4 255
  FileWriteByte $R4 254
  FileClose $R4
  SetOutPath $fourpsStage
  DetailPrint "Preparando e verificando o pacote antes de alterar a versão anterior..."
  !insertmacro extractEmbeddedAppPackage
  File "/oname=$PLUGINSDIR\fourps-stage\${UNINSTALL_FILENAME}" "${UNINSTALLER_OUT_FILE}"
  !ifdef UNINSTALLER_ICON
    File "/oname=$PLUGINSDIR\fourps-stage\uninstallerIcon.ico" "${UNINSTALLER_ICON}"
  !endif
  !ifdef FOURPS_UPGRADE_LAB
    StrCpy $newDesktopLink "$INSTDIR\..\lab-desktop\${SHORTCUT_NAME}.lnk"
    StrCpy $newStartMenuLink "$INSTDIR\..\lab-programs\${SHORTCUT_NAME}.lnk"
    StrCpy $R4 "$INSTDIR\..\lab-cache\installer.exe"
  !else
    StrCpy $R4 "$LOCALAPPDATA\${APP_INSTALLER_STORE_FILE}"
  !endif
  WriteINIStr $fourpsConfig "Update" "AppId" "${APP_ID}"
  WriteINIStr $fourpsConfig "Update" "Guid" "${APP_GUID}"
  WriteINIStr $fourpsConfig "Update" "Target" "$INSTDIR"
  WriteINIStr $fourpsConfig "Update" "Stage" "$fourpsStage"
  WriteINIStr $fourpsConfig "Update" "Installer" "$EXEPATH"
  WriteINIStr $fourpsConfig "Update" "Cache" "$R4"
  WriteINIStr $fourpsConfig "Update" "DesktopLink" "$newDesktopLink"
  WriteINIStr $fourpsConfig "Update" "StartLink" "$newStartMenuLink"
  StrCpy $fourpsStarted "true"
  !insertmacro fourpsCallUpgrade prepare
  StrCmp $R2 "0" fourps_prepared
    !insertmacro fourpsAbortUpgrade
  fourps_prepared:
!macroend

!macro fourpsInstallUpgrade
  DetailPrint "Atualizando arquivos com proteção de recuperação..."
  !insertmacro fourpsCallUpgrade install
  StrCmp $R2 "0" fourps_installed
    !insertmacro fourpsAbortUpgrade
  fourps_installed:
  !ifdef FOURPS_UPGRADE_FAIL_AFTER_COPY
    !ifndef FOURPS_UPGRADE_LAB
      !error "Fault injection is only permitted with the isolated lab identity."
    !endif
    StrCpy $R2 "2"
    StrCpy $R3 "Falha simulada após cópia e validação no laboratório."
    !insertmacro fourpsAbortUpgrade
  !endif
!macroend

!macro fourpsCommitUpgrade
  !ifdef FOURPS_UPGRADE_FAIL_AFTER_REGISTRY
    !ifndef FOURPS_UPGRADE_LAB
      !error "Fault injection is only permitted with the isolated lab identity."
    !endif
    StrCpy $R2 "2"
    StrCpy $R3 "Falha simulada após registro no laboratório."
    !insertmacro fourpsAbortUpgrade
  !endif
  !insertmacro fourpsCallUpgrade commit
  StrCmp $R2 "0" fourps_committed
    !insertmacro fourpsAbortUpgrade
  fourps_committed:
  StrCpy $fourpsCommitted "true"
!macroend

!macro fourpsLabShortcuts
  CreateDirectory "$INSTDIR\..\lab-desktop"
  CreateDirectory "$INSTDIR\..\lab-programs"
  CreateShortCut "$newDesktopLink" "$appExe"
  CreateShortCut "$newStartMenuLink" "$appExe"
!macroend

!ifndef BUILD_UNINSTALLER
  Function .onInstFailed
    !insertmacro fourpsRollbackUpgrade
  FunctionEnd
  Function fourpsUserAbort
    !insertmacro fourpsRollbackUpgrade
  FunctionEnd
  Function .onGUIEnd
    !insertmacro fourpsRollbackUpgrade
  FunctionEnd
!endif
