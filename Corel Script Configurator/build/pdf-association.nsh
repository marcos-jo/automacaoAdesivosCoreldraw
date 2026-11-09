; Registra capacidade/Open With sem tocar no default de .pdf ou em UserChoice.
!define PDF_PROGID "${APP_ID}.PDF"
!define /ifndef FOURPS_CLASSES "Software\Classes"
!define /ifndef FOURPS_REGISTERED "Software\RegisteredApplications"
!define /ifndef PDF_CAPABILITIES "Software\${APP_ID}\Capabilities"

!macro customInstall
  !ifmacrodef verifyDistribution
    !insertmacro verifyDistribution
  !endif
  ClearErrors
  WriteRegStr SHCTX "${FOURPS_CLASSES}\${PDF_PROGID}" "" "Documento PDF (${PRODUCT_NAME})"
  WriteRegStr SHCTX "${FOURPS_CLASSES}\${PDF_PROGID}\DefaultIcon" "" '$\"$INSTDIR\resources\icons\pdf.ico$\",0'
  WriteRegStr SHCTX "${FOURPS_CLASSES}\${PDF_PROGID}\shell\open\command" "" '$\"$INSTDIR\${APP_EXECUTABLE_FILENAME}$\" --open-pdf $\"%1$\"'
  WriteRegStr SHCTX "${FOURPS_CLASSES}\.pdf\OpenWithProgids" "${PDF_PROGID}" ""
  WriteRegStr SHCTX "${FOURPS_CLASSES}\Applications\${APP_EXECUTABLE_FILENAME}" "FriendlyAppName" "${PRODUCT_NAME}"
  WriteRegStr SHCTX "${FOURPS_CLASSES}\Applications\${APP_EXECUTABLE_FILENAME}\SupportedTypes" ".pdf" ""
  WriteRegStr SHCTX "${FOURPS_CLASSES}\Applications\${APP_EXECUTABLE_FILENAME}\shell\open\command" "" '$\"$INSTDIR\${APP_EXECUTABLE_FILENAME}$\" --open-pdf $\"%1$\"'
  WriteRegStr SHCTX "${PDF_CAPABILITIES}" "ApplicationName" "${PRODUCT_NAME}"
  WriteRegStr SHCTX "${PDF_CAPABILITIES}" "ApplicationDescription" "Organizar e editar PDFs localmente"
  WriteRegStr SHCTX "${PDF_CAPABILITIES}" "ApplicationIcon" '$\"$INSTDIR\${APP_EXECUTABLE_FILENAME}$\",0'
  WriteRegStr SHCTX "${PDF_CAPABILITIES}\FileAssociations" ".pdf" "${PDF_PROGID}"
  WriteRegStr SHCTX "${FOURPS_REGISTERED}" "${APP_ID}" "${PDF_CAPABILITIES}"
  IfErrors fourps_association_failed fourps_association_ready
  fourps_association_failed:
    !ifmacrodef fourpsAbortUpgrade
      StrCpy $R3 "Não foi possível registrar a associação própria de PDFs."
      StrCpy $R2 "2"
      !insertmacro fourpsAbortUpgrade
    !else
      SetErrorLevel 2
      Abort
    !endif
  fourps_association_ready:
  System::Call 'shell32::SHChangeNotify(i 0x08000000, i 0, p 0, p 0)'
  !ifmacrodef fourpsCommitUpgrade
    !insertmacro fourpsCommitUpgrade
  !endif
!macroend

!macro customUnInstall
  ; No registration changes before file removal succeeds.
  !ifdef FOURPS_UPGRADE_LAB
    StrCpy $oldDesktopLink "$INSTDIR\..\lab-desktop\${SHORTCUT_NAME}.lnk"
    StrCpy $oldStartMenuLink "$INSTDIR\..\lab-programs\${SHORTCUT_NAME}.lnk"
  !endif
!macroend

!macro fourpsUnregisterPdf
  ReadRegStr $R0 SHCTX "${FOURPS_CLASSES}\${PDF_PROGID}\shell\open\command" ""
  ${if} $R0 == '$\"$INSTDIR\${APP_EXECUTABLE_FILENAME}$\" --open-pdf $\"%1$\"'
    DeleteRegKey SHCTX "${FOURPS_CLASSES}\${PDF_PROGID}"
    DeleteRegValue SHCTX "${FOURPS_CLASSES}\.pdf\OpenWithProgids" "${PDF_PROGID}"
    DeleteRegKey /ifempty SHCTX "${FOURPS_CLASSES}\.pdf\OpenWithProgids"
    DeleteRegKey /ifempty SHCTX "${FOURPS_CLASSES}\.pdf"
    DeleteRegKey SHCTX "${FOURPS_CLASSES}\Applications\${APP_EXECUTABLE_FILENAME}"
    DeleteRegValue SHCTX "${FOURPS_REGISTERED}" "${APP_ID}"
    DeleteRegKey SHCTX "${PDF_CAPABILITIES}"
    !ifndef FOURPS_UPGRADE_LAB
      DeleteRegKey /ifempty SHCTX "Software\${APP_ID}"
    !endif
  ${endif}
  System::Call 'shell32::SHChangeNotify(i 0x08000000, i 0, p 0, p 0)'
!macroend

!macro fourpsUnregisterCallback
!ifdef BUILD_UNINSTALLER
Function un.onUninstSuccess
  !insertmacro fourpsUnregisterPdf
FunctionEnd
!endif
!macroend
