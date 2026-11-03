; Registra capacidade/Open With sem tocar no default de .pdf ou em UserChoice.
!define PDF_PROGID "${APP_ID}.PDF"
!define PDF_CAPABILITIES "Software\${APP_ID}\Capabilities"

!macro customInstall
  WriteRegStr SHCTX "Software\Classes\${PDF_PROGID}" "" "Documento PDF (${PRODUCT_NAME})"
  WriteRegStr SHCTX "Software\Classes\${PDF_PROGID}\DefaultIcon" "" '$\"$INSTDIR\resources\icons\pdf.ico$\",0'
  WriteRegStr SHCTX "Software\Classes\${PDF_PROGID}\shell\open\command" "" '$\"$INSTDIR\${APP_EXECUTABLE_FILENAME}$\" --open-pdf $\"%1$\"'
  WriteRegStr SHCTX "Software\Classes\.pdf\OpenWithProgids" "${PDF_PROGID}" ""
  WriteRegStr SHCTX "Software\Classes\Applications\${APP_EXECUTABLE_FILENAME}" "FriendlyAppName" "${PRODUCT_NAME}"
  WriteRegStr SHCTX "Software\Classes\Applications\${APP_EXECUTABLE_FILENAME}\SupportedTypes" ".pdf" ""
  WriteRegStr SHCTX "Software\Classes\Applications\${APP_EXECUTABLE_FILENAME}\shell\open\command" "" '$\"$INSTDIR\${APP_EXECUTABLE_FILENAME}$\" --open-pdf $\"%1$\"'
  WriteRegStr SHCTX "${PDF_CAPABILITIES}" "ApplicationName" "${PRODUCT_NAME}"
  WriteRegStr SHCTX "${PDF_CAPABILITIES}" "ApplicationDescription" "Organizar e editar PDFs localmente"
  WriteRegStr SHCTX "${PDF_CAPABILITIES}" "ApplicationIcon" '$\"$INSTDIR\${APP_EXECUTABLE_FILENAME}$\",0'
  WriteRegStr SHCTX "${PDF_CAPABILITIES}\FileAssociations" ".pdf" "${PDF_PROGID}"
  WriteRegStr SHCTX "Software\RegisteredApplications" "${APP_ID}" "${PDF_CAPABILITIES}"
  System::Call 'shell32::SHChangeNotify(i 0x08000000, i 0, p 0, p 0)'
!macroend

!macro customUnInstall
  ReadRegStr $R0 SHCTX "Software\Classes\${PDF_PROGID}\shell\open\command" ""
  ${if} $R0 == '$\"$INSTDIR\${APP_EXECUTABLE_FILENAME}$\" --open-pdf $\"%1$\"'
    DeleteRegKey SHCTX "Software\Classes\${PDF_PROGID}"
    DeleteRegValue SHCTX "Software\Classes\.pdf\OpenWithProgids" "${PDF_PROGID}"
    DeleteRegKey /ifempty SHCTX "Software\Classes\.pdf\OpenWithProgids"
    DeleteRegKey /ifempty SHCTX "Software\Classes\.pdf"
    DeleteRegKey SHCTX "Software\Classes\Applications\${APP_EXECUTABLE_FILENAME}"
    DeleteRegValue SHCTX "Software\RegisteredApplications" "${APP_ID}"
    DeleteRegKey SHCTX "Software\${APP_ID}\Capabilities"
    DeleteRegKey /ifempty SHCTX "Software\${APP_ID}"
  ${endif}
  System::Call 'shell32::SHChangeNotify(i 0x08000000, i 0, p 0, p 0)'
!macroend
