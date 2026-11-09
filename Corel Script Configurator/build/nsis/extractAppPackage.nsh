; Adaptador para electron-builder26.18.0: sem editar node_modules, instalador
; padrão, assinatura do desinstalador ou lógica de atualização/perfil.
!include "${PROJECT_DIR}\node_modules\app-builder-lib\templates\nsis\include\extractAppPackage.nsh"
!macroundef extractUsing7za
!macro extractUsing7za FILE
  Push $OUTDIR
  SetOutPath "$PLUGINSDIR\4ps-extract"
  ClearErrors
  Nsis7z::Extract "${FILE}"
  Pop $R0
  SetOutPath $R0
  IfFileExists "$PLUGINSDIR\4ps-extract\${APP_EXECUTABLE_FILENAME}" 0 distribution_copy_failed
  IfFileExists "$PLUGINSDIR\4ps-extract\lo\runtime-manifest.json" 0 distribution_copy_failed

  ; Robocopy é do Windows. Sem /256, /MIR, /PURGE, ACL ou preferências globais.
  ; /IS /IT impede pular arquivos por tamanho/data; verificação criptográfica
  ; integral é feita antes de o instalador anunciar sucesso.
  nsExec::ExecToStack '"$SYSDIR\robocopy.exe" "$PLUGINSDIR\4ps-extract" "$OUTDIR" /E /COPY:DAT /R:1 /W:1 /IS /IT /XJ /MT:4 /NFL /NDL /NJH /NJS /NP'
  Pop $R2
  Pop $R3
  StrCmp $R2 "error" distribution_copy_failed
  StrCmp $R2 "timeout" distribution_copy_failed
  IntCmp $R2 8 distribution_copy_failed distribution_copy_valid distribution_copy_failed
  distribution_copy_failed:
    DetailPrint "$R3"
    MessageBox MB_OK|MB_ICONSTOP "Não foi possível copiar todos os arquivos do aplicativo. A instalação foi interrompida, sem alterar seu perfil." /SD IDOK
    SetErrorLevel 2
    Abort
  distribution_copy_valid:
  ; This is only our fixed private extraction directory, never the installation.
  RMDir /r "$PLUGINSDIR\4ps-extract"
!macroend
