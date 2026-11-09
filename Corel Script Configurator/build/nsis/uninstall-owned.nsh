; Selective, long-path-aware uninstall. Unlisted/user files and profile stay.
!ifdef BUILD_UNINSTALLER
!include FileFunc.nsh
!include StrFunc.nsh
${UnStrRep}
${UnStrStr}
Var /GLOBAL fourpsUnBackup
Var /GLOBAL fourpsUnList
Var /GLOBAL fourpsUnCount
Var /GLOBAL fourpsUnIndex
Var /GLOBAL fourpsUnRelative

Function un.fourpsReadRelative
  ReadINIStr $fourpsUnRelative $fourpsUnList "files" $fourpsUnIndex
  Call un.fourpsValidateRelative
FunctionEnd

Function un.fourpsValidateRelative
  StrCmp $fourpsUnRelative "" fourps_un_bad_path
  ${UnStrStr} $R0 $fourpsUnRelative ".."
  StrCmp $R0 "" 0 fourps_un_bad_path
  ${UnStrStr} $R0 $fourpsUnRelative ":"
  StrCmp $R0 "" 0 fourps_un_bad_path
  StrCpy $R0 $fourpsUnRelative 1
  StrCmp $R0 "\" fourps_un_bad_path
  StrCmp $R0 "/" fourps_un_bad_path
  ${UnStrRep} $fourpsUnRelative $fourpsUnRelative "/" "\"
  ClearErrors
  Return
  fourps_un_bad_path:
    SetErrors
FunctionEnd

Function un.fourpsRestoreOwned
  ; The backup is persistent, never discarded if restoration fails.
  StrCpy $fourpsUnIndex 0
  fourps_un_restore_loop:
    IntCmp $fourpsUnIndex $fourpsUnCount fourps_un_restore_done fourps_un_restore_item fourps_un_restore_done
    fourps_un_restore_item:
    Call un.fourpsReadRelative
    IfErrors fourps_un_restore_failed
    IfFileExists "\\?\$fourpsUnBackup\$fourpsUnRelative" 0 fourps_un_restore_next
      ${GetParent} "$INSTDIR\$fourpsUnRelative" $R0
      CreateDirectory "\\?\$R0"
      ClearErrors
      Rename "\\?\$fourpsUnBackup\$fourpsUnRelative" "\\?\$INSTDIR\$fourpsUnRelative"
      IfErrors fourps_un_restore_failed
    fourps_un_restore_next:
    IntOp $fourpsUnIndex $fourpsUnIndex + 1
    Goto fourps_un_restore_loop
  fourps_un_restore_done:
    ClearErrors
    Return
  fourps_un_restore_failed:
    SetErrors
FunctionEnd
!endif

!macro customRemoveFiles
  InitPluginsDir
  IfFileExists "$INSTDIR\4ps-installed-files.ini" 0 fourps_un_invalid
  GetTempFileName $fourpsUnBackup "$INSTDIR\.."
  GetFullPathName $fourpsUnBackup $fourpsUnBackup
  Delete $fourpsUnBackup
  CreateDirectory $fourpsUnBackup
  StrCpy $fourpsUnList "$fourpsUnBackup\ownership.ini"
  ClearErrors
  CopyFiles /SILENT "$INSTDIR\4ps-installed-files.ini" $fourpsUnList
  IfErrors fourps_un_invalid
  ReadINIStr $R0 $fourpsUnList "identity" "appId"
  StrCmp $R0 "${APP_ID}" 0 fourps_un_invalid
  ReadINIStr $R0 $fourpsUnList "identity" "guid"
  StrCmp $R0 "${APP_GUID}" 0 fourps_un_invalid
  ReadINIStr $fourpsUnCount $fourpsUnList "files" "count"
  IntCmp $fourpsUnCount 1 fourps_un_count_ok fourps_un_invalid fourps_un_count_ok
  fourps_un_count_ok:
  IntCmp $fourpsUnCount 100000 fourps_un_begin fourps_un_begin fourps_un_invalid
  fourps_un_begin:
  ; Validate the complete manifest before moving the first file.
  StrCpy $fourpsUnIndex 0
  fourps_un_validate_files:
    IntCmp $fourpsUnIndex $fourpsUnCount fourps_un_validate_dirs fourps_un_validate_file fourps_un_validate_dirs
    fourps_un_validate_file:
    Call un.fourpsReadRelative
    IfErrors fourps_un_invalid
    IntOp $fourpsUnIndex $fourpsUnIndex + 1
    Goto fourps_un_validate_files
  fourps_un_validate_dirs:
  ReadINIStr $R1 $fourpsUnList "directories" "dirCount"
  IntCmp $R1 0 fourps_un_validated fourps_un_invalid fourps_un_validate_dir_count
  fourps_un_validate_dir_count:
  IntCmp $R1 100000 fourps_un_validate_dir_start fourps_un_validate_dir_start fourps_un_invalid
  fourps_un_validate_dir_start:
  StrCpy $R2 0
  fourps_un_validate_dir_loop:
    IntCmp $R2 $R1 fourps_un_validated fourps_un_validate_dir fourps_un_validated
    fourps_un_validate_dir:
    ReadINIStr $fourpsUnRelative $fourpsUnList "directories" "d$R2"
    Call un.fourpsValidateRelative
    IfErrors fourps_un_invalid
    IntOp $R2 $R2 + 1
    Goto fourps_un_validate_dir_loop
  fourps_un_validated:
  StrCpy $fourpsUnIndex 0
  fourps_un_loop:
    IntCmp $fourpsUnIndex $fourpsUnCount fourps_un_files_done fourps_un_item fourps_un_files_done
    fourps_un_item:
    Call un.fourpsReadRelative
    IfErrors fourps_un_failed
    StrCpy $7 "\\?\$INSTDIR\$fourpsUnRelative"
    System::Call 'kernel32::GetFileAttributesW(w r7) i.r9 ?e'
    Pop $8
    StrCmp $9 "-1" 0 fourps_un_move
      StrCmp $8 "2" fourps_un_next
      StrCmp $8 "3" fourps_un_next
      Goto fourps_un_failed
    fourps_un_move:
    ${GetParent} "$fourpsUnBackup\$fourpsUnRelative" $R0
    CreateDirectory "\\?\$R0"
    ClearErrors
    Rename "\\?\$INSTDIR\$fourpsUnRelative" "\\?\$fourpsUnBackup\$fourpsUnRelative"
    IfErrors fourps_un_failed
    fourps_un_next:
    IntOp $fourpsUnIndex $fourpsUnIndex + 1
    Goto fourps_un_loop
  fourps_un_files_done:
  ReadINIStr $R1 $fourpsUnList "directories" "dirCount"
  StrCpy $R2 0
  fourps_un_dirs:
    IntCmp $R2 $R1 fourps_un_dirs_done fourps_un_dir fourps_un_dirs_done
    fourps_un_dir:
    ReadINIStr $fourpsUnRelative $fourpsUnList "directories" "d$R2"
    Call un.fourpsValidateRelative
    IfErrors fourps_un_failed
    RMDir "\\?\$INSTDIR\$fourpsUnRelative"
    IntOp $R2 $R2 + 1
    Goto fourps_un_dirs
  fourps_un_dirs_done:
  SetOutPath $TEMP
  RMDir "\\?\$INSTDIR"
  Goto fourps_un_done
  fourps_un_failed:
    Call un.fourpsRestoreOwned
    MessageBox MB_OK|MB_ICONSTOP "A remoção foi interrompida. Arquivos de usuário e perfil foram preservados. Confira a recuperação em $fourpsUnBackup antes de tentar novamente." /SD IDOK
    SetErrorLevel 2
    Abort
  fourps_un_invalid:
    MessageBox MB_OK|MB_ICONSTOP "Inventário da instalação ausente ou inválido. Nenhum arquivo fora do inventário será removido." /SD IDOK
    SetErrorLevel 2
    Abort
  fourps_un_done:
!macroend
