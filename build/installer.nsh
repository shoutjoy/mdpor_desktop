; Web2Electron — NSIS (electron-builder)
; customCheckAppRunning 정의 시 빌더가 getProcessInfo.nsh 생략 → 여기서 포함
!include "getProcessInfo.nsh"
!include "nsDialogs.nsh"
!include "LogicLib.nsh"
!include "StdUtils.nsh"
Var pid
!ifndef BUILD_UNINSTALLER
Var DesktopShortcutCheckbox
Var TaskbarShortcutCheckbox
Var WantDesktopShortcut
Var WantTaskbarShortcut
!endif

!macro customCheckAppRunning
  Var /GLOBAL IsPowerShellAvailable
  nsExec::Exec `"$SYSDIR\WindowsPowerShell\v1.0\powershell.exe" -C "if (Get-Command Get-CimInstance -ErrorAction SilentlyContinue) { exit 0 } else { exit 1 }"`
  Pop $0
  ${if} $0 == 0
    nsExec::Exec `"$SYSDIR\WindowsPowerShell\v1.0\powershell.exe" -C "if ((Get-ExecutionPolicy -Scope Process) -eq 'Restricted') { exit 1 } else { exit 0 }"`
    Pop $0
  ${endIf}
  ${if} $0 != 0
    StrCpy $0 1
  ${endIf}
  StrCpy $IsPowerShellAvailable $0
  ; /T: 자식 프로세스까지 (GPU·렌더러 등이 exe 폴더 DLL 잠금)
  nsExec::Exec `taskkill /F /T /IM "${APP_EXECUTABLE_FILENAME}"`
  Pop $0
  Sleep 600
  ${if} $IsPowerShellAvailable == 0
    nsExec::Exec `"$SYSDIR\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -Command "Get-CimInstance Win32_Process -ErrorAction SilentlyContinue | Where-Object { $$_.ExecutablePath -and $$_.ExecutablePath.StartsWith('$INSTDIR', 'CurrentCultureIgnoreCase') } | ForEach-Object { Stop-Process -Id $$_.ProcessId -Force -ErrorAction SilentlyContinue }"`
    Pop $0
    Sleep 600
  ${endIf}
  !insertmacro _CHECK_APP_RUNNING
!macroend

!macro customInit
  !ifndef BUILD_UNINSTALLER
  StrCpy $WantDesktopShortcut "1"
  StrCpy $WantTaskbarShortcut "0"
  !endif
  IfSilent silent_skip
  MessageBox MB_YESNO|MB_ICONQUESTION "이전에 설치된 ${PRODUCT_NAME}이(가) 있으면 설치 과정에서 자동으로 제거한 뒤 새 버전을 설치합니다. 계속하시겠습니까?" IDYES silent_skip
  Quit
  silent_skip:
!macroend

!ifndef BUILD_UNINSTALLER
Function ShortcutOptionsPageCreate
  nsDialogs::Create 1018
  Pop $0
  ${If} $0 == error
    Abort
  ${EndIf}

  ${NSD_CreateLabel} 0 0 100% 24u "${PRODUCT_NAME} 바로가기 옵션을 선택하세요."
  Pop $0
  ${NSD_CreateCheckbox} 0 34u 100% 14u "바탕화면에 바로가기 만들기"
  Pop $DesktopShortcutCheckbox
  ${NSD_Check} $DesktopShortcutCheckbox
  ${NSD_CreateCheckbox} 0 56u 100% 14u "작업표시줄에 고정하기"
  Pop $TaskbarShortcutCheckbox
  ${NSD_CreateLabel} 0 82u 100% 28u "시작메뉴 바로가기는 자동으로 등록됩니다. 작업표시줄 고정은 Windows 정책에 따라 일부 환경에서 적용되지 않을 수 있습니다."
  Pop $0

  nsDialogs::Show
FunctionEnd

Function ShortcutOptionsPageLeave
  ${NSD_GetState} $DesktopShortcutCheckbox $WantDesktopShortcut
  ${NSD_GetState} $TaskbarShortcutCheckbox $WantTaskbarShortcut
FunctionEnd

!macro customPageAfterChangeDir
  Page custom ShortcutOptionsPageCreate ShortcutOptionsPageLeave
!macroend

!macro customInstall
  ${If} $WantDesktopShortcut == ${BST_CHECKED}
    CreateShortCut "$DESKTOP\${PRODUCT_NAME}.lnk" "$INSTDIR\${APP_EXECUTABLE_FILENAME}" "" "$INSTDIR\${APP_EXECUTABLE_FILENAME}" 0
  ${EndIf}

  ${If} $WantTaskbarShortcut == ${BST_CHECKED}
    ${StdUtils.InvokeShellVerb} $0 "$INSTDIR" "${APP_EXECUTABLE_FILENAME}" ${StdUtils.Const.ShellVerb.PinToTaskbar}
  ${EndIf}
!macroend
!endif
