; 论文格式医生 · Windows 安装包 (NSIS)
; 把 Nuitka standalone 文件夹打包成单个 exe 安装器。
; 在仓库根目录运行：`makensis installer.nsi`（dist/thesis-format-doctor-desktop/ 须已存在）
; 版本号可用 `makensis /DVERSION=1.3.103 installer.nsi` 覆盖。
!ifndef VERSION
  !define VERSION "dev"
!endif

!define APPNAME "论文格式医生"
!define APPDIR  "ThesisFormatDoctor"
!define EXE     "thesis-format-doctor-desktop.exe"
!define DIST    "dist\thesis-format-doctor-desktop"

Name "${APPNAME}"
OutFile "thesis-format-doctor-desktop-setup.exe"
InstallDir "$LOCALAPPDATA\Programs\${APPDIR}"
RequestExecutionLevel user          ; 每用户安装，无需 UAC 提权弹窗

!include "MUI2.nsh"
!define MUI_ABORTWARNING
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES
!insertmacro MUI_LANGUAGE "SimpChinese"

Section "Main" SecMain
  SetOutPath "$INSTDIR"
  File /r "${DIST}\*"

  ; 开始菜单 + 桌面快捷方式（用户只看得到一个图标，背后文件夹不可见）
  CreateDirectory "$SMPROGRAMS\${APPNAME}"
  CreateShortCut "$SMPROGRAMS\${APPNAME}\${APPNAME}.lnk" "$INSTDIR\${EXE}"
  CreateShortCut "$DESKTOP\${APPNAME}.lnk" "$INSTDIR\${EXE}"

  ; 卸载程序 + 注册表项
  WriteUninstaller "$INSTDIR\Uninstall.exe"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPDIR}" "DisplayName" "${APPNAME}"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPDIR}" "UninstallString" "$INSTDIR\Uninstall.exe"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPDIR}" "DisplayVersion" "${VERSION}"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPDIR}" "Publisher" "芦苇不熬夜"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPDIR}" "URLInfoAbout" "https://reedskill.com"
SectionEnd

Section "Uninstall"
  Delete "$DESKTOP\${APPNAME}.lnk"
  RMDir /r "$SMPROGRAMS\${APPNAME}"
  RMDir /r "$INSTDIR"
  DeleteRegKey HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPDIR}"
SectionEnd
