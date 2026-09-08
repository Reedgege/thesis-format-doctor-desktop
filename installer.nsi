; 论文格式医生 · Windows 安装包 (NSIS)
; 把 Nuitka standalone 文件夹打包成单个 exe 安装器。
; 在仓库根目录运行：`makensis installer.nsi`（dist/thesis-format-doctor-desktop/ 须已存在）
; 版本号可用 `makensis /DVERSION=1.3.105 installer.nsi` 覆盖。
; 🔴 编码铁律：本文件必须保存为 UTF-8 带 BOM！NSIS 3 只有读到 BOM 才按 UTF-8
; 解析脚本；否则按系统 ANSI 码页读（CI 英文 runner ACP=1252），中文全变乱码
; （症状：安装界面/快捷方式/注册表名显示 è®ºæ–‡... 乱码）。
Unicode true
!ifndef VERSION
  !define VERSION "dev"
!endif

!define APPNAME "论文格式医生"
!define APPDIR  "ThesisFormatDoctor"
!define EXE     "1.论文格式医生.exe"
!define DIST    "dist\论文格式医生（免安装版）"

Name "${APPNAME}"
OutFile "thesis-format-doctor-desktop-setup.exe"
InstallDir "$LOCALAPPDATA\Programs\${APPDIR}"
RequestExecutionLevel user          ; 每用户安装，无需 UAC 提权弹窗

; 安装包自身的"正当性"版本资源（文件属性→详细信息）。VERNUM 由 CI 传纯数字版本
; （/DVERNUM=1.3.105），本地不带参数时为 0.0.0 占位，不影响构建。
!ifndef VERNUM
  !define VERNUM "0.0.0"
!endif
VIProductVersion "${VERNUM}.0"
VIAddVersionKey "ProductName" "${APPNAME}"
VIAddVersionKey "FileDescription" "${APPNAME} 安装程序"
VIAddVersionKey "FileVersion" "${VERNUM}"
VIAddVersionKey "ProductVersion" "${VERNUM}"
VIAddVersionKey "CompanyName" "芦苇不熬夜"
VIAddVersionKey "LegalCopyright" "Copyright (c) 芦苇不熬夜"

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
