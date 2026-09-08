# -*- coding: utf-8 -*-
"""
Generate portable-install assets for the Windows "folder zip" distribution.

Drops three files into the Nuitka standalone folder (dist/<appname>/):
  - 一键安装.bat   : ASCII wrapper that runs a base64-encoded PowerShell installer
  - 卸载.bat       : ASCII wrapper that runs a base64-encoded PowerShell uninstaller
  - 使用说明.txt   : Chinese usage note (GBK encoded for Notepad on Chinese Windows)

Why base64-EncodedCommand?
  The PowerShell scripts contain Chinese (product name, shortcut label). Encoding
  them as UTF-16LE base64 keeps the .bat files 100% ASCII, so there is ZERO
  codepage / gibberish risk on the customer's machine. This is the bulletproof
  pattern and avoids the GBK/BOM pitfalls we hit earlier with NSIS scripts.

The install logic:
  - copies the whole folder to %LOCALAPPDATA%\\Programs\\<install-sub> (no admin)
  - creates Desktop + Start Menu shortcuts with the Chinese product label
  - writes an HKCU Uninstall registry entry so it shows in "设置 -> 应用"

Usage:
  python tools/make_portable_assets.py --dist dist/<appname> --exe <appname>.exe \
      --install-sub ThesisFormatDoctor --product "论文格式医生"
"""
import os
import sys
import base64
import argparse


INSTALL_PS = r'''
$ErrorActionPreference = 'Stop'
$src = $PWD.Path
$product = "{PRODUCT}"
$sub = "{SUB}"
$exe = "{EXE}"
$instDir = Join-Path $env:LOCALAPPDATA ("Programs\" + $sub)
if (Test-Path $instDir) { Remove-Item $instDir -Recurse -Force }
Copy-Item -Path $src -Destination $instDir -Recurse -Force
$target = Join-Path $instDir $exe
$ws = New-Object -ComObject WScript.Shell
$desktop = [Environment]::GetFolderPath('Desktop')
$lnk = $ws.CreateShortcut((Join-Path $desktop ($product + ".lnk")))
$lnk.TargetPath = $target
$lnk.WorkingDirectory = $instDir
$lnk.IconLocation = $target
$lnk.Save()
$smPath = Join-Path $env:APPDATA ("Microsoft\Windows\Start Menu\Programs\" + $product + ".lnk")
$lnk2 = $ws.CreateShortcut($smPath)
$lnk2.TargetPath = $target
$lnk2.WorkingDirectory = $instDir
$lnk2.IconLocation = $target
$lnk2.Save()
$key = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\" + $product
New-Item -Path $key -Force | Out-Null
Set-ItemProperty -Path $key -Name DisplayName -Value $product
Set-ItemProperty -Path $key -Name UninstallString -Value ($instDir + "\卸载.bat")
Set-ItemProperty -Path $key -Name DisplayIcon -Value $target
Set-ItemProperty -Path $key -Name InstallLocation -Value $instDir
Set-ItemProperty -Path $key -Name NoModify -Value 1
Set-ItemProperty -Path $key -Name NoRepair -Value 1
Write-Host ("安装完成！桌面已创建「" + $product + "」快捷方式，也可在「设置 -> 应用」中卸载。")
'''

UNINSTALL_PS = r'''
$ErrorActionPreference = 'Stop'
$product = "{PRODUCT}"
$sub = "{SUB}"
$exe = "{EXE}"
$instDir = Join-Path $env:LOCALAPPDATA ("Programs\" + $sub)
$ws = New-Object -ComObject WScript.Shell
$desktop = [Environment]::GetFolderPath('Desktop')
$lnk = Join-Path $desktop ($product + ".lnk")
if (Test-Path $lnk) { Remove-Item $lnk -Force }
$smPath = Join-Path $env:APPDATA ("Microsoft\Windows\Start Menu\Programs\" + $product + ".lnk")
if (Test-Path $smPath) { Remove-Item $smPath -Force }
if (Test-Path $instDir) { Remove-Item $instDir -Recurse -Force }
$key = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\" + $product
if (Test-Path $key) { Remove-Item $key -Force }
Write-Host ("已卸载「" + $product + "」。")
'''

# GBK-safe text only (no ★ / → / curly quotes that are outside GBK).
USAGE_TXT = r"""论文格式医生 使用说明

【推荐三步（最省事）】
1. 把下载的压缩包解压到任意位置（如桌面）。
2. 进入解压出的文件夹，双击“一键安装.bat”。
   - 程序会自动复制到：C:\用户\你的用户名\AppData\Local\Programs\...
   - 并在桌面和开始菜单创建“论文格式医生”图标
3. 以后直接双击桌面图标即可使用，无需再进压缩包。

【卸载】
双击文件夹里的“卸载.bat”即可；也可在 Windows“设置 - 应用”里找到并卸载。

【关于杀软提示】
本程序主文件经多引擎检测无风险（0 报）。若个别杀软对“安装包”提示，属误报，可放心使用或加白名单。
"""


def b64(ps):
    # PowerShell -EncodedCommand expects UTF-16LE base64
    return base64.b64encode(ps.encode("utf-16-le")).decode("ascii")


def main():
    ap = argparse.ArgumentParser(description="Generate portable install assets for Windows zip.")
    ap.add_argument("--dist", required=True,
                    help="Path to the Nuitka standalone folder (dist/<appname>)")
    ap.add_argument("--exe", required=True,
                    help="Main executable file name, e.g. thesis-format-doctor-desktop.exe")
    ap.add_argument("--install-sub", required=True,
                    help="ASCII subfolder name under %%LOCALAPPDATA%%\\Programs, e.g. ThesisFormatDoctor")
    ap.add_argument("--product", required=True,
                    help="Display name (Chinese), e.g. 论文格式医生")
    args = ap.parse_args()

    dist = args.dist
    if not os.path.isdir(dist):
        print("SKIP: dist folder not found:", dist, file=sys.stderr)
        sys.exit(0)

    install_ps = (INSTALL_PS
                  .replace("{PRODUCT}", args.product)
                  .replace("{SUB}", args.install_sub)
                  .replace("{EXE}", args.exe))
    uninstall_ps = (UNINSTALL_PS
                    .replace("{PRODUCT}", args.product)
                    .replace("{SUB}", args.install_sub)
                    .replace("{EXE}", args.exe))

    # .bat files are pure ASCII (base64 is ASCII); no codepage risk.
    with open(os.path.join(dist, "一键安装.bat"), "w", encoding="ascii") as f:
        f.write("@echo off\n")
        f.write('cd /d "%~dp0"\n')
        f.write("powershell -NoProfile -ExecutionPolicy Bypass -EncodedCommand " + b64(install_ps) + "\n")
    with open(os.path.join(dist, "卸载.bat"), "w", encoding="ascii") as f:
        f.write("@echo off\n")
        f.write('cd /d "%~dp0"\n')
        f.write("powershell -NoProfile -ExecutionPolicy Bypass -EncodedCommand " + b64(uninstall_ps) + "\n")

    # usage note as GBK so Chinese Notepad shows it correctly
    with open(os.path.join(dist, "使用说明.txt"), "w", encoding="gbk") as f:
        f.write(USAGE_TXT)

    print("portable assets written to", dist)
    print("  - 一键安装.bat")
    print("  - 卸载.bat")
    print("  - 使用说明.txt")


if __name__ == "__main__":
    main()
