# -*- coding: utf-8 -*-
"""
Generate the usage note (使用说明.txt) for the Windows "folder zip" build.

v1.3.108 redesign: the zip is now a PURE PORTABLE build. There are NO
.vbs/.bat install scripts any more. The app itself creates its own desktop
shortcut via the win32com that is already bundled inside the exe (see gui.py
"桌面图标" link + first-run prompt) - so the customer machine needs ZERO extra
components (no VBScript, no PowerShell, no .NET) and there is no console
window anywhere in the customer journey.

This script only writes 使用说明.txt (GBK encoded so Chinese Notepad shows it
correctly). Keep ALL console output ASCII-safe: GitHub Actions Windows runners
use a cp1252 console, and printing Chinese filenames crashed CI in v1.3.107
(UnicodeEncodeError). We also reconfigure stdout/stderr to UTF-8 as
belt-and-braces.

Usage:
  python tools/make_portable_assets.py --dist dist/<appname> --exe <appname>.exe \
      --product "论文格式医生"
"""
import argparse
import os
import sys

# GitHub Actions Windows runner console is cp1252; printing Chinese would raise
# UnicodeEncodeError (this exact bug broke the v1.3.107 build). Reconfigure to
# UTF-8 AND keep prints ASCII-safe below.
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

# GBK-safe text only (no ★ / → / emoji). Curly quotes and full-width brackets
# are inside GBK so they are fine.
TMPL = """{product} 使用说明（免安装绿色版）

【启动】
1. 双击本文件夹里的“{exe}”即可启动，无需安装任何组件。
2. 首次启动会自动在桌面创建“{product}”快捷方式；
   之后双击桌面图标就能打开，不用每次进文件夹。
   （若桌面没有图标，打开软件后点右上角“桌面图标”文字链接即可补建。）

【提醒】
· 文件夹里除“{exe}”和“2.使用说明.txt”外的其它文件，是软件运行的
  必需组件，一个都不能删，但也不用管它们。

【卸载】
1. 关闭软件，删除桌面上的快捷方式；
2. 删除整个文件夹即完成卸载，不残留任何系统文件。

【提示】
· 软件完全在本机处理，论文文件不会上传任何服务器。
· 若需要“安装到电脑、可在系统设置中卸载”的版本，
  可到官网下载安装版（setup.exe）。
· 若个别杀毒软件提示，属误报（主程序多引擎检测 0 报），
  可放心使用或加入白名单。
"""


def main():
    ap = argparse.ArgumentParser(description="Generate usage note for Windows zip.")
    ap.add_argument("--dist", required=True,
                    help="Path to the Nuitka standalone folder (dist/<appname>)")
    ap.add_argument("--exe", required=True,
                    help="Main executable file name, e.g. 1.论文格式医生.exe")
    ap.add_argument("--product", required=True,
                    help="Display name (Chinese), e.g. 论文格式医生")
    # kept for backward CLI compatibility (old CI command lines pass it)
    ap.add_argument("--install-sub", default="",
                    help="Deprecated, ignored.")
    args = ap.parse_args()

    dist = args.dist
    if not os.path.isdir(dist):
        print("SKIP: dist folder not found:", dist, file=sys.stderr)
        sys.exit(0)

    note = TMPL.replace("{product}", args.product).replace("{exe}", args.exe)
    with open(os.path.join(dist, "2.使用说明.txt"), "w", encoding="gbk") as f:
        f.write(note)

    print("usage note written to", dist)


if __name__ == "__main__":
    main()
