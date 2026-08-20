# -*- coding: utf-8 -*-
"""
跨平台构建脚本：把论文格式医生打包为单文件可执行程序。
  Windows -> thesis-format-doctor-desktop.exe
  macOS   -> thesis-format-doctor-desktop.app (--onefile 下为 Unix 可执行)
  Linux   -> thesis-format-doctor-desktop (可执行)

用法:
  python build.py            # 按当前平台构建到 dist/
  python build.py --clean   # 先清理再构建
"""
import os
import sys
import shutil
import subprocess

HERE = os.path.dirname(os.path.abspath(__file__))
SEP = ";" if sys.platform.startswith("win") else ":"

APP_NAME = "thesis-format-doctor-desktop"
ENTRY = os.path.join(HERE, "main.py")

# 引擎模块（标准库，但用 importlib 动态加载，必须显式声明）
HIDDEN = [
    "docxutils", "format_checker", "headings_fix",
    "ref_reformat", "format_profile", "report_docx", "format_check",
]

# 需作为"数据文件"打包的目录: (源目录, 打包后目录名)
DATA_DIRS = [
    (os.path.join(HERE, "tfd_app", "frontend"), "frontend"),
    (os.path.join(HERE, "tfd_app", "assets"), "assets"),
]


def build(clean=False):
    pyinstaller = os.path.join(
        os.path.dirname(sys.executable), "pyinstaller"
    )
    # 优先使用 venv 中的 pyinstaller
    venv_scripts = os.path.join(
        os.path.dirname(os.path.dirname(sys.executable)), "Scripts", "pyinstaller.exe"
    )
    if os.path.exists(venv_scripts):
        pyinstaller = venv_scripts
    elif not shutil.which("pyinstaller"):
        # 回退：用 python -m PyInstaller
        pyinstaller = None

    cmd = [sys.executable, "-m", "PyInstaller"] if pyinstaller is None else [pyinstaller]
    cmd += [
        "--name", APP_NAME,
        "--onefile",
        "--noconsole",
        "--clean" if clean else "",
        "--paths", HERE,
        "--paths", os.path.join(HERE, "tfd_app", "core"),
    ]
    for h in HIDDEN:
        cmd += ["--hidden-import", h]
    for src, dst in DATA_DIRS:
        if os.path.isdir(src):
            cmd += ["--add-data", f"{src}{SEP}{dst}"]

    cmd.append(ENTRY)

    # 过滤空串
    cmd = [c for c in cmd if c != ""]

    print(">>> " + " ".join(cmd))
    env = os.environ.copy()
    env["PYTHONHASHSEED"] = "1"
    rc = subprocess.call(cmd, cwd=HERE, env=env)
    if rc != 0:
        print("构建失败，返回码", rc)
        sys.exit(rc)
    print(f"\n构建完成 -> {os.path.join(HERE, 'dist', APP_NAME + ('.exe' if sys.platform.startswith('win') else ''))}")


if __name__ == "__main__":
    build(clean="--clean" in sys.argv)
