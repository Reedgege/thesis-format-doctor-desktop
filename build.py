# -*- coding: utf-8 -*-
"""
跨平台构建脚本：把论文格式医生打包为单文件可执行程序（原生 Tkinter 窗口，无浏览器）。
  Windows -> thesis-format-doctor-desktop.exe
  macOS   -> thesis-format-doctor-desktop
  Linux   -> thesis-format-doctor-desktop

用法:
  python build.py            # 按当前平台构建到 dist/
  python build.py --clean   # 先清理再构建
"""
import os
import sys
import subprocess

# Windows CI 控制台默认 cp1252，无法输出中文会抛 UnicodeEncodeError，
# 强制 stdout/stderr 用 utf-8，避免构建脚本自身因中文打印以非零退出。
try:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

HERE = os.path.dirname(os.path.abspath(__file__))
SEP = ";" if sys.platform.startswith("win") else ":"

APP_NAME = "thesis-format-doctor-desktop"
ENTRY = os.path.join(HERE, "main.py")

# v1.3.52：按平台选小羽毛图标（统一各平台 exe 文件图标）
# Windows: .ico（PyInstaller 直接嵌入为 exe 资源）
# macOS:   .icns（CI 步骤从 icon.png 生成后提交，见 .github/workflows/build.yml）
# Linux:   .png（PyInstaller 6+ 支持，onefile 内嵌为 ELF 资源）
ICON_ICO = os.path.join(HERE, "tfd_app", "assets", "icon.ico")
ICON_ICNS = os.path.join(HERE, "tfd_app", "assets", "icon.icns")
ICON_PNG = os.path.join(HERE, "tfd_app", "assets", "icon.png")


def _icon_for_build():
    """按当前平台返回可用的图标路径；缺失则退回 PNG（避免构建失败）。"""
    if sys.platform.startswith("win") and os.path.isfile(ICON_ICO):
        return ICON_ICO
    if sys.platform == "darwin" and os.path.isfile(ICON_ICNS):
        return ICON_ICNS
    return ICON_PNG

# 引擎模块（标准库，但用 importlib 动态加载，必须显式声明 hiddenimport）
# 注意：tfd_app.engine / tfd_app.gui 必须显式列出，否则 PyInstaller 静态分析
# 追踪不到"裸 import engine"这类运行时才解析的导入，打包后运行报 No module named 'engine'。
HIDDEN = [
    "docxutils", "format_checker", "headings_fix",
    "ref_reformat", "format_profile", "report_docx", "format_check",
    "tfd_app.engine", "tfd_app.gui", "tfd_app.license",
    # v1.3.57：试用版计数 + 水印（tfd_app 包内模块，gui 用相对导入 from . import）
    "tfd_app.trial", "tfd_app.watermark",
    # Windows 专用：进程内调用本机 Word/WPS 转换 .doc/.wps（pywin32 COM）
    "win32com", "win32com.client", "pythoncom", "pywintypes",
]

# 需作为"数据文件"打包的目录: (源目录, 打包后目录名)
DATA_DIRS = [
    (os.path.join(HERE, "tfd_app", "assets"), "tfd_app/assets"),
]

# 明确排除用不到的模块，减小 exe 体积（unittest/在线文档/演示程序等）。
# 保守起见只排除确定不用的；urllib 等网络组件依赖的模块一律保留。
EXCLUDES = [
    "unittest", "pydoc", "pydoc_data", "lib2to3", "idlelib",
    "turtledemo", "ensurepip", "test", "tkinter.test", "distutils",
    "http.server", "webbrowser", "xmlrpc", "telnetlib",
]


def build(clean=False):
    # 始终用 `python -m PyInstaller`，避免依赖 PATH 中的 pyinstaller 可执行文件。
    cmd = [sys.executable, "-m", "PyInstaller"]
    cmd += [
        "--name", APP_NAME,
        "--onefile",
        "--noconsole",
        "--clean" if clean else "",
        "--paths", HERE,
        "--paths", os.path.join(HERE, "tfd_app", "core"),
        # Tk 资源：让 PyInstaller 把 tcl/tk 运行时一并打进单文件
        "--collect-all", "tkinter",
    ]
    for h in HIDDEN:
        cmd += ["--hidden-import", h]
    for m in EXCLUDES:
        cmd += ["--exclude-module", m]
    for src, dst in DATA_DIRS:
        if os.path.isdir(src):
            cmd += ["--add-data", f"{src}{SEP}{dst}"]
    # v1.3.52：统一各平台 exe 文件图标为小羽毛（gui 窗口图标已用 icon.png 跨平台）
    cmd += ["--icon", _icon_for_build()]

    cmd.append(ENTRY)
    cmd = [c for c in cmd if c != ""]  # 过滤空串

    print(">>> " + " ".join(cmd))
    env = os.environ.copy()
    env["PYTHONHASHSEED"] = "1"
    rc = subprocess.call(cmd, cwd=HERE, env=env)
    if rc != 0:
        print("Build failed with return code", rc)
        sys.exit(rc)
    print("\nBuild complete -> " + os.path.join(HERE, "dist",
          APP_NAME + (".exe" if sys.platform.startswith("win") else "")))


if __name__ == "__main__":
    build(clean="--clean" in sys.argv)
