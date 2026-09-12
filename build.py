# -*- coding: utf-8 -*-
"""
跨平台构建脚本（Nuitka 版）
==========================
把论文格式医生编译为原生二进制「文件夹版」，再用 NSIS 打安装包（Windows）。
彻底规避 PyInstaller 单文件（onefile）运行时自解压导致的杀软误报。

  Windows -> dist/论文格式医生（免安装版）/  + installer.nsi -> thesis-format-doctor-desktop-setup.exe
  macOS   -> dist/thesis-format-doctor-desktop.app  (--macos-create-app-bundle)
  Linux   -> dist/thesis-format-doctor-desktop/

原理：Nuita 把 Python 编译成 C -> 原生机器码。产物在 Windows Defender 眼里等同
普通 C++ 程序，报毒率极低（实测 2/71），且源码被编译保护、无法被 pyinstxtractor
扒出。客户侧无需安装 Python —— Nuitka standalone 会把 Python 运行时 + Tcl/Tk 打进文件夹。

用法:
  python build.py            # 按当前平台构建到 dist/
  python build.py --clean   # 先清理 dist/ 再构建
"""
import os
import sys
import shutil
import subprocess

# Windows CI 控制台默认 cp1252，强制 utf-8，避免中文打印以非零退出。
try:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

HERE = os.path.dirname(os.path.abspath(__file__))
WIN = sys.platform.startswith("win")
DARWIN = sys.platform == "darwin"

APP_NAME = "thesis-format-doctor-desktop"
# v1.3.109：Windows 分发名中文化——exe 带数字前缀排文件列表最前、dist 目录（即 zip
# 顶层文件夹）用中文产品名，客户解压后第一眼看到「1.论文格式医生.exe」。macOS/Linux
# 保持英文 APP_NAME（.app 包内/文件夹无"找 exe"问题，避免引入非 ASCII 路径的 CI 风险）。
WIN_EXE_NAME = "1.论文格式医生"
WIN_DIR_NAME = "论文格式医生（免安装版）"
ENTRY = os.path.join(HERE, "main.py")

# Windows 可执行文件的"正当性"版本资源（文件属性里的产品名/公司名/版本）。
# 未签名新 exe 被 Defender ML 误报的常见诱因之一是"来路不明、无版本信息"；
# 补全后特征更像正规软件，可压低 Wacatac/Sabsik 类启发式分数。
PRODUCT_NAME = "论文格式医生"
COMPANY_NAME = "芦苇不熬夜"

ICON_ICO = os.path.join(HERE, "tfd_app", "assets", "icon.ico")
ICON_ICNS = os.path.join(HERE, "tfd_app", "assets", "icon.icns")
ICON_PNG = os.path.join(HERE, "tfd_app", "assets", "icon.png")

# 本地引擎模块：core/ 下的纯 Python 模块，被代码以 `import docxutils` 这种
# 顶层导入方式引用（依赖运行时把 core 加入 sys.path）。Nuitka 静态分析追踪不到，
# 必须显式 --include-module 强制打包，并在编译期把 core 放入 PYTHONPATH 使其可解析。
CORE_MODULES = [
    "docxutils", "format_checker", "headings_fix",
    "ref_reformat", "format_profile", "report_docx", "format_check",
]


def _read_version():
    """读 VERSION 文件并去 'v' 前缀，返回纯数字版本（如 1.3.105）。"""
    try:
        with open(os.path.join(HERE, "VERSION"), encoding="utf-8") as fh:
            return fh.read().strip().lstrip("vV")
    except Exception:
        return ""


def _nuitka_options():
    opt = [
        sys.executable, "-m", "nuitka", ENTRY,
        "--standalone",                 # 文件夹模式：运行时不自解压 -> 避免 dropper 指纹
        "--enable-plugins=tk-inter",    # tkinter 支持（Tcl/Tk 一并打包）
        "--assume-yes-for-downloads",   # Nuitka 需下载 ccache/补丁时自动同意，不卡交互
        "--output-dir=" + os.path.join(HERE, "dist"),
        "--output-filename=" + (WIN_EXE_NAME if WIN else APP_NAME) + (".exe" if WIN else ""),
        "--remove-output",              # 构建中间目录 .build 用完即删，仅留 .dist/.app
        "--show-progress",
    ]
    # 强制包含本地模块（顶层导入，须显式声明）
    for m in CORE_MODULES:
        opt += ["--include-module=" + m]
    opt += ["--include-package=tfd_app"]
    # 运行时资源（窗口图标 / 二维码 / 水印图 / 主界面背景图）必须显式带进产物 ——
    # gui / watermark / backdrop 都会在运行时按 <包目录>/assets/<name> 读取，而读取方
    # 普遍带 isfile 守卫 → 漏带就会「二维码不显示、水印丢失、背景图缺失」且全程静默。
    # 清单以 tfd_app/assetpath.py 的 REQUIRED_ASSETS 为唯一事实来源；缺文件直接构建
    # 失败（fail-closed），宁可构建红灯，也不发出一个"看着正常、功能缺件"的包。
    sys.path.insert(0, HERE)
    from tfd_app.assetpath import REQUIRED_ASSETS
    for _name in REQUIRED_ASSETS:
        _src = os.path.join(HERE, "tfd_app", "assets", _name)
        if not os.path.isfile(_src):
            raise SystemExit(
                "构建中止：缺少运行时资源 %s（清单见 tfd_app/assetpath.py 的 REQUIRED_ASSETS）" % _src)
        opt += ["--include-data-files=" + _src + "=tfd_app/assets/" + _name]
    # 平台图标
    if WIN and os.path.isfile(ICON_ICO):
        opt += ["--windows-icon-from-ico=" + ICON_ICO]
    # Windows GUI 程序：禁用控制台子系统，否则启动时黑框一闪（Nuitka 默认 console）。
    # 并注入版本资源（文件属性里的公司/产品/版本），给未签名 exe 增加正当性信号。
    if WIN:
        opt += ["--windows-console-mode=disable"]
        _ver = _read_version()
        if _ver:
            opt += ["--company-name=" + COMPANY_NAME,
                    "--product-name=" + PRODUCT_NAME,
                    "--product-version=" + _ver,
                    "--file-description=" + PRODUCT_NAME + "（模板驱动论文格式检查与修正工具）",
                    "--file-version=" + _ver,
                    "--copyright=Copyright (c) " + COMPANY_NAME]
    if DARWIN:
        if os.path.isfile(ICON_ICNS):
            opt += ["--macos-app-icon=" + ICON_ICNS]
        opt += ["--macos-create-app-bundle",
                "--macos-app-mode=gui",
                "--macos-app-name=论文格式医生"]
    # Windows 专用：进程内调用本机 Word/WPS 转换 .doc/.wps（pywin32 COM）。
    # win32com 在函数内局部 import，Nuitka 不会自动收录，且仅 Windows 构建才装 pywin32。
    if WIN:
        try:
            import win32com  # noqa: F401
            opt += ["--include-module=win32com",
                    "--include-module=win32com.client",
                    "--include-module=pythoncom",
                    "--include-module=pywintypes"]
        except Exception:
            pass
    return opt


def _normalize_output():
    """Nuitka 的 standalone 产物目录名固定为 <入口Basename>.dist / .app，
    重命名为发布用目录名方便 NSIS / 压缩与发布。Windows 用中文目录名（zip 顶层
    直接就是中文文件夹，客户解压即见 exe）；macOS/Linux 保持英文 APP_NAME。"""
    dist = os.path.join(HERE, "dist")
    if DARWIN:
        src, dst = os.path.join(dist, "main.app"), os.path.join(dist, APP_NAME + ".app")
    elif WIN:
        src, dst = os.path.join(dist, "main.dist"), os.path.join(dist, WIN_DIR_NAME)
    else:
        src, dst = os.path.join(dist, "main.dist"), os.path.join(dist, APP_NAME)
    if os.path.isdir(src) and not os.path.exists(dst):
        shutil.move(src, dst)
        return dst
    if os.path.isdir(dst):
        return dst
    return src


def _report():
    out = _normalize_output()
    print("\nBuild complete -> " + out)
    if WIN:
        print("主程序           -> " + os.path.join(out, WIN_EXE_NAME + ".exe"))
        print("安装包           -> 仓库根目录运行 `makensis installer.nsi` 生成 setup.exe")
    elif DARWIN:
        print("App 包           -> " + out)
    else:
        print("主程序           -> " + os.path.join(out, APP_NAME))


def build(clean=False):
    if clean:
        d = os.path.join(HERE, "dist")
        if os.path.isdir(d):
            shutil.rmtree(d)
    # 编译期让 `import docxutils`（位于 tfd_app/core）可解析
    env = os.environ.copy()
    extra = os.pathsep.join([os.path.join(HERE, "tfd_app", "core"),
                             os.path.join(HERE, "tfd_app")])
    env["PYTHONPATH"] = (extra + os.pathsep + env["PYTHONPATH"]) if env.get("PYTHONPATH") else extra
    env["PYTHONHASHSEED"] = "1"
    cmd = _nuitka_options()
    print(">>> " + " ".join(cmd))
    rc = subprocess.call(cmd, cwd=HERE, env=env)
    if rc != 0:
        print("Build failed with return code", rc)
        sys.exit(rc)
    _report()


if __name__ == "__main__":
    build(clean="--clean" in sys.argv)
