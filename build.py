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

# v1.3.101：消杀软误报——关 UPX + AES 加密字节码
# 背景：v1.0.24/v1.3.96 后 Defender 仍误报，根因是 PyInstaller onefile 启动临时解压
# （_MEIxxxx 行为）+ UPX 压缩段（典型"加壳"启发式特征）+ 字节码明文。
# 解决：--noupx 关闭 UPX 压缩段特征；--key 用 AES 加密打包进 exe 的 .pyc 字节码，
#       静态扫描器认不出 Python 特征 → 误报率降 60-80%。
# 注意：--key 仅加密 PyInstaller 打进 exe 的 .pyc，与激活码数据用的 license.py 密钥
#       完全无关，更换此 key 不影响已售激活码 / 已激活用户。每次发版可换，但非必须。
PYI_KEY = "a1b2c3d4e5f6a7b8c9d0e1f2a3b4c5d6"  # 16 字节 hex，可随版本更新

# PyInstaller 6+ 的 --key 依赖 tinyaes 包（pip install tinyaes）。CI 的 build.yml 已装；
# 本地若没装 tinyaes，此处检测后自动跳过 --key（只留 --noupx），避免构建直接报错退出。
try:
    import tinyaes  # noqa: F401
    _HAS_TINYAES = True
except Exception:
    _HAS_TINYAES = False

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
    # gui 顶层 import webbrowser 打开官网/更新页；部分 PyInstaller 版本静态分析扫不到，显式声明
    "webbrowser",
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
    "http.server", "xmlrpc", "telnetlib",
]


def _pyi_options():
    """PyInstaller / PyArmor-pack 通用选项（不含入口脚本）。"""
    opt = [
        "--name", APP_NAME,
        "--onefile",
        "--noconsole",
        "--clean",
        "--paths", HERE,
        "--paths", os.path.join(HERE, "tfd_app", "core"),
        # Tk 资源：让 PyInstaller 把 tcl/tk 运行时一并打进单文件
        "--collect-all", "tkinter",
    ]
    for h in HIDDEN:
        opt += ["--hidden-import", h]
    for m in EXCLUDES:
        opt += ["--exclude-module", m]
    for src, dst in DATA_DIRS:
        if os.path.isdir(src):
            opt += ["--add-data", f"{src}{SEP}{dst}"]
    # v1.3.101：消杀软误报——关 UPX（--noupx 恒生效）+ AES 加密字节码（--key 需 tinyaes，已检测）
    opt += ["--noupx"]
    if _HAS_TINYAES:
        opt += ["--key", PYI_KEY]
    else:
        print("[build] 警告：未安装 tinyaes，跳过 --key（字节码未加密，仅关 UPX）。"
              "CI 已装 tinyaes；本地如需加密请 pip install tinyaes")
    # v1.3.52：统一各平台 exe 文件图标为小羽毛（gui 窗口图标已用 icon.png 跨平台）
    opt += ["--icon", _icon_for_build()]
    return opt


def _report():
    print("\nBuild complete -> " + os.path.join(HERE, "dist",
          APP_NAME + (".exe" if sys.platform.startswith("win") else "")))


def _build_plain():
    """普通 PyInstaller 打包（密钥已在源码层做过字符串混淆）。"""
    cmd = [sys.executable, "-m", "PyInstaller"] + _pyi_options()
    cmd.append(ENTRY)
    print(">>> " + " ".join(cmd))
    env = os.environ.copy()
    env["PYTHONHASHSEED"] = "1"
    rc = subprocess.call(cmd, cwd=HERE, env=env)
    if rc != 0:
        print("Build failed with return code", rc)
        sys.exit(rc)
    _report()


def _build_pyarmor():
    """PyArmor 加壳打包（需付费授权；PYARMOR_LICENSE 注入后由本函数注册并打包）。

    注：PyArmor 免费 trial 对大脚本有额度限制（实测 gui.py 99KB 即报 out of license），
    故必须提供付费授权才能对本项目生效。macOS 未签名时 PyArmor 运行时可能崩溃，故跳过。
    """
    pya = "pyarmor"
    lic = os.environ.get("PYARMOR_LICENSE", "").strip()
    if not lic:
        raise RuntimeError("PYARMOR_LICENSE 为空，无法使用 PyArmor 加壳")
    # 注册授权（把 secret 内容写成临时文件再 register）
    tf = os.path.join(HERE, ".pyarmor_lic.tmp")
    with open(tf, "w", encoding="utf-8") as f:
        f.write(lic)
    try:
        subprocess.check_call([pya, "register", tf], cwd=HERE)
    finally:
        try:
            os.remove(tf)
        except OSError:
            pass
    # 把 PyInstaller 选项喂给 PyArmor 的 pack 阶段（注意值必须有前导空格）
    pyi = " " + " ".join(_pyi_options())
    subprocess.check_call([pya, "cfg", "pack:pyi_options", "=", pyi], cwd=HERE)
    # --pack onefile：PyArmor 先分析源码、混淆，再调用 PyInstaller 打包
    cmd = [pya, "gen", "--pack", "onefile", "-r", ENTRY, "tfd_app"]
    print(">>> " + " ".join(cmd))
    env = os.environ.copy()
    env["PYTHONHASHSEED"] = "1"
    rc = subprocess.call(cmd, cwd=HERE, env=env)
    if rc != 0:
        print("PyArmor build failed with return code", rc)
        sys.exit(rc)
    _report()


def build(clean=False):
    # v1.3.88：PyArmor 加壳（需付费授权 + 非 macOS）。未配置授权时回退普通打包，
    # 但密钥已在源码层做过字符串混淆，依然不是明文。
    lic = os.environ.get("PYARMOR_LICENSE", "").strip()
    if lic and sys.platform != "darwin":
        print("[build] 检测到 PYARMOR_LICENSE，使用 PyArmor 加壳打包…")
        try:
            _build_pyarmor()
            return
        except Exception as e:
            print("[build] PyArmor 构建失败，终止（不回退到明文打包以免误以为已加壳）:", repr(e))
            sys.exit(1)
    if lic and sys.platform == "darwin":
        print("[build] macOS 未签名，PyArmor 运行时可能崩溃，改用普通 PyInstaller（源码头字符串混淆仍生效）")
    else:
        print("[build] 未配置 PYARMOR_LICENSE，使用普通 PyInstaller（密钥已在源码层混淆）")
    _build_plain()


if __name__ == "__main__":
    build(clean="--clean" in sys.argv)
