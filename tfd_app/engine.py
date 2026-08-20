# -*- coding: utf-8 -*-
"""
论文格式医生 · 引擎封装层（桌面版内部使用，不直接给终端用户）

职责：
  1. 把 5 个纯标准库引擎脚本（format_profile / format_checker / format_check /
     headings_fix / ref_reformat）封装成「吃文件路径、返回文本」的干净函数，
     避免上层 GUI 去拼 sys.argv、也避免 subprocess 在冻结（PyInstaller onefile）
     环境下踩坑。
  2. 兼容老格式 .doc / WPS .wps：检测到 OLE2 复合文档时，若本机装有 LibreOffice，
     自动转成 .docx 再处理；否则给出明确提示让用户先另存为 .docx。

设计原则：不修改用户的原始引擎脚本（原创作品），只在外部做路径/输出编排。
"""
import os
import sys
import io
import json
import shutil
import subprocess
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
CORE = os.path.join(HERE, "core")
if CORE not in sys.path:
    sys.path.insert(0, CORE)

# 先确保 docxutils 可导入（其余引擎脚本 import 它）
import docxutils  # noqa: E402,F401
import format_checker  # noqa: E402
import headings_fix  # noqa: E402
import ref_reformat  # noqa: E402
import format_profile  # noqa: E402
import format_check  # noqa: E402


# ---------------------------------------------------------------------------
# stdout 捕获：引擎的 main() 用 print 输出报告，重定向到内存后返回给 GUI
# ---------------------------------------------------------------------------
def _run_module_main(module, argv):
    """在受控环境里执行引擎模块的 main()，捕获其 print 输出并返回文本字符串。

    无论引擎内部是否 sys.exit / 抛异常，都会恢复 sys.argv 与 sys.stdout。
    """
    old_argv = sys.argv
    old_stdout = sys.stdout
    buf = io.StringIO()
    sys.argv = argv
    sys.stdout = buf
    try:
        module.main()
    except SystemExit:
        pass
    except Exception:
        # 把异常原样抛出，让 GUI 捕获并提示
        raise
    finally:
        sys.stdout = old_stdout
        sys.argv = old_argv
    return buf.getvalue()


# ---------------------------------------------------------------------------
# 老格式（.doc / .wps）自动转换
# ---------------------------------------------------------------------------
_OLE2_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


def _looks_like_ole2(path):
    try:
        with open(path, "rb") as f:
            head = f.read(8)
        return head == _OLE2_MAGIC
    except Exception:
        return False


def find_soffice():
    """查找 LibreOffice 可执行文件（用于 .doc/.wps 转 .docx）。返回路径或 None。"""
    # 1) 环境变量 PATH 里直接能找到
    for name in ("soffice", "libreoffice"):
        p = shutil.which(name)
        if p:
            return p
    # 2) 常见安装位置
    candidates = []
    if sys.platform.startswith("win"):
        candidates = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ]
    elif sys.platform == "darwin":
        candidates = [
            "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        ]
    else:
        candidates = [
            "/usr/bin/libreoffice",
            "/usr/bin/soffice",
            "/snap/bin/libreoffice",
            "/opt/libreoffice/program/soffice",
        ]
    for c in candidates:
        if os.path.isfile(c):
            return c
    return None


def _convert_doc_to_docx(path, soffice, workdir):
    """调用 LibreOffice 把 .doc/.wps 转成 .docx，返回生成的 .docx 路径。"""
    out_dir = tempfile.mkdtemp(prefix="tfd_conv_", dir=workdir or None)
    cmd = [soffice, "--headless", "--convert-to", "docx", "--outdir", out_dir, path]
    try:
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=240)
    except Exception as e:
        raise RuntimeError("调用 LibreOffice 失败：" + str(e))
    if r.returncode != 0:
        err = (r.stderr or r.stdout or "").strip()[:600]
        raise RuntimeError("LibreOffice 转换失败：" + err)
    produced = None
    for f in os.listdir(out_dir):
        if f.lower().endswith(".docx"):
            produced = os.path.join(out_dir, f)
            break
    if not produced:
        raise RuntimeError("LibreOffice 未生成 .docx 文件，请确认文件未损坏。")
    return produced


# 用本机 Word / WPS（Windows COM，经系统脚本宿主调 VBScript）把 .doc/.wps 转 .docx。
# 纯标准库实现：脚本宿主是 Windows 自带，无需 pywin32；VBS 用 UTF-16 编码以支持中文路径。
# 第三个参数为可选的保存格式：空 = 由 .docx 扩展名决定；12 = Word 2007 XML(.docx)；16 = 默认 docx。
_VBS_CONVERT = r'''
Option Explicit
Dim app, doc, ok, dst, ff
ok = False
dst = WScript.Arguments(1)
ff = ""
If WScript.Arguments.Count > 2 Then
    ff = WScript.Arguments(2)
End If
On Error Resume Next
Set app = CreateObject("Word.Application")
If Err.Number <> 0 Then
    Err.Clear
    Set app = CreateObject("KWPS.Application")
End If
If Err.Number <> 0 Then
    Err.Clear
    Set app = CreateObject("WPS.Application")
End If
If Err.Number <> 0 Then
    WScript.Echo "NO_APP"
    WScript.Quit 1
End If
app.Visible = False
app.DisplayAlerts = 0
Set doc = app.Documents.Open(WScript.Arguments(0), False, True)
If Err.Number = 0 Then
    On Error Resume Next
    If ff = "" Then
        doc.SaveAs2 dst
        If Err.Number <> 0 Then
            Err.Clear
            doc.SaveAs dst
        End If
    Else
        doc.SaveAs2 dst, CInt(ff)
        If Err.Number <> 0 Then
            Err.Clear
            doc.SaveAs dst, CInt(ff)
        End If
    End If
    If Err.Number = 0 Then
        ok = True
    End If
    doc.Close False
End If
app.Quit
If ok Then
    WScript.Echo "OK"
Else
    WScript.Echo "FAIL"
End If
'''


def _is_valid_docx(path):
    """判断文件是否为合法 .docx：ZIP 包且含 word/document.xml。"""
    try:
        with zipfile.ZipFile(path) as z:
            return "word/document.xml" in z.namelist()
    except Exception:
        return False


def _convert_doc_via_ms_app(path, workdir):
    """用本机已装的 Word / WPS 把 .doc/.wps 转成 .docx（Windows 专用）。

    依次尝试多种保存策略，每次都在 Python 端校验产出的 .docx 是否合法，
    直到得到含 word/document.xml 的合法文件。全部失败返回 None。
    """
    if not sys.platform.startswith("win"):
        return None
    out_dir = tempfile.mkdtemp(prefix="tfd_conv_", dir=workdir or None)
    dst = os.path.join(out_dir, "_converted.docx")
    vbs = os.path.join(out_dir, "_convert.vbs")
    # 脚本宿主需 UTF-16（带 BOM）才能正确读含中文的脚本与路径
    with open(vbs, "w", encoding="utf-16", newline="\r\n") as f:
        f.write(_VBS_CONVERT)
    strategies = ["", "12", "16"]   # 扩展名驱动 → Word2007 docx → 默认 docx
    for ff in strategies:
        if os.path.exists(dst):
            try:
                os.remove(dst)
            except OSError:
                pass
        cmd = ["cscript.exe", "//nologo", vbs, path, dst]
        if ff:
            cmd.append(ff)
        try:
            r = subprocess.run(cmd, capture_output=True, timeout=240)
        except Exception as e:
            raise RuntimeError("调用本机 Word/WPS 转换失败：" + str(e))
        out = r.stdout.decode("utf-8", errors="replace")
        if "OK" in out and os.path.isfile(dst) and _is_valid_docx(dst):
            return dst
    return None


def normalize_input(path):
    """把输入文件整理成引擎能吃的标准 .docx。

    返回 (docx_path, note)。
      - 已是 .docx：原样返回。
      - .doc / .wps 或 OLE2 二进制：尝试用 LibreOffice 转成临时 .docx；
        找不到 LibreOffice 则抛 RuntimeError 提示用户先另存为 .docx。
    注意：绝不覆盖用户原始文件。
    """
    if not os.path.isfile(path):
        raise FileNotFoundError("文件不存在：" + path)
    ext = os.path.splitext(path)[1].lower()
    if ext == ".docx":
        return path, None
    if ext in (".doc", ".wps") or _looks_like_ole2(path):
        workdir = os.path.dirname(os.path.abspath(path))
        # 优先级 1：LibreOffice（跨平台）
        soffice = find_soffice()
        if soffice:
            conv = _convert_doc_to_docx(path, soffice, workdir)
            return conv, "已自动将旧格式转换为 .docx 后处理（LibreOffice，原文件未改动）"
        # 优先级 2：本机已装的 Word / WPS（Windows COM）
        conv = _convert_doc_via_ms_app(path, workdir)
        if conv:
            return conv, "已自动将旧格式转换为 .docx 后处理（本机 Word/WPS，原文件未改动）"
        raise RuntimeError(
            "检测到旧格式文件（.doc / .wps），自动转换未成功。\n"
            "已尝试：LibreOffice（未安装）、本机 Word/WPS（转换结果异常，产物不是合法 .docx）。\n"
            "解决办法（任选其一）：\n"
            "① 安装免费的 LibreOffice（推荐，装好后本软件会自动转换）；\n"
            "② 用 Word 或 WPS 打开该文件，手动执行「另存为 → Word 文档(.docx)」后再用本软件处理。"
        )
    # 其它扩展名：交给引擎，让它报“无法识别”之类的错
    return path, None


# ---------------------------------------------------------------------------
# 对外 API：四种处理模式
# ---------------------------------------------------------------------------
def run_check(path, profile_path=None, out_md=None):
    """格式体检（通用或模板驱动）。返回 markdown 报告文本。

    path: 已转为 docx 的文件（调用方先 normalize_input）。
    profile_path: 学校模板画像 .json（可选，提供则走模板驱动诊断）。
    out_md: 报告保存路径（可选）。

    注意：不向引擎传 -o。format_checker 在传 -o 时只写文件、不打印正文，
    会导致 GUI 日志看不到报告。改为捕获完整 markdown 后由本函数写文件。
    """
    argv = ["format_checker.py", path]
    if profile_path:
        argv += ["--profile", profile_path]
    report = _run_module_main(format_checker, argv)
    if out_md:
        with open(out_md, "w", encoding="utf-8") as f:
            f.write(report)
    return report


def run_fix_headings(src, dst, profile_path=None, report_docx=None, add_comments=True):
    """一键套标题样式（通用或模板驱动）。

    src: 输入 docx；dst: 输出 docx（新文件，不改动 src）；
    report_docx: 修改明细 docx（可选）；add_comments: 是否在文档写批注。
    返回 markdown 修改清单文本。
    """
    argv = ["headings-fix.py", src, dst]
    if profile_path:
        argv += ["--profile", profile_path]
    if report_docx:
        argv += ["--report", report_docx]
    if not add_comments:
        argv += ["--no-comments"]
    return _run_module_main(headings_fix, argv)


def run_reformat_refs(src, dst):
    """参考文献按 GB/T 7714 重排。src 输入 docx；dst 输出 docx。返回报告文本。"""
    argv = ["ref-reformat.py", src, "-o", dst]
    return _run_module_main(ref_reformat, argv)


def run_build_profile(template_path, out_json):
    """从学校模板 .docx 生成格式画像 profile.json。返回画像 JSON 文本。

    不向引擎传 -o：format_profile 传 -o 时只写文件、不打印正文，
    会让 GUI 日志看不到画像内容；改为捕获完整 JSON 后由本函数写文件。
    """
    argv = ["format-profile.py", template_path]
    text = _run_module_main(format_profile, argv)
    with open(out_json, "w", encoding="utf-8") as f:
        f.write(text)
    return text


# 方便 GUI 做模式枚举
MODES = {
    "check": "格式体检（不修改，出报告）",
    "fix": "一键套标题样式（生成新文档）",
    "ref": "参考文献重排（生成新文档）",
    "profile": "从模板生成格式画像(.json)",
}

__all__ = [
    "normalize_input",
    "run_check",
    "run_fix_headings",
    "run_reformat_refs",
    "run_build_profile",
    "MODES",
]
