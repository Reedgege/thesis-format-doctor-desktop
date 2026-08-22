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
import time
import zipfile
import traceback

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


def _convert_doc_via_com_inproc(path, out_dir):
    """在软件进程内直接调用本机 Word / WPS（pywin32 COM）把 .doc/.wps 转 .docx。

    优点：没有外部进程、不弹任何窗口（本机办公软件也以隐藏方式运行）。
    返回生成的 .docx 路径；未装 pywin32、本机无 Word/WPS 或转换失败时返回 None。

    注意：用户开着 WPS/Word 时 COM 调用偶发失败（如 <unknown>.Add / 发生意外），
    因此整体最多重试 3 轮，每轮都用全新实例（早绑定优先）。
    """
    try:
        import win32com.client
        import pythoncom
    except Exception:
        return None
    dst = os.path.join(out_dir, "_converted.docx")
    progids = ("Word.Application", "KWPS.Application", "WPS.Application")
    for attempt in range(3):
        app = None
        try:
            pythoncom.CoInitialize()
            for progid in progids:
                try:
                    app = win32com.client.gencache.EnsureDispatch(progid)
                    break
                except Exception:
                    app = None
                try:
                    app = win32com.client.DispatchEx(progid)
                    break
                except Exception:
                    app = None
                try:
                    app = win32com.client.Dispatch(progid)
                    break
                except Exception:
                    app = None
            if app is None:
                return None
            try:
                app.Visible = False
            except Exception:
                pass
            try:
                app.DisplayAlerts = 0
            except Exception:
                pass
            for strategy in (None, 12, 16):   # 扩展名驱动 → Word2007 docx → 默认 docx
                if os.path.exists(dst):
                    try:
                        os.remove(dst)
                    except OSError:
                        pass
                doc = None
                try:
                    doc = app.Documents.Open(path, False, True)
                    try:
                        if strategy is None:
                            doc.SaveAs2(dst)
                        else:
                            doc.SaveAs2(dst, strategy)
                    except Exception:
                        if strategy is None:
                            doc.SaveAs(dst)
                        else:
                            doc.SaveAs(dst, strategy)
                    doc.Close(False)
                    doc = None
                except Exception:
                    try:
                        if doc is not None:
                            doc.Close(False)
                    except Exception:
                        pass
                if _is_valid_docx(dst):
                    return dst
        except Exception:
            pass
        finally:
            try:
                if app is not None:
                    app.Quit()
            except Exception:
                pass
        if attempt < 2:
            time.sleep(1.0)
    return None


def _convert_doc_via_script_host(path, out_dir):
    """经系统脚本宿主（隐藏控制台窗口）调用本机 Word / WPS 转换。

    备用通道：仅当进程内 COM 不可用时才走这里。返回 docx 路径或 None。
    """
    dst = os.path.join(out_dir, "_converted.docx")
    vbs = os.path.join(out_dir, "_convert.vbs")
    # 脚本宿主需 UTF-16（带 BOM）才能正确读含中文的脚本与路径
    with open(vbs, "w", encoding="utf-16", newline="\r\n") as f:
        f.write(_VBS_CONVERT)
    kwargs = {}
    if sys.platform.startswith("win"):
        # 关键：不让子进程弹出黑色控制台窗口（cscript 默认会闪一下黑框）
        kwargs["creationflags"] = getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000)
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
            r = subprocess.run(cmd, capture_output=True, timeout=240, **kwargs)
        except Exception as e:
            raise RuntimeError("调用本机 Word/WPS 转换失败：" + str(e))
        out = r.stdout.decode("utf-8", errors="replace")
        if "OK" in out and os.path.isfile(dst) and _is_valid_docx(dst):
            return dst
    return None


def _convert_doc_via_ms_app(path, workdir):
    """用本机已装的 Word / WPS 把 .doc/.wps 转成 .docx（Windows 专用）。

    优先在软件进程内直接调用（pywin32 COM，无任何弹窗、无外部进程）；
    若进程内调用不可用，退回脚本宿主通道并强制隐藏控制台窗口。
    返回 (docx_path, note)；全部失败返回 (None, None)。
    """
    if not sys.platform.startswith("win"):
        return None, None
    out_dir = tempfile.mkdtemp(prefix="tfd_conv_", dir=workdir or None)
    conv = _convert_doc_via_com_inproc(path, out_dir)
    if conv:
        return conv, "已自动将旧格式转换为 .docx 后处理（软件内置转换，原文件未改动）"
    conv = _convert_doc_via_script_host(path, out_dir)
    if conv:
        return conv, "已自动将旧格式转换为 .docx 后处理（备用通道转换，原文件未改动）"
    return None, None


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
        conv, cnote = _convert_doc_via_ms_app(path, workdir)
        if conv:
            return conv, cnote
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


def md_to_docx(md_text, out_docx):
    """把 Markdown 报告文本转成 Word 文档（纯标准库生成，无需 python-docx）。

    关键：本软件运行环境（PyInstaller 单文件 exe）未安装 python-docx，若依赖它，
    检查报告 / 修正检查报告在客户机器上会 ImportError 写不出文件。
    这里直接按 OOXML 规范用 zipfile 生成最小可用的 .docx（Word/WPS 均可正常打开）。

    支持：# / ## / ### 标题、- / * 列表、普通段落；`**加粗**`、行内代码标记去除。
    """
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    def esc(t):
        return (t.replace("&", "&amp;").replace("<", "&lt;")
                 .replace(">", "&gt;").replace('"', "&quot;"))

    body = []
    for raw in md_text.splitlines():
        line = raw.rstrip()
        if not line.strip():
            continue
        s = line.strip()
        if s.startswith("### "):
            ppr, text = '<w:pPr><w:pStyle w:val="Heading3"/></w:pPr>', s[4:]
        elif s.startswith("## "):
            ppr, text = '<w:pPr><w:pStyle w:val="Heading2"/></w:pPr>', s[3:]
        elif s.startswith("# "):
            ppr, text = '<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>', s[2:]
        else:
            ppr, text = "", s
        text = text.replace("**", "").replace("`", "")
        if s.startswith("- ") or s.startswith("* "):
            text = "• " + text
        run = '<w:r><w:t xml:space="preserve">%s</w:t></w:r>' % esc(text)
        body.append('<w:p>%s%s</w:p>' % (ppr, run))

    document = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:document xmlns:w="%s"><w:body>%s'
        '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
        'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
        '</w:body></w:document>' % (W, "".join(body))
    )
    styles = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:styles xmlns:w="%s">'
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
        '<w:name w:val="Normal"/><w:rPr><w:rFonts w:ascii="Calibri" '
        'w:eastAsia="宋体"/><w:sz w:val="21"/></w:rPr></w:style>'
        '<w:style w:type="paragraph" w:styleId="Heading1">'
        '<w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>'
        '<w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>'
        '<w:style w:type="paragraph" w:styleId="Heading2">'
        '<w:name w:val="heading 2"/><w:basedOn w:val="Normal"/>'
        '<w:rPr><w:b/><w:sz w:val="26"/></w:rPr></w:style>'
        '<w:style w:type="paragraph" w:styleId="Heading3">'
        '<w:name w:val="heading 3"/><w:basedOn w:val="Normal"/>'
        '<w:rPr><w:b/><w:sz w:val="22"/></w:rPr></w:style>'
        '</w:styles>' % W
    )
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" '
        'ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.'
        'wordprocessingml.document.main+xml"/>'
        '<Override PartName="/word/styles.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.'
        'wordprocessingml.styles+xml"/>'
        '</Types>'
    )
    rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships '
        'xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/'
        'officeDocument/2006/relationships/officeDocument" '
        'Target="word/document.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/'
        'officeDocument/2006/relationships/styles" Target="word/styles.xml"/>'
        '</Relationships>'
    )
    if os.path.exists(out_docx):
        try:
            os.remove(out_docx)
        except OSError:
            pass
    with zipfile.ZipFile(out_docx, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", rels)
        z.writestr("word/document.xml", document)
        z.writestr("word/styles.xml", styles)


def _log_fix_error(src, label, exc):
    """把一键修正的真实异常写入临时日志（便于回传定位），同时打印到 stderr。"""
    try:
        log_path = os.path.join(tempfile.gettempdir(), "tfd_fix_error.log")
        with open(log_path, "a", encoding="utf-8") as f:
            f.write("=" * 60 + "\n")
            f.write("时间: %s\n" % time.strftime("%Y-%m-%d %H:%M:%S"))
            f.write("输入: %s\n" % src)
            f.write("尝试模式: %s\n" % label)
            f.write("异常: %s\n" % repr(exc))
            f.write(traceback.format_exc())
            f.write("\n")
        sys.stderr.write("[一键修正异常·已记录到 %s]\n%s\n" % (log_path, traceback.format_exc()))
    except Exception:
        pass


def run_fix_headings(src, dst, profile_path=None, report_docx=None, add_comments=True):
    """一键套标题样式（通用或模板驱动）。

    src: 输入 docx；dst: 输出 docx（新文件，不改动 src）；
    report_docx: 修改明细 docx（可选）；add_comments: 是否在文档写批注。
    返回 markdown 修改清单文本。

    模板驱动原则（v1.3.42）：本软件是模板驱动的修正工具——客户选择了学校模板后，
    修正必须严格按【从模板提取的画像】执行。因此当 profile_path 非空时，
    **绝不静默退化为通用规范**：带模板的两次尝试（模板+修改报告 → 模板无报告）
    全部失败则直接报错（GUI 弹窗 + %TEMP%/tfd_fix_error.log 留痕），由客户/客服
    定位，而不是输出一份"没按模板改"的论文让客户困惑。
    未选模板（profile_path 为空）时，base 本身即通用规范——那是客户预期的行为，
    不属于降级。
    """
    base = ["headings-fix.py", src, dst]
    if profile_path:
        base += ["--profile", profile_path]
    if not add_comments:
        base += ["--no-comments"]
    attempts = []
    if report_docx:
        attempts.append((base + ["--report", report_docx], "模板+修改报告"))
    attempts.append((base, "模板（无修改报告）"))
    last = None
    for argv, label in attempts:
        try:
            return _run_module_main(headings_fix, argv)
        except Exception as e:
            last = e
            _log_fix_error(src, label, e)
            continue
    raise last


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
