# -*- coding: utf-8 -*-
"""
论文格式医生 · 原生 Tkinter 界面（客户安心版 v2）
==============================================
面向客户的简洁流程，不展示任何执行细节；一屏放得下，无需滚动：

  顶部品牌区  —— 线描 mark + 大标题「论文格式医生」+ 朱砂竖章「学生」+ 副标题
  左栏        —— 文件选择：上传区 → 学校模板 → 模板说明 → 单选状态 → 标语
  右栏        —— 处理步骤：步进器 → ① 提取学校模板要求 ② 论文格式检查 ③ 一键格式修正
                 → 主按钮 → 底部信任卡
  页脚        —— 邮箱 ｜ 标语 ｜ 官网 + 客服入口 + 本机处理声明 + 版权

界面文字唯一来源：`_ui_ref/COPY_TABLE.md`（逐字照抄，不许改写）。

交互约定：
  - 「提取学校模板要求」完成后，弹出“学校模板要求”确认页（关键项可修改），
    客户确认后才生效；放弃则本次不使用画像。
  - 确认页字段先用“批注要求”，没有批注则回退到“样式定义”填充（字体/字号/
    行距/缩进/页边距/对齐等），保证提取结果不空、可核对可修改。
  - 「按学校要求一键修正」导出地址由客户选择；完成后弹确认页，
    并附检查报告 + 修改报告。
  - .doc / WPS .wps 一律先自动转 .docx 再处理（进程内转换，无黑框、不弹窗），
    原文件绝不被修改。
"""
import os
import sys
import json
import time
import queue
import tempfile
import hashlib
import threading
import subprocess
import shutil
import webbrowser

# 把 tfd_app 加入搜索路径（打包成 exe 后 sys.path 已含，但开发态下保险）
HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

from . import trial       # v1.3.56：试用计数（机器码绑定，1 次）——相对导入，PyInstaller 才收集
from . import watermark   # v1.3.56：试用水印（页眉页脚+正文穿插）
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import tkinter.font as tkfont

from . import engine, license
from .ui import get_theme, RADIUS, TYPE
from .ui import widgets, brand, backdrop
from .ui.modal import ModalShell, ConfirmModal


ICON = os.path.join(HERE, "assets", "icon.png")
QRCODE = os.path.join(HERE, "assets", "qrcode.png")
MINIAPP_QRCODE = os.path.join(HERE, "assets", "miniapp_qrcode.png")
MINIAPP_NAME = "芦苇论文格式"
# 品牌 / 客服（文案统一来源，避免散落硬编码）
WECHAT_NAME = "芦苇不熬夜"
WECHAT_ID = "reedskill"
ABOUT_MAIL = "hi@reedskill.com"
HELP_HINT = "\n\n遇到问题？看帮助或关注公众号【%s】留言。" % WECHAT_NAME
OFFICIAL_SITE = "https://reedskill.com"        # 官网（大本营）；给人看的省略 https，代码里打开用全址
OFFICIAL_SITE_TEXT = "官网：reedskill.com"

# ---------------------------------------------------------------------------
# 界面文案常量 —— 唯一来源 `_ui_ref/COPY_TABLE.md`（装修包实物逐字提取）
#
# 铁律：以下每个字符串都与 COPY_TABLE **一字不差**（含标点、空格、`·`、`→`、`｜`、
# 括号全半角）。改文案必须同步改 COPY_TABLE，不许在这里「顺手优化措辞」。
# ---------------------------------------------------------------------------
COPY_TITLE = "论文格式医生"                 # 顶部大标题（旧版字符间加空格的写法一律作废）
COPY_SUBTITLE = "让论文格式更简单"          # 顶部副标题
COPY_UPGRADE = "升级正式版 →"               # 右上实心按钮
# 右上入口行原文（顺序固定）；LinkRow 用「 ｜ 」原样分词，渲染与装修包一致，
# 也保证权威字符串在代码里逐字可查（COPY_TABLE §1）。
COPY_ENTRY_LINE = "客服微信 ｜ 官方公众号 ｜ 关于 ｜ 帮助"
COPY_ENTRY_LINKS = tuple(COPY_ENTRY_LINE.split(" ｜ "))
COPY_TRIAL_FMT = "试用版 · 剩余 %d 次"      # 试用徽章（X = 剩余次数）
COPY_TRIAL_USED = "试用版 · 已用完"
# 正式徽章（COPY_TABLE §5：正式版 · 月卡）——按卡种给出成品短文案
COPY_FORMAL_LABELS = {
    "次卡": "正式版 · 次卡",
    "周卡": "正式版 · 周卡",
    "月卡": "正式版 · 月卡",
    "永久版": "正式版 · 永久版",
}
COPY_FORMAL_FMT = "正式版 · %s"             # 未登记卡种兜底（行为同旧实现）
COPY_FORMAL_FALLBACK = "正式版"
# 右上入口行整体渲染为：客服微信 ｜ 官方公众号 ｜ 关于 ｜ 帮助
# 试用 / 正式徽章形如：试用版 · 剩余 X 次 ／ 正式版 · 月卡

COPY_LEFT_TITLE = "文件选择"
COPY_LEFT_HINT = "支持多种文档格式，智能识别格式要求"
COPY_TPL_TITLE = "学校模板"
COPY_TPL_BADGE = "可选"
COPY_TPL_DESC = "用于按学校要求检查／修正，更贴合要求"
COPY_TPL_BTN = "选择模板"
COPY_NOTE_TITLE = "模板说明"
COPY_NOTE_BODY = ("请优先使用学校官方模板（通常批注中写明了格式要求）；若模板无批注，"
                  "将按模板格式定义／通用规范处理，可能与学校要求有出入。")
COPY_PICK_THESIS = "未选择论文"
COPY_PICK_TPL = "模板未选（可选）"
COPY_LEFT_SLOGAN = "一键规范格式，专注论文内容"

COPY_RIGHT_TITLE = "处理步骤"
COPY_SLANT_LINES = ("三步完成", "论文格式检查与修正")
COPY_COUNTER_FMT = "%d / %d"
COPY_COUNTER_FIRST = "1 / 3"                # 右上页码初值（视觉稿原文）
# 右上页码形如：1 / 3
COPY_CTA_FIRST = "下一步：开始检查 →"
COPY_CTA_FIX = "下一步：一键修正 →"
COPY_TRUST_TITLE = "学术规范 · 专业高效"
COPY_TRUST_BODY = "精准识别格式问题，助力您的论文顺利通过审核。"

# 状态与提示（COPY_TABLE §5）
COPY_STATUS_RUNNING = "正在处理，请稍候…"
COPY_STATUS_DONE = "处理完成"
COPY_STATUS_FAIL = "处理失败，请检查文件"
COPY_TOAST_SAVED = "已保存 · 你的设置已更新"

# ---------------------------------------------------------------------------
# 配色：全部走 ui.theme 的 student 调色板（与导师版 _ADV 同一套机制，只换 edition）
# ---------------------------------------------------------------------------
_T = get_theme("student")
PAPER   = _T.bg          # 宣纸底
INK     = _T.navy        # 主文字（深）—— 标题 / 重点
BODY    = "#444444"      # 正文（灰）—— 描述性文字
MUTED   = _T.muted       # 次要文字 / 页脚（浅灰）
LINE    = _T.border      # 细线
ACCENT  = _T.primary     # 学术蓝（强调）
OKC     = _T.success_dot # 完成（墨绿点）
ERRC    = _T.error_text  # 出错（朱红）

# 字体：主界面一律走 ui.theme 的 serif / sans 令牌（与装修包一致）；下面这几个
# F_XXX 命名对象只留给「关于 / 使用帮助」等 tk 原生窗口，运行时由 _init_fonts()
# 创建。字号**只从 TYPE 表取名**（唯一权威表，见 _ui_ref/TYPE_SCALE.md）——
# 这里没有第二套字号，也没有运行期缩放。
from .buildinfo import APP_VERSION   # 版本号唯一来源（与 VERSION 文件的一致性由 tests 钉住）

# v1.3.108：绿色 zip 版由软件自建桌面快捷方式（win32com 已内置，客户零依赖、零黑框）。
APP_SHORTCUT_NAME = "论文格式医生"       # 桌面快捷方式显示名
SHORTCUT_FLAG_TAG = "ThesisFormatDoctor"  # %APPDATA% 下询问标记目录（区分学生/导师版）


def _is_frozen_exe():
    """打包后的主程序才显示“桌面图标”入口 / 首启询问；
    开发态（python.exe/pythonw.exe）与 mac/Linux 一律不建 Windows 快捷方式。"""
    base = os.path.basename(sys.executable or "").lower()
    return base.endswith(".exe") and not (base.startswith("python")
                                          or base.startswith("pythonw"))


_FONTS = {}      # name -> (Font, size)

def _init_fonts(root):
    """在 root 创建后调用：把 F_XXX 全局名绑定为 TYPE 字号的 Font 命名对象。

    角色表写的是 TYPE 的**键名**（不是磅值字面量）——字号只有 theme.TYPE 一处定义。
    """
    import platform as _platform
    roles = {
        "F_TITLE":    ("KaiTi", "page_title", "bold"),      # 主标题（楷体）
        "F_HDR":      ("KaiTi", "section_title", "bold"),   # 章节标题（楷体）
        "F_BODY":     ("Microsoft YaHei", "body"),          # 正文 / 步骤说明
        "F_SMALL":    ("Microsoft YaHei", "caption"),       # 底部提示 / 次要
        "F_SUBTITLE": ("Microsoft YaHei", "body"),          # 元信息
        "F_FOOT":     ("Microsoft YaHei", "footer"),        # 页脚（最小层级）
    }
    # 跨平台中文字体兜底：Windows 自带楷体/雅黑；mac/Linux 无则映射到系统 CJK 字体，
    # 避免 tkinter 静默回退到无中文的默认字体出现方块（tofu）。Windows 上字体存在，走原值。
    _fams = set(tkfont.families())
    _KAITI = {"Darwin": "STKaiti", "Linux": "Noto Serif CJK SC"}
    _HEITI = {"Darwin": "PingFang SC", "Linux": "Noto Sans CJK SC"}
    def _resolve(fam):
        if fam in _fams:
            return fam
        cand = (_KAITI if fam == "KaiTi" else _HEITI).get(_platform.system())
        if cand and cand in _fams:
            return cand
        return fam
    for name, spec in roles.items():
        size = TYPE[spec[1]]
        kw = {"family": _resolve(spec[0]), "size": size}
        if len(spec) > 2:
            kw["weight"] = spec[2]
        f = tkfont.Font(root=root, **kw)
        _FONTS[name] = (f, size)
        globals()[name] = f

# 页边距字段：编辑厘米值时同步写 twips（top/bottom/left/right），兼容两套读取方
_MARGIN_TWIPS = {
    ("spec", "page", "top_cm"): "top",
    ("spec", "page", "bottom_cm"): "bottom",
    ("spec", "page", "left_cm"): "left",
    ("spec", "page", "right_cm"): "right",
}
_TWIPS_PER_CM = 567.0   # 1 厘米 ≈ 567 twips

ALIGN_DISPLAY = {"center": "居中", "left": "左对齐", "right": "右对齐", "both": "两端对齐"}
ALIGN_CODE = {v: k for k, v in ALIGN_DISPLAY.items()}

# 确认页可修改项：(中文标签, [候选取值路径…], 控件类型)
# 取值优先级：批注要求(spec.*) → 样式定义(profile 根的 body/headings/headingStyles/page)
# 字段取值路径：levels 段最完整（样式定义挖全要素 + 批注覆盖），spec 兜底，根级再兜底。
# levels 键：zh_font/sz(半磅)/align/line_val/indent_chars/before_pt/after_pt/bold 等。
EDIT_FIELDS = [
    ("正文字体",     [("levels", "body", "zh_font"), ("spec", "body", "zh_font"),
                     ("spec", "body", "font"), ("body", "font")], "text"),
    # v1.3.49：字号候选把 sz（半磅数字）提到 size（中文字号名）之前——
    # 此前 size 优先导致表单显示"小四"、客户改数字写回 size 字段而引擎只读 sz，
    # 客户改字号不生效（自查发现的逻辑 bug）。
    ("正文字号",     [("levels", "body", "sz"), ("levels", "body", "size"),
                     ("spec", "body", "sz"), ("spec", "body", "size")], "text"),
    ("正文行距(磅)", [("levels", "body", "line_val"), ("spec", "body", "line_val"),
                     ("body", "line_val"), ("body", "line")], "text"),
    ("首行缩进(字符)", [("levels", "body", "indent_chars"), ("spec", "body", "indent_chars"),
                       ("body", "indent_chars"), ("body", "firstLineChars")], "text"),
    ("一级标题字体", [("levels", "1", "zh_font"), ("spec", "h1", "zh_font"),
                     ("spec", "h1", "font"), ("headings", "1", "font")], "text"),
    ("一级标题字号", [("levels", "1", "sz"), ("levels", "1", "size"),
                     ("spec", "h1", "sz"), ("spec", "h1", "size")], "text"),
    ("一级标题对齐", [("levels", "1", "align"), ("spec", "h1", "align")], "align"),
    ("二级标题字体", [("levels", "2", "zh_font"), ("spec", "h2", "zh_font"),
                     ("spec", "h2", "font"), ("headings", "2", "font")], "text"),
    ("二级标题字号", [("levels", "2", "sz"), ("levels", "2", "size"),
                     ("spec", "h2", "sz"), ("spec", "h2", "size")], "text"),
    ("三级标题字体", [("levels", "3", "zh_font"), ("spec", "h3", "zh_font"),
                     ("spec", "h3", "font"), ("headings", "3", "font")], "text"),
    ("三级标题字号", [("levels", "3", "sz"), ("levels", "3", "size"),
                     ("spec", "h3", "sz"), ("spec", "h3", "size")], "text"),
    ("页边距 上(厘米)", [("spec", "page", "top_cm"), ("page", "top")], "text"),
    ("页边距 下(厘米)", [("spec", "page", "bottom_cm"), ("page", "bottom")], "text"),
    ("页边距 左(厘米)", [("spec", "page", "left_cm"), ("page", "left")], "text"),
    ("页边距 右(厘米)", [("spec", "page", "right_cm"), ("page", "right")], "text"),
    ("参考文献格式", [("levels", "reference"), ("spec", "reference"),
                     ("refExample",), ("reference",)], "ref"),
    # v1.3.78：结构页（摘要/致谢/附录）确认项——模板画像已拆「标题/正文」双 spec，
    # 弹窗中可逐项核对与修改；写回 levels.xxx_title / levels.xxx 供引擎套用。
    ("摘要标题字体", [("levels", "abstract_title", "zh_font"),
                     ("spec", "abstract_title", "zh_font"),
                     ("levels", "abstract", "zh_font")], "text"),
    ("摘要标题字号", [("levels", "abstract_title", "sz"), ("levels", "abstract_title", "size"),
                     ("spec", "abstract_title", "sz"), ("levels", "abstract", "sz")], "text"),
    ("摘要正文字体", [("levels", "abstract", "zh_font"), ("spec", "abstract", "zh_font")], "text"),
    ("摘要正文字号", [("levels", "abstract", "sz"), ("levels", "abstract", "size"),
                     ("spec", "abstract", "sz")], "text"),
    ("摘要正文行距(磅)", [("levels", "abstract", "line_val"), ("spec", "abstract", "line_val")], "text"),
    ("致谢标题字体", [("levels", "ack_title", "zh_font"), ("spec", "ack_title", "zh_font"),
                    ("levels", "ack", "zh_font")], "text"),
    ("致谢标题字号", [("levels", "ack_title", "sz"), ("levels", "ack_title", "size"),
                    ("spec", "ack_title", "sz"), ("levels", "ack", "sz")], "text"),
    ("致谢正文字体", [("levels", "ack", "zh_font"), ("spec", "ack", "zh_font")], "text"),
    ("附录标题字体", [("levels", "appendix_title", "zh_font"), ("spec", "appendix_title", "zh_font"),
                    ("levels", "appendix", "zh_font")], "text"),
    ("附录标题字号", [("levels", "appendix_title", "sz"), ("levels", "appendix_title", "size"),
                    ("spec", "appendix_title", "sz"), ("levels", "appendix", "sz")], "text"),
    # v1.3.79：附录正文一律不自动修改（各校差异极大），故不提供"附录正文字体"确认项，
    # 避免客户修改后不生效造成误导；附录标题处会以批注说明正文请自行对照学校要求处理。
]


def _base_no_ext(path):
    return os.path.splitext(path)[0]


def _deep_get(d, path):
    for k in path:
        if not isinstance(d, dict) or k not in d:
            return None
        d = d[k]
    return d


def _deep_set(d, path, value):
    for k in path[:-1]:
        if not isinstance(d.get(k), dict):
            d[k] = {}
        d = d[k]
    d[path[-1]] = value


_CN_SIZE_PT = {  # 中文字号名 → 磅（通用印刷规范，非学校特定值）
    "初号": 42, "小初": 36, "一号": 26, "小一": 24, "二号": 22, "小二": 18,
    "三号": 16, "小三": 15, "四号": 14, "小四": 12, "五号": 10.5, "小五": 9,
    "六号": 7.5, "小六": 6.5, "七号": 5.5, "八号": 5,
}


def _cn_size_to_pt(text):
    """把客户输入的字号转成磅：支持数字（如 12）与中文字号名（如 小四=12）。

    v1.3.49：确认弹窗字号输入此前只认数字，客户填"小四"会被当非法忽略。
    """
    s = (text or "").strip()
    if not s:
        return None
    try:
        return float(s)
    except (TypeError, ValueError):
        pass
    if s in _CN_SIZE_PT:
        return _CN_SIZE_PT[s]
    # 容忍"小四号""小五号"等写法：先匹配"小X"（小四=12 优先于 四号=14），再匹配普通字号名
    for name in ("小初", "小一", "小二", "小三", "小四", "小五", "小六", "小七"):
        if name in s:
            return _CN_SIZE_PT[name]
    for name, pt in _CN_SIZE_PT.items():
        if name in s:
            return pt
    return None


def _field_value(profile, candidates):
    """取确认页字段值：批注要求优先，样式定义兜底；含单位换算。"""
    for path in candidates:
        v = _deep_get(profile, path)
        if v in (None, ""):
            continue
        s = str(v)
        # levels/spec 里的 sz 是半磅（OOXML 单位）→ 显示为磅
        if path[-1] == "sz":
            try:
                return str(int(float(v)) // 2)
            except (ValueError, TypeError):
                return s
        # 样式定义里的页边距是 twips → 换算成厘米显示
        if len(path) == 2 and path[0] == "page" and path[1] in ("top", "bottom", "left", "right"):
            try:
                return "%.2f" % (float(v) / _TWIPS_PER_CM)
            except (ValueError, TypeError):
                return s
        # 样式定义里的正文行距是 twips → 换算成磅显示
        if tuple(path) == ("body", "line"):
            try:
                return str(int(float(v) / 20.0))
            except (ValueError, TypeError):
                return s
        return s
    return ""


def _profile_summary(profile):
    """把格式画像渲染成客户看得懂的“学校模板要求”文本行。

    批注要求(spec.*)优先，样式定义(profile 根)兜底，保证提取结果不空。
    """
    lines = []
    spec = profile.get("spec") or {}
    levels = profile.get("levels") or {}

    # 页面
    page = spec.get("page") or {}
    if not page.get("top_cm") and not page.get("top"):
        page = profile.get("page") or {}
    if page.get("top_cm") is not None:
        lines.append("· 页边距：上 %s 下 %s 左 %s 右 %s（厘米）" % (
            page.get("top_cm"), page.get("bottom_cm"),
            page.get("left_cm"), page.get("right_cm")))
    elif page.get("top") is not None:
        try:
            t = "%.2f" % (float(page["top"]) / _TWIPS_PER_CM)
            b = "%.2f" % (float(page["bottom"]) / _TWIPS_PER_CM)
            l = "%.2f" % (float(page["left"]) / _TWIPS_PER_CM)
            r = "%.2f" % (float(page["right"]) / _TWIPS_PER_CM)
            lines.append("· 页边距：上 %s 下 %s 左 %s 右 %s（厘米）" % (t, b, l, r))
        except (ValueError, TypeError):
            lines.append("· 页边距：已提取")

    # 正文（levels 完整画像优先，spec/根级兜底）
    body = levels.get("body") or spec.get("body") or profile.get("body") or {}
    if body:
        parts = []
        font = body.get("zh_font") or body.get("font")
        if font:
            parts.append("字体 %s" % font)
        size = body.get("size")
        if size:
            parts.append("字号 %s" % size)
        elif body.get("sz"):
            try:
                parts.append("字号 %s" % str(int(float(body["sz"])) // 2))
            except (ValueError, TypeError):
                pass
        indent = body.get("indent_chars") or body.get("firstLineChars")
        if indent:
            parts.append("首行缩进 %s 字符" % indent)
        line_val = body.get("line_val")
        if line_val is None and body.get("line"):
            try:
                line_val = int(float(body["line"]) / 20.0)
            except (ValueError, TypeError):
                line_val = None
        if line_val:
            parts.append("行距 %s 磅" % line_val)
        lines.append("· 正文：%s" % ("，".join(parts) if parts else "样式已提取"))

    # 各级标题（levels 优先，spec/headings 兜底）
    lv_keys = {"h1": "1", "h2": "2", "h3": "3"}
    for lv, name in (("h1", "一级标题"), ("h2", "二级标题"), ("h3", "三级标题")):
        h = (levels.get(lv_keys[lv]) or spec.get(lv)
             or profile.get("headings", {}).get(lv_keys[lv]) or {})
        if not h:
            continue
        parts = []
        font = h.get("zh_font") or h.get("font")
        if font:
            parts.append("字体 %s" % font)
        size = h.get("size")
        if size:
            parts.append("字号 %s" % size)
        elif h.get("sz"):
            try:
                parts.append("字号 %s" % str(int(float(h["sz"])) // 2))
            except (ValueError, TypeError):
                pass
        if h.get("align"):
            parts.append("对齐 %s" % ALIGN_DISPLAY.get(h["align"], h["align"]))
        if h.get("bold"):
            parts.append("加粗")
        lines.append("· %s：%s" % (name, "，".join(parts) if parts else "样式已提取"))

    # 其它分类（levels 优先，spec 兜底；dict 或文本都展示）
    for key, label in (("abstract", "摘要"), ("keywords", "关键词"), ("toc", "目录"),
                       ("reference", "参考文献"), ("title", "论文题目"),
                       ("table", "表格"), ("figure", "插图"), ("footnote", "脚注")):
        v = levels.get(key) or spec.get(key)
        if isinstance(v, dict) and v:
            s = "，".join("%s %s" % (k, val) for k, val in list(v.items())[:4])
            lines.append("· %s：%s" % (label, s))
        elif isinstance(v, str) and v.strip():
            lines.append("· %s：%s" % (label, v.strip()[:40]))

    if not lines:
        lines.append("（未从模板提取到明确的格式要求，将按通用规范处理。）")
    return lines


class App:
    def __init__(self, root):
        global _APP_REF
        self.root = root
        _APP_REF = self
        self.root.title("论文格式医生 · 桌面版")
        # v1.3.97：窗口策略与导师版 v1.0.10~v1.0.12 同款——根治两栏 50:50 布局下
        # 右栏「保存修正后论文」等按钮被裁的问题。根因：两栏各 minsize=470（合计约 980），
        # 旧窗口默认 900 / 可缩到 800，宽度不足时右栏被挤出可视区，按钮显示不全。
        # 1) 默认尺寸按屏幕自适应（不超过屏幕 86%×88%），小屏也能容纳；
        # 2) 最小宽度 1100：50:50 等分后每栏仍有约 520px，按钮/文件名不被挤变形；
        # 3) 最小高度 690：保证底部按钮区与页脚完整可见（五个验收尺寸的最小档）；
        # 4) Windows 且屏幕 ≥1440×900 时启动即最大化（小屏跳过，保留窗口控制权）。
        _sw = self.root.winfo_screenwidth()
        _sh = self.root.winfo_screenheight()
        self.root.geometry("%dx%d" % (min(1280, int(_sw * 0.86)),
                                      min(860, int(_sh * 0.88))))
        self.root.minsize(1100, 690)
        try:
            if sys.platform.startswith("win"):
                if self.root.winfo_screenwidth() >= 1440 and self.root.winfo_screenheight() >= 900:
                    self.root.state("zoomed")
        except tk.TclError:
            pass
        try:
            if os.path.isfile(ICON):
                self.root.iconphoto(True, tk.PhotoImage(file=ICON))
        except Exception:
            pass

        # 窗口尺寸变化只重算左右留白 / 顶栏折行；字号一律取 TYPE 表，
        # 不做运行期缩放（TYPE_SCALE.md §4.4 禁止 int(size*k) 补丁）。
        self._gutters_now = None     # 当前左右留白（避免反复 pack_configure）
        self._gutter_hosts = []      # 需要跟随留白的容器（顶栏/主体/状态栏/页脚）
        self.root.bind("<Configure>", self._on_resize)

        self.thesis_path = tk.StringVar()
        self.template_path = tk.StringVar()
        self.profile_path = tk.StringVar()
        # 状态行初值为空 → 整行隐藏（视觉稿左卡底部只有标语，不摆状态栏）
        self.status_var = tk.StringVar(value="")
        self.running = False
        self._dialog_open = False   # 保存对话框打开期间防重复弹窗
        self._profile_confirmed = False  # 画像是否已被客户确认（检查/修正前弹出确认页）
        self._profile_abandoned = False  # v1.3.49：客户在确认页点"放弃"后，本次不再自动重新提取画像
        # 已保存文件记录：同一会话内再次保存到同一文件时弹“已保存过，是否再次保存”
        self._check_report_saved = None   # 第②步检查报告已保存路径
        self._fix_saved = set()           # 第③步修正产出文件已保存路径集合
        self._fix_out = None              # 第③步修正完成后的临时产出 (dst, chk, rep)
        self._fix_phase = "idle"          # 第③步子状态：idle→fixed→saved（驱动按钮三态）
        self._errored = False
        self._msgs = []
        self.step_defs = [("profile", "提取学校模板要求"),
                          ("check", "论文格式检查"),
                          ("fix", "一键格式修正")]
        self.step_desc = ["识别字号、页边距与格式规范",
                          "快速定位格式问题与待修正项",
                          "确认后生成符合规范的论文文件"]
        self.step_index = 0

        # UI 版本标识与主题令牌（学生版）
        self.edition = "student"
        self._theme = get_theme(self.edition)

        # 先铺官方背景图：底图要在最底层，内容容器创建时才取得到它，
        # 才能各自画出“自己位置那一块”宣纸纹理（见 ui/backdrop.py）。
        self._backdrop = backdrop.Backdrop(self.edition)
        self._backdrop.attach(self.root)

        self._build_style()
        self._build_widgets()

        # v1.3.109：绿色 zip 版首次启动自动创建桌面快捷方式（静默，仅 Windows 正式版生效）
        self.root.after(800, self._maybe_auto_shortcut)

    # -------------------------------------------------- 窗口缩放自适应
    def _on_resize(self, _evt=None):
        try:
            w = self.root.winfo_width()
        except Exception:
            return
        if w < 60:
            return
        self._apply_gutters(w)
        try:
            _tb = getattr(self, "_topbar", None)
            shell_w = _tb.winfo_width() if _tb is not None else 0
        except Exception:
            shell_w = 0        # 窗口正在销毁时控件已不存在，忽略即可
        if shell_w > 40:
            self._apply_topbar_mode(self._topbar_needs_stack(shell_w))

    # -------------------------------------------------- 主体左右留白（露出底图）
    def _gutters(self, w):
        """左右留白：给背景的水墨装饰留出可见边距（照视觉稿 14:9 的比例），
        窗口变窄时自动收窄，保证两张卡片仍放得下按钮与文件名。"""
        avail = w - 2 * 430 - 14
        left = max(20, min(int(w * 0.10), 200))
        right = max(20, min(int(w * 0.06), 132))
        if left + right > avail - 24:
            room = max(40, avail - 24)
            left = max(20, int(room * 0.62))
            right = max(20, room - left)
        return left, right

    def _apply_topbar_mode(self, stacked):
        """顶栏一行 / 两行的落位（只在模式真的变了时才动 pack，避免每次 resize 重排）。"""
        if stacked == self._top_stacked:
            return
        self._top_stacked = stacked
        for row in (self._brand_row, self._entry_row):
            try:
                row.pack_forget()
            except Exception:
                pass
        if stacked:
            self._brand_row.pack(side="top", fill="x")
            self._entry_row.pack(side="top", fill="x", pady=(6, 0))
        else:
            self._brand_row.pack(side="left")
            self._entry_row.pack(side="right")

    def _topbar_needs_stack(self, avail):
        """品牌行与入口行并排放不下时返回 True（按真实测量宽度算，不猜）。"""
        try:
            need = (self._brand.winfo_reqwidth() + self._seal.winfo_reqwidth() + 10
                    + self._entry_row.winfo_reqwidth() + 24)
        except Exception:
            return False
        return avail < need

    def _apply_gutters(self, w):
        left, right = self._gutters(w)
        if (left, right) == self._gutters_now:
            return
        self._gutters_now = (left, right)
        for wdg in self._gutter_hosts:
            try:
                wdg.pack_configure(padx=(left, right))
            except Exception:
                pass

    # ------------------------------------------------------------- style
    def _build_style(self):
        """主界面按钮 / 卡片一律用 ui.widgets 零件，这里只配置仍在用的 ttk.Progressbar。

        旧的 ttk 按钮样式（四个自定义样式名）随老界面一并删除，避免新旧两套并存。
        """
        t = self._theme
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("Tfd.Horizontal.TProgressbar",
                        troughcolor=t.primary_soft, background=t.primary,
                        bordercolor=t.primary_soft, lightcolor=t.primary,
                        darkcolor=t.primary, thickness=6)

    # ------------------------------------------------------------- layout
    def _build_widgets(self):
        t = self._theme
        self.root.configure(bg=t.bg)

        # 所有内容套在一层宣纸“外壳”上：左右留白只改这一层，
        # 顶栏/主体/页脚的 pack 顺序不会被 pack_configure 打乱。
        shell = self._tex(self.root)
        shell.pack(fill="both", expand=True)

        # 顶部：左侧成组的品牌区（线描书本 mark + 大标题 + 朱砂竖章 + 副标题）；
        #       右侧一条入口行（试用状态胶囊 + 升级正式版 → + 客服微信｜官方公众号｜关于｜帮助）
        top_h = max(brand.brand_size(self.edition)[1],
                    brand.seal_size(self.root, self.edition)[1]) + 16
        topbar = self._tex(shell, height=top_h)
        topbar.pack(fill="x", pady=(10, 0))
        self._topbar = topbar
        # 品牌行（mark + 大标题 + 朱砂竖章）与入口行（试用胶囊 + 升级按钮 + 四个文字入口）
        # 各自装一个容器：窗口够宽时并成一行（＝视觉稿），窄了就让入口行折到第二行，
        # 宁可多一行也不把「升级正式版 / 关于 / 帮助」挤出画面。
        brand_row = self._tex(topbar)
        self._brand_row = brand_row
        self._brand = brand.build_brand(brand_row, self.edition)
        self._brand.pack(side="left", pady=(8, 6))
        self._seal = brand.build_seal(brand_row, self.edition)
        self._seal.pack(side="left", padx=(10, 0), pady=(8, 6))

        entry_row = self._tex(topbar)
        self._entry_row = entry_row
        self._links = widgets.LinkRow(entry_row, [
            (COPY_ENTRY_LINKS[0], self._show_contact),
            (COPY_ENTRY_LINKS[1], self._show_contact),
            (COPY_ENTRY_LINKS[2], lambda: show_about(self.root)),
            (COPY_ENTRY_LINKS[3], lambda: show_help(self.root)),
        ], edition=self.edition)
        self._links.pack(side="right", padx=(0, 2))
        self._activate_btn = widgets.RoundButton(
            entry_row, edition=self.edition, text=COPY_UPGRADE, style="primary",
            height=36, font=t.sans(TYPE["btn"], bold=True), paper=True,
            command=self._show_upgrade_qr)
        self._activate_btn.pack(side="right", padx=(12, 14))
        self._trial_badge = widgets.Badge(entry_row, kind="trial", edition=self.edition,
                                          text=COPY_TRIAL_USED, bg=t.bg, paper=True)
        self._trial_badge.pack(side="right")
        self._top_stacked = None
        self._apply_topbar_mode(False)

        # 主体两栏（文件选择 | 处理步骤）
        main = self._tex(shell)
        main.pack(fill="both", expand=True, pady=(12, 8))
        main.columnconfigure(0, weight=1, uniform="half", minsize=430)
        main.columnconfigure(1, weight=1, uniform="half", minsize=430)
        main.rowconfigure(0, weight=1)
        left = self._tex(main)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 14))
        right = self._tex(main)
        right.grid(row=0, column=1, sticky="nsew")
        self._build_left(left)
        self._build_right(right)

        # 底部：页脚（邮箱 ｜ 标语 ｜ 官网 + 客服入口 + 本机处理声明 + 版权）。
        # 旧版那条独立「状态栏」已删除：状态改由左卡底部状态行承载（见 _set_bar）。
        footer = widgets.Footer(shell, edition=self.edition,
                                on_wechat=self._show_contact,
                                on_official=self._show_contact)
        footer.pack(side="bottom", fill="x", pady=(8, 4))

        # 左右留白交给 _on_resize 按窗口宽度驱动（露出背景的水墨装饰）
        self._gutter_hosts = [shell]
        self._gutters_now = None
        self._apply_gutters(self.root.winfo_width() or 1180)

        self._set_bar("idle")
        # 首次布局时卡片高度还没最终确定，_fit_*_blocks 可能按“还没长开”的高度把
        # 「模板说明 / 如何工作」收起，之后内层尺寸不再变化 → 不会再触发 Configure，
        # 就一直停在收起态（实测：默认尺寸启动时左卡底部空一大块）。
        # 窗口真正显示后再复算一次；_fit_*_blocks 幂等，重复调用无副作用。
        self.root.after(120, self._fit_all_blocks)
        self.root.after(520, self._fit_all_blocks)

        self._update_trial_badge()

    # ----------------------------------------------------- 宣纸“透明”容器
    def _tex(self, parent, **kw):
        """内容容器：宣纸底图画布（卡片/按钮之间的空隙透出底图纹理）。"""
        kw.setdefault("edition", self.edition)
        return backdrop.TextureCanvas(parent, **kw)

    # ---------------------------------------------------------- left panel
    def _build_left(self, parent):
        """左栏「文件选择」（视觉稿顺序：上传区 → 学校模板 → 模板说明 → 单选 → 标语）。"""
        t = self._theme
        card = widgets.RoundCard(parent, edition=self.edition, padx=20, pady=14)
        card.pack(fill="both", expand=True)

        # 卡片头：文件选择 + 右上注释
        hdr = tk.Frame(card.inner, bg=t.surface)
        hdr.pack(fill="x")
        tk.Frame(hdr, bg=t.primary, width=4, height=18).pack(side="left", padx=(0, 8))
        _title = tk.Label(hdr, text=COPY_LEFT_TITLE, bg=t.surface, fg=t.navy,
                          font=t.serif(TYPE["section_title"], bold=True))
        _title.pack(side="left")
        self._left_title = _title
        self._left_hdr = hdr
        self._left_hint = _hint_fit(
            tk.Label(hdr, text=COPY_LEFT_HINT, bg=t.surface, fg=t.muted,
                     font=t.sans(TYPE["caption"]), anchor="e", justify="right"), hdr, _title)
        self._left_hint.pack(side="right")

        # 上传区（主/副文案 +「选择文件」次按钮；整块可点，接既有 _pick_input）
        self._thesis_drop = widgets.DropZone(card.inner, edition=self.edition,
                                             command=self._pick_input)
        self._thesis_drop.pack(fill="both", expand=True, pady=(10, 0))

        # 学校模板行：标题 +「可选」徽章 + 说明 +「选择模板」按钮。
        # 按钮**先 pack**（side="right"）：packer 按打包顺序分配空间，先打包的先拿到
        # 自己请求的宽度；否则窄窗下左侧说明会先把宽度吃光，按钮被裁成「择模」。
        self._tpl_row = tk.Frame(card.inner, bg=t.surface, highlightthickness=1,
                                 highlightbackground=t.border)
        self._tpl_row.pack(fill="x", pady=(10, 0))
        widgets.RoundButton(self._tpl_row, edition=self.edition, text=COPY_TPL_BTN,
                            style="secondary", height=34,
                            font=t.sans(TYPE["btn"], bold=True),
                            command=self._pick_template).pack(side="right", padx=12)
        tpl_txt = tk.Frame(self._tpl_row, bg=t.surface)
        tpl_txt.pack(side="left", fill="x", expand=True, padx=(14, 8), pady=8)
        tpl_hdr = tk.Frame(tpl_txt, bg=t.surface)
        tpl_hdr.pack(fill="x")
        tk.Label(tpl_hdr, text=COPY_TPL_TITLE, bg=t.surface, fg=t.navy,
                 font=t.sans(TYPE["body"], bold=True)).pack(side="left")
        widgets.Badge(tpl_hdr, kind="soft", edition=self.edition, text=COPY_TPL_BADGE,
                      bg=t.surface).pack(side="left", padx=(8, 0))
        self._tpl_desc = _auto_wrap(
            tk.Label(tpl_txt, text=COPY_TPL_DESC, bg=t.surface, fg=t.muted,
                     font=t.sans(TYPE["caption"])), tpl_txt, 0)
        self._tpl_desc.pack(anchor="w", pady=(2, 0))

        # 模板说明（模板优先级：批注说明 > 样式定义 > 通用规范）
        self._note = tk.Frame(card.inner, bg=t.info_fill, highlightthickness=1,
                              highlightbackground=t.info_border)
        self._note.pack(fill="x", pady=(8, 0))
        self._note_title = tk.Label(self._note, text=COPY_NOTE_TITLE, bg=t.info_fill,
                                    fg=t.stepper_label_active,
                                    font=t.sans(TYPE["body"], bold=True))
        self._note_title.pack(anchor="w", padx=14, pady=(6, 2))
        self._note_body = _auto_wrap(tk.Label(
            self._note, text=COPY_NOTE_BODY, bg=t.info_fill, fg=t.muted,
            font=t.sans(TYPE["caption"]), justify="left", wraplength=380),
            self._note, 14)
        self._note_body.pack(anchor="w", padx=14, pady=(0, 6))

        # 画像状态（提取后显示；沿用既有 _update_profile_box 驱动）
        self.profile_box = tk.Frame(card.inner, bg=t.success_fill,
                                    highlightthickness=1, highlightbackground=t.success_border)
        tk.Label(self.profile_box, text="●", bg=t.success_fill, fg=t.success_dot,
                 font=t.sans(TYPE["caption"])).pack(side="left", padx=(12, 6), pady=6)
        self.profile_info_var = tk.StringVar(value="已载入格式画像")
        tk.Label(self.profile_box, textvariable=self.profile_info_var, bg=t.success_fill,
                 fg=t.success_text, font=t.sans(TYPE["caption"])).pack(
                     side="left", fill="x", expand=True)
        widgets.RoundButton(
            self.profile_box, edition=self.edition, text="清除", style="secondary",
            height=28, font=t.sans(TYPE["btn_small"], bold=True),
            command=self._clear_profile).pack(side="right", padx=12, pady=4)
        self._update_profile_box()

        # 单选行：未选择论文 / 模板未选（可选）
        self._picks = tk.Frame(card.inner, bg=t.surface)
        self._picks.pack(fill="x", pady=(8, 0))
        self._thesis_radio = widgets.Radio(self._picks, edition=self.edition,
                                           text="", empty_text=COPY_PICK_THESIS,
                                           selected=True, bg=t.surface)
        self._thesis_radio.pack(side="left")
        self._tpl_radio = widgets.Radio(self._picks, edition=self.edition,
                                        text="", empty_text=COPY_PICK_TPL,
                                        bg=t.surface)
        self._tpl_radio.pack(side="left", padx=(28, 0))
        # 兼容既有 Handler：_pick_template / _reset_wizard 都是调 _template_sel.set_text(...)
        self._template_sel = self._tpl_radio

        # 卡片底部标语
        self._slogan_row = tk.Frame(card.inner, bg=t.surface)
        self._slogan_row.pack(fill="x", pady=(6, 0))
        tk.Label(self._slogan_row, text=COPY_LEFT_SLOGAN, bg=t.surface, fg=t.muted,
                 font=t.sans(TYPE["caption"])).pack()

        # 状态行（处理中 / 完成 / 失败）：无状态时整行不显示，底部只留标语
        self._status_row = tk.Frame(card.inner, bg=t.surface)
        self.status_dot = tk.Label(self._status_row, text="●", bg=t.surface, fg=t.muted,
                                   font=t.sans(TYPE["caption"]))
        self.status_dot.pack(side="left", padx=(0, 6))
        self.status_lbl = tk.Label(self._status_row, textvariable=self.status_var,
                                   bg=t.surface, fg=t.navy,
                                   font=t.sans(TYPE["caption"], bold=True),
                                   anchor="w", justify="left")
        self.status_lbl.pack(side="left", fill="x", expand=True)

        # 已选论文 / 模板名同步到单选行（只加观察者，不改既有 Handler）
        self.thesis_path.trace_add("write", lambda *a: self._sync_pick_rows())
        self.template_path.trace_add("write", lambda *a: self._sync_pick_rows())
        self._sync_pick_rows()

        card.inner.bind("<Configure>", self._fit_left_blocks, add="+")

    def _sync_pick_rows(self):
        """把已选论文 / 模板名显示到左卡单选行；未选则回落权威文案。"""
        try:
            tp = self.thesis_path.get().strip()
            self._thesis_radio.set_text(os.path.basename(tp) if tp else "")
            sp = self.template_path.get().strip()
            self._tpl_radio.set_text(os.path.basename(sp) if sp else "")
            self._tpl_radio.set_selected(bool(sp))
        except Exception:
            pass
        self._update_profile_anchor()

    def _update_profile_anchor(self):
        """画像状态框的落位基准（模板说明面板被收起时改挂到模板行下面）。"""
        self._profile_after = (self._note if self._note.winfo_manager()
                              else self._tpl_row)

    def _fit_left_blocks(self, _evt=None):
        """按左卡可用高度决定「能力说明 / 模板说明」是否展示（主内容永不被裁）。"""
        try:
            host = self._note.master
            avail = host.winfo_height()
            if avail < 4:
                return
        except Exception:
            return
        # 卡片过窄/过矮：把卡片头右侧说明整条收起（折行会压住上传区）
        self._toggle_hint(self._left_hint, "left", avail, 430, 470,
                          _hint_fits(self._left_hdr, self._left_title, self._left_hint))
        try:
            panel_w = max(140, host.winfo_width() - 28)
            note_h = (self._note_title.winfo_reqheight()
                      + widgets.wrap_height(self._note_body, panel_w) + 22)
            fixed = (self._left_hdr.winfo_reqheight()
                     + self._tpl_row.winfo_reqheight()
                     + self._picks.winfo_reqheight()
                     + self._slogan_row.winfo_reqheight() + 60)
        except Exception:
            return
        show_note = (avail - fixed - (note_h + 10)) >= self._thesis_drop.min_height
        if show_note:
            if not self._note.winfo_manager():
                self._note.pack(fill="x", pady=(10, 0), before=self._picks)
        elif self._note.winfo_manager():
            self._note.pack_forget()
        self._update_profile_anchor()
    # --------------------------------------------------------- right panel
    def _build_right(self, parent):
        """右栏「处理步骤」（视觉稿顺序：标题 → 步进器 → 三步骤 → 主按钮 → 信任卡）。"""
        t = self._theme
        card = widgets.RoundCard(parent, edition=self.edition, padx=20, pady=16)
        card.pack(fill="both", expand=True)

        hdr = tk.Frame(card.inner, bg=t.surface)
        hdr.pack(fill="x")
        tk.Frame(hdr, bg=t.primary, width=4, height=18).pack(side="left", padx=(0, 8))
        _title = tk.Label(hdr, text=COPY_RIGHT_TITLE, bg=t.surface, fg=t.navy,
                          font=t.serif(TYPE["section_title"], bold=True))
        _title.pack(side="left")
        self._step_counter = tk.Label(hdr, text=COPY_COUNTER_FIRST,
                                      bg=t.surface, fg=t.muted,
                                      font=t.sans(TYPE["caption"]))
        self._step_counter.pack(side="right")
        self._right_title = _title
        self._right_hdr = hdr
        # 右上斜排小字（三步完成 / 论文格式检查与修正）
        self._slant = widgets.SlantNote(hdr, COPY_SLANT_LINES, edition=self.edition)
        self._slant.pack(side="right", padx=(0, 16))
        self._right_hint = self._slant

        self._build_stepper(card.inner)

        self.progress = ttk.Progressbar(card.inner, mode="indeterminate",
                                        style="Tfd.Horizontal.TProgressbar")

        # 可压缩留白：空间不够时先压它，按钮行既不会被挤出卡片、也不压上方内容
        self._right_filler = tk.Frame(card.inner, bg=t.surface)
        self._right_filler.pack(fill="both", expand=True)

        # 按钮行：独占一行（上一步 + 主按钮通栏），与上方内容互不重叠
        self._btn_row = tk.Frame(card.inner, bg=t.surface)
        self._btn_row.pack(fill="x", pady=(6, 6))
        self._prev_btn = widgets.RoundButton(
            self._btn_row, edition=self.edition, text="上一步", style="secondary",
            height=44, width=120, font=t.sans(TYPE["btn"], bold=True),
            command=self._go_prev)
        self._prev_btn.pack(side="left")
        self._next_btn = widgets.RoundButton(
            self._btn_row, edition=self.edition, text=COPY_CTA_FIRST,
            style="primary", height=44, font=t.sans(TYPE["btn"], bold=True),
            command=self._run_step)
        self._next_btn.pack(side="left", fill="x", expand=True, padx=(12, 0))
        self._set_btn_text = lambda btn, text, **kw: btn.config(text=text, **kw)

        # 底部信任卡（视觉稿：学术规范 · 专业高效 / 精准识别格式问题…）——
        # 复用既有 InfoPanel 零件，同一语义不造第二种样式
        self._trust = widgets.InfoPanel(
            card.inner, edition=self.edition, title=COPY_TRUST_TITLE,
            line1=COPY_TRUST_BODY, line2="", height=72)
        self._trust.pack(fill="x")
        card.inner.bind("<Configure>", self._fit_right_blocks, add="+")

        self._refresh_wizard()

    def _fit_right_blocks(self, _evt=None):
        """按右卡可用高度决定底部信任卡 / 步骤说明是否展示（按钮永不被裁）。"""
        try:
            avail = self._trust.master.winfo_height()
            if avail < 4:
                return
        except Exception:
            return
        # 卡片过矮：把卡片头右侧斜排小字整条收起（避免压住标题）
        self._toggle_hint(self._right_hint, "right", avail, 370, 392,
                          _hint_fits(self._right_hdr, self._right_title,
                                     self._right_hint, (self._step_counter,)))
        try:
            need = self._right_hdr.winfo_reqheight() \
                + self._stepper.height_with(False) \
                + self._btn_row.winfo_reqheight() + 18
            need_panel = self._trust.winfo_reqheight()
        except Exception:
            return
        # 三级取舍：① 放得下就显示信任卡；② 只放得下步骤条就把信任卡收起；
        # ③ 连展开版步骤条都放不下，才收起步骤说明文字。
        # 注意②必须拿**展开版**的 need 比，不能拿紧凑版的去比——否则会误判「放得下」，
        # 底部按钮行被挤出卡片外（实测 1920→1040 缩窗时复现）。
        if avail >= need + need_panel:
            self._stepper.set_compact(False)
            if not self._trust.winfo_manager():
                self._trust.pack(fill="x")
        elif avail >= need + 10:
            self._stepper.set_compact(False)
            if self._trust.winfo_manager():
                self._trust.pack_forget()
        else:
            self._stepper.set_compact(True)
            if self._trust.winfo_manager():
                self._trust.pack_forget()

    def _toggle_hint(self, hint, side, avail, hide_below, show_above, fits=True):
        """卡片头说明的显隐（带回滞；``fits`` 为宽度是否够）。

        宽度不够时**整条收起**而不是让它折成两行——折行会把卡片头撑高，
        把下面的上传区/按钮挤出去（实测窄窗下会压住「模板说明」面板）。
        """
        shown = bool(hint.winfo_manager())
        want = bool(fits) and avail >= (show_above if not shown else hide_below)
        if want == shown:
            return
        if want:
            hint.pack(side=side, padx=(0, 14) if side == "right" else 0)
        else:
            hint.pack_forget()
        try:
            self.root.after(30, self._fit_all_blocks)
        except Exception:
            pass


    def _fit_all_blocks(self):
        self._fit_left_blocks()
        self._fit_right_blocks()

    # ----------------------------------------------------- 步骤条（向导）
    def _build_stepper(self, parent):
        t = self._theme
        # 竖排步骤条：完整步骤名 + 说明（圆点/连线/配色照 stepper.svg）
        steps = [(label, self.step_desc[i])
                 for i, (mode, label) in enumerate(self.step_defs)]
        self._stepper = widgets.Stepper(parent, steps=steps, edition=self.edition,
                                        bg=t.surface)
        self._stepper.pack(fill="x", pady=(8, 2))
        self._step_circle = self._stepper.circles
        self._step_title = self._stepper.titles
        self._step_line = self._stepper.lines
        self._step_desc = self._stepper.descs

    def _refresh_wizard(self):
        """根据 step_index / _errored / _fix_phase 重绘时间线、计数与按钮三态。"""
        t = self._theme
        n = len(self.step_defs)
        for i, (mode, label) in enumerate(self.step_defs):
            circ = self._step_circle[i]
            title = self._step_title[i]
            line = self._step_line[i]
            # 第三步（最后一步）修正完成（fixed/saved）也视为完成态
            is_done = i < self.step_index or (
                i == n - 1 and self._fix_phase in ("fixed", "saved"))
            if is_done:
                circ.config(bg=t.surface, fg=t.success_dot,
                            highlightbackground=t.success_dot, text="✔")
                title.config(fg=t.navy)
                if line: line.config(bg=t.success_dot)
            elif i == self.step_index and self._errored:
                circ.config(bg=t.surface, fg=t.error_dot,
                            highlightbackground=t.error_dot, text="✕")
                title.config(fg=t.navy)
                if line: line.config(bg=t.stepper_line)
            elif i == self.step_index:
                # 当前步：实心学术蓝圆点（stepper.svg），编号反白
                circ.config(bg=t.primary, fg="#FFFFFF",
                            highlightbackground=t.primary, text=str(i + 1))
                title.config(fg=t.navy)
                if line: line.config(bg=t.stepper_line)
            else:
                circ.config(bg=t.step_pending_fill, fg=t.muted,
                            highlightbackground=t.step_pending_border,
                            text=str(i + 1))
                title.config(fg=t.muted)
                if line: line.config(bg=t.stepper_line)
        for i in range(n):
            self._step_desc[i].config(
                fg=t.stepper_label_active if i == self.step_index else t.muted)
        self._step_counter.config(
            text=COPY_COUNTER_FMT % (min(self.step_index + 1, n), n))
        if self.step_index >= n:
            self._set_btn_text(self._next_btn, "再处理一篇", command=self._reset_wizard, state="normal")
            self._set_btn_text(self._prev_btn, "上一步", state="disabled", command=self._go_prev)
        elif self.step_index == n - 1:
            # 第三步：一键修正 → 保存修正后论文 → 完成（由 _fix_phase 驱动）
            if self._fix_phase == "idle":
                self._set_btn_text(self._next_btn, "一键修正", command=self._run_step,
                                      state="disabled" if self.running else "normal")
                self._set_btn_text(self._prev_btn, "上一步",
                                      state="disabled" if (self.running or self.step_index == 0) else "normal",
                                      command=self._go_prev)
            elif self._fix_phase == "fixed":
                self._set_btn_text(self._next_btn, "保存修正后论文", command=self._export_fix,
                                      state="disabled" if self.running else "normal")
                self._set_btn_text(self._prev_btn, "上一步", state="normal", command=self._go_prev)
            else:  # saved
                self._set_btn_text(self._next_btn, "完成", state="disabled")
                self._set_btn_text(self._prev_btn, "再处理一篇", state="normal",
                                      command=self._reset_wizard)
        else:
            cta = (COPY_CTA_FIRST if self.step_index == 0 else COPY_CTA_FIX)
            self._set_btn_text(self._next_btn, cta, command=self._run_step,
                                  state="disabled" if self.running else "normal")
            self._set_btn_text(self._prev_btn, "上一步",
                                  state="disabled" if (self.running or self.step_index == 0) else "normal",
                                  command=self._go_prev)

    def _set_bar(self, state, hint=None):
        """状态提示：idle / running / done / error。

        旧的独立「底部状态栏」已删除，四态统一落到左卡底部状态行；idle 不占位
        （视觉稿底部只有标语）。``hint`` 参数保留以兼容既有调用点，文案仍取权威值。
        """
        t = self._theme
        tmap = {
            "idle": ("", t.muted),
            "running": (COPY_STATUS_RUNNING, t.processing_dot),
            "done": (COPY_STATUS_DONE, t.success_dot),
            "error": (COPY_STATUS_FAIL, t.error_dot),
        }
        text, c = tmap.get(state, tmap["idle"])
        self._set_status(text, c)

    def _on_step_done(self, idx, mode):
        t = self._theme
        self._step_circle[idx].config(bg=t.surface, fg=t.success_dot,
                                      highlightbackground=t.success_dot, text="✔")
        self._step_title[idx].config(fg=t.navy)
        if self._step_line[idx]:
            self._step_line[idx].config(bg=t.success_dot)
        if mode == "fix":
            # 第三步：修正完成→进入“保存修正后论文”子状态（先修正、后导出）。
            # 不前进到“再处理一篇”，按钮由 _fix_phase 驱动为「保存修正后论文」。
            self._fix_phase = "fixed"
            self._set_status("修正完成，请点击「保存修正后论文」", t.success_text)
            self._set_bar("done")
        else:
            self.step_index = idx + 1
            self._set_status("已完成", t.success_text)
            self._set_bar("done")
        self._refresh_wizard()

    def _on_step_error(self, idx, mode, err):
        t = self._theme
        self._errored = True
        if mode == "fix":
            self._fix_phase = "idle"
        self._set_status("未能完成，请查看提示", t.error_text)
        self._set_bar("error")
        self._refresh_wizard()
        title = "处理出错"
        msg = err
        if mode == "fix":
            title = "一键修正失败"
            log_path = os.path.join(tempfile.gettempdir(), "tfd_fix_error.log")
            msg = ("一键修正未能完成，论文原文件未被改动。\n\n"
                   "错误信息：\n%s\n\n"
                   "完整报错已记录到：\n%s\n\n"
                   "请把这段信息与该日志文件发给客服，以便定位原因。"
                   % (err, log_path)) + HELP_HINT
        messagebox.showerror(title, msg)

    # --------------------------------------------------------------- picks
    def _pick_input(self):
        p = filedialog.askopenfilename(
            title="选择论文文件",
            filetypes=[("Word 文档", "*.docx *.doc *.wps"), ("所有文件", "*.*")])
        if p:
            self.thesis_path.set(p)
            self._thesis_drop.set_subtitle(os.path.basename(p))
            self._set_status("论文已选择", self._theme.success_text)

    def _pick_template(self):
        p = filedialog.askopenfilename(
            title="选择学校模板",
            filetypes=[("Word 文档", "*.docx *.doc *.wps"), ("所有文件", "*.*")])
        if p:
            self.template_path.set(p)
            self._template_sel.set_text(os.path.basename(p))
            # 换了新模板：重新允许提取画像（清除"放弃"标记）
            self._profile_abandoned = False
            self._profile_confirmed = False

    def _clear_profile(self):
        self.profile_path.set("")
        self._update_profile_box()

    # ---------------------------------------------------------------- run（向导）
    def _go_prev(self):
        """上一步：回退一个步骤，该步及其后的进度重置为待办，可重新执行。

        第三步（修正）内的子状态：先退回“未修正”，停留在第三步便于重新修正，
        而非跳回检查步骤。
        """
        if self.running or self.step_index <= 0:
            return
        t = self._theme
        if self.step_index == len(self.step_defs) - 1 and self._fix_phase != "idle":
            self._fix_phase = "idle"
            self._errored = False
            self._set_status("请按步骤操作", t.muted)
            self._set_bar("idle")
            self._refresh_wizard()
            return
        self.step_index -= 1
        self._errored = False
        # v1.3.49：回退后允许重新确认画像（清除"放弃"标记）
        self._profile_abandoned = False
        self._set_status("请按步骤操作", t.muted)
        self._set_bar("idle")
        self._refresh_wizard()

    def _run_step(self):
        if self._dialog_open:
            return  # 保存对话框已打开，防连点重复弹窗
        if self.running:
            messagebox.showinfo("正在处理", "上一步还在处理中，请稍候…")
            return
        if self.step_index >= len(self.step_defs):
            self._reset_wizard()
            return
        idx = self.step_index
        mode = self.step_defs[idx][0]

        # 校验：论文文件必选（模板可选）
        src = self.thesis_path.get().strip()
        if not src or not os.path.isfile(src):
            messagebox.showerror("缺少输入", "请先选择要处理的论文文件。")
            return

        # 第③步的「保存位置」放到修正【完成之后】再弹（先修正、后导出）：
        # 这里统一 dst=None，_do_fix 先把结果产出到临时目录，_export_fix 收尾时再让客户选位置。
        dst = None

        self._errored = False
        self.running = True
        self._set_running(True)
        status = {"profile": "正在提取学校模板要求…",
                  "check": "正在检查论文格式…",
                  "fix": "正在按学校要求修正论文…"}[mode]
        t = self._theme
        self._set_status(status, t.processing_dot)
        self._set_bar("running", self.step_defs[idx][1])
        # 点亮当前步
        self._step_circle[idx].config(bg=t.surface, fg=t.primary,
                                      highlightbackground=t.primary, text=str(idx + 1))
        self._step_title[idx].config(fg=t.navy)
        threading.Thread(target=self._worker, args=(idx, mode, src, dst), daemon=True).start()

    def _worker(self, idx, mode, src, dst=None):
        try:
            if mode == "profile":
                self._do_profile(src)
                self.root.after(0, lambda: self._on_step_done(idx, mode))
            else:
                docx_path, note = engine.normalize_input(src)
                if note:
                    self._debug(note)
                if mode == "check":
                    # 第二步：先在「学校模板要求」确认窗核对/修改（修改写回画像），
                    # 确认后才开始检查；保存报告放到检查完成后由主线程弹框（_save_check_report）。
                    confirmed = self._confirm_profile_if_needed()
                    if confirmed is None:
                        self.profile_path.set("")   # 客户放弃使用画像 → 按通用规范检查
                    report = self._do_check(src, docx_path)
                    self.root.after(0, lambda: self._save_check_report(report, idx))
                elif mode == "fix":
                    # 修正前同样先确认/修改画像（第一次进修正时）
                    confirmed = self._confirm_profile_if_needed()
                    if confirmed is None:
                        self.profile_path.set("")
                    done = self._do_fix(src, docx_path, dst)
                    if done:
                        self.root.after(0, lambda: self._on_step_done(idx, mode))
                    else:
                        # 试用被拦截（次数用完/取消）：停留在当前步并提示
                        self.root.after(0, self._on_trial_blocked)
        except Exception as e:
            self._debug("[错误] " + str(e))
            self.root.after(0, lambda: self._on_step_error(idx, mode, str(e)))
        finally:
            self.running = False
            # 步骤处理结束：清理旧格式转换产生的临时目录（客户无感，不残留 tfd_conv_*）
            engine.cleanup_conv_dirs()
            self.root.after(0, lambda: self._set_running(False))

    def _set_running(self, running):
        def _apply():
            if running:
                self.progress.pack(fill="x", pady=(8, 4), after=self._stepper)
                self.progress.start(12)
                self._next_btn.config(state="disabled")
            else:
                self.progress.stop()
                self.progress.pack_forget()
                # 完成态允许“再处理一篇”，否则恢复可用
                self._next_btn.config(state="normal")
        self.root.after(0, _apply)

    # --------------------------------------------------- profile 提取与确认
    def _extract_profile(self, src):
        """提取学校模板要求并弹确认页（客户可修改）。返回画像路径或 None（放弃）。"""
        t_docx, note = engine.normalize_input(src)
        if note:
            self._debug(note)
        # 画像 json 放到系统临时目录（不污染客户源目录），按源文件 hash 命名避免多论文覆盖
        out = os.path.join(tempfile.gettempdir(),
                           "tfd_profile_%s.json"
                           % hashlib.sha1(src.encode("utf-8")).hexdigest()[:10])
        text = engine.run_build_profile(t_docx, out)
        try:
            profile = json.loads(text)
        except Exception:
            profile = None
        if profile is None:
            self.profile_path.set("")
            return None
        # 提取成功先不弹确认页：确认页挪到「检查/修正」步骤执行前统一弹出，
        # 客户在真正处理前核对/修改（修改会写回画像并生效）。
        self.profile_path.set(out)
        self._profile_confirmed = False
        self._profile_abandoned = False   # 新画像就绪：清除"放弃"标记
        self.root.after(0, self._update_profile_box)
        return out

    def _ensure_profile_ready(self):
        # v1.3.49：客户已明确"放弃"画像 → 本次会话不再自动重新提取（按通用规范处理），
        # 否则放弃后又会从模板重新提取，放弃形同虚设。
        if self._profile_abandoned:
            return None
        p = self.profile_path.get().strip()
        if p and os.path.isfile(p):
            return p
        t = self.template_path.get().strip()
        if t and os.path.isfile(t):
            return self._extract_profile(t)
        return None

    def _confirm_profile_if_needed(self):
        """检查/修正前：若画像存在且尚未确认，弹「学校模板要求」确认窗（可修改）。

        客户修改的字段会写回画像 json，后续检查/修正均按修改后的要求执行。
        返回画像路径；客户选择“放弃”时返回 None（本次按通用规范处理）。
        """
        p = self._ensure_profile_ready()
        if not p or not os.path.isfile(p) or self._profile_confirmed:
            return p
        try:
            with open(p, encoding="utf-8") as f:
                profile = json.load(f)
        except Exception:
            return p
        ok, edits = self._ask_profile_confirm(profile)
        if not ok:
            # v1.3.49：客户点"放弃"→ 本次会话不再自动重新提取画像，按通用规范处理
            self._profile_abandoned = True
            return None
        self._profile_abandoned = False
        if edits:
            for path_keys, value in edits:
                _deep_set(profile, path_keys, value)
                twips_key = _MARGIN_TWIPS.get(tuple(path_keys))
                if twips_key:
                    try:
                        _deep_set(profile, ("spec", "page", twips_key),
                                  str(int(float(value) * _TWIPS_PER_CM)))
                    except (ValueError, TypeError):
                        pass
            try:
                with open(p, "w", encoding="utf-8") as f:
                    json.dump(profile, f, ensure_ascii=False, indent=2)
            except Exception:
                pass
        self._profile_confirmed = True
        return p

    def _ask_profile_confirm(self, profile):
        ev = threading.Event()
        box = {}

        def show():
            try:
                ok, edits = self._profile_confirm_dialog(profile)
                box["ok"] = ok
                box["edits"] = edits
            except Exception:
                box["ok"] = False
                box["edits"] = []
            finally:
                ev.set()

        self.root.after(0, show)
        # 注意：不能用 ev.wait(timeout=...) 限时等待——客户读完弹窗再点“确认”必然
        # 超过限定时长，导致 ev 超时、box 仍为空、确认结果被丢弃（画像被清空、退回通用规范）。
        # 弹窗通过 wait_window 阻塞主线程，客户关闭后 show() 才返回并 ev.set()，
        # 这里无限等待直到客户做出选择即可（工作线程等待、主线程照常处理弹窗，无死锁）。
        ev.wait()
        return box.get("ok", False), box.get("edits", [])

    def _profile_confirm_dialog(self, profile):
        """“学校模板要求 · 请确认”弹窗：只读摘要 + 可修改关键项。返回 (ok, edits)。"""
        result = {"ok": False, "edits": []}
        t = self._theme
        # 高度按屏幕夹取，避免 768 高屏把「确认 / 放弃」按钮挤出屏外（overrideredirect 无标题栏拖不动）
        try:
            _sh = self.root.winfo_screenheight()
        except Exception:
            _sh = 768
        _dlg_h = min(780, max(520, _sh - 100))
        top = ModalShell(self.root, edition=self.edition,
                         title="学校模板要求 · 请确认", width=660, height=_dlg_h)
        tk.Label(top.body, text="请核对是否与学校规定一致；如有不准，可直接修改后确认。",
                 bg=t.modal_fill, fg=t.modal_sub, font=t.sans(TYPE["caption"]),
                 justify="left", anchor="w").pack(anchor="w", pady=(0, 6))

        # v1.3.46：批注来源提示——有批注=绿色"以批注为准"；无批注=红色警告（防误解、减纠纷）
        # 颜色统一走主题 success / error 语义令牌，不手写第二套。
        _n_cmt = int((profile or {}).get("comment_count", 0) or 0)
        if _n_cmt > 0:
            note_bg, note_bd, note_fg = t.success_fill, t.success_border, t.success_text
            src_txt = "已读取学校模板批注 %d 条 —— 以下要求以【批注】为准（最权威）。" % _n_cmt
        else:
            note_bg, note_bd, note_fg = t.error_fill, t.error_border, t.error_text
            src_txt = ("该模板【未检测到批注】。学校批注是最权威的格式要求；"
                       "无批注时以下要求来自模板样式定义 / 通用规范，"
                       "可能与学校规定有出入，请仔细核对后再确认。")
        src_note = tk.Frame(top.body, bg=note_bg, highlightthickness=1,
                            highlightbackground=note_bd)
        src_note.pack(fill="x", pady=(0, 6))
        tk.Label(src_note, text=src_txt, bg=note_bg, fg=note_fg,
                 font=t.sans(TYPE["caption"]), justify="left", anchor="w",
                 wraplength=580).pack(fill="x", padx=10, pady=8)

        sum_f = tk.Frame(top.body, bg=t.surface, highlightthickness=1,
                         highlightbackground=t.border)
        sum_f.pack(fill="x", pady=3)
        sum_txt = tk.Text(sum_f, height=6, wrap="word", bg=t.surface, fg=t.modal_sub,
                          font=t.sans(TYPE["caption"]), relief="flat", padx=10, pady=6)
        sum_txt.insert("1.0", "\n".join(_profile_summary(profile)))
        sum_txt.config(state="disabled")
        sum_txt.pack(fill="x")

        tk.Label(top.body, text="如需修正，直接修改下列项目（留空表示保持提取结果）",
                 bg=t.modal_fill, fg=t.primary, font=t.sans(TYPE["caption"])).pack(
            anchor="w", pady=(8, 2))

        # 表单区：canvas 与滚动条同在一个 frame 内，滚动条贴右侧整个高度（不沉到右下角）
        form_area = tk.Frame(top.body, bg=t.modal_fill)
        form_area.pack(fill="both", expand=True, pady=3)
        canvas = tk.Canvas(form_area, bg=t.modal_fill, highlightthickness=0)
        vbar = ttk.Scrollbar(form_area, orient="vertical", command=canvas.yview)
        form = tk.Frame(canvas, bg=t.modal_fill)
        form.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=form, anchor="nw")
        canvas.configure(yscrollcommand=vbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        vbar.pack(side="right", fill="y")

        # v1.3.76：确认页表单滚动区同样支持滚轮（15 个字段，鼠标翻页更顺手）
        def _form_wheel(e):
            d = getattr(e, "delta", 0) or 0
            step = -3 if d > 0 else 3
            if abs(d) >= 120:
                step = -int(d / 120) * 3
            canvas.yview_scroll(step, "units")
        for w in (canvas, form):
            w.bind("<MouseWheel>", _form_wheel)
            w.bind("<Button-4>", _form_wheel)
            w.bind("<Button-5>", _form_wheel)

        entries = {}
        for i, (label, candidates, kind) in enumerate(EDIT_FIELDS):
            tk.Label(form, text=label, bg=t.modal_fill, fg=t.modal_sub,
                     font=t.sans(TYPE["caption"])).grid(
                row=i, column=0, sticky="e", padx=(0, 10), pady=3)
            cur = _field_value(profile, candidates)
            if kind == "align":
                var = tk.StringVar(value=ALIGN_DISPLAY.get(cur, cur))
                cb = ttk.Combobox(form, textvariable=var, width=20,
                                  font=t.sans(TYPE["caption"]),
                                  values=list(ALIGN_DISPLAY.values()), state="readonly")
                cb.grid(row=i, column=1, sticky="w", pady=3)
                entries[i] = (candidates, var, "align", var.get())
            elif kind == "ref":
                has_ref = bool(cur)
                var = tk.StringVar(value="已提取" if has_ref else "未提取（按通用规范检查）")
                tk.Entry(form, textvariable=var, width=28,
                         font=t.sans(TYPE["caption"]),
                         relief="flat", bd=0, bg=t.modal_fill, state="disabled",
                         disabledforeground=t.success_text if has_ref else t.muted).grid(
                    row=i, column=1, sticky="w", pady=3)
                entries[i] = (candidates, var, "ref", var.get())
            else:
                var = tk.StringVar(value=cur)
                tk.Entry(form, textvariable=var, width=22,
                         font=t.sans(TYPE["caption"]), relief="solid", bd=1,
                         highlightthickness=1, highlightbackground=t.select_border,
                         bg=t.surface).grid(row=i, column=1, sticky="w", pady=3)
                entries[i] = (candidates, var, "text", var.get())

        _NUM_FIELDS = ("indent_chars", "line_val", "top_cm", "bottom_cm",
                       "left_cm", "right_cm", "before_pt", "after_pt", "hanging_cm")

        def on_confirm():
            edits = []
            for (candidates, var, kind, init_text) in entries.values():
                if kind == "ref":
                    continue  # 参考文献格式为只读提示，不参与修改
                text = var.get().strip()
                if text == init_text or not text:
                    continue
                if kind == "align":
                    code = ALIGN_CODE.get(text, text)
                    edits.append((candidates[0], code))
                    continue
                k = candidates[0][-1]
                if k == "sz":
                    pt = _cn_size_to_pt(text)
                    if pt is None:
                        continue  # 非法输入：保持提取结果
                    edits.append((candidates[0], str(int(round(pt * 2)))))
                elif k in _NUM_FIELDS:
                    try:
                        float(text)
                    except (TypeError, ValueError):
                        continue  # 非法输入：保持提取结果
                    edits.append((candidates[0], text))
                    if k == "indent_chars":
                        edits.append((candidates[0][:-1] + ("indent_type",), "first"))
                    if k == "line_val":
                        edits.append((candidates[0][:-1] + ("line_rule",), "exact"))
                else:
                    edits.append((candidates[0], text))
            result["ok"] = True
            result["edits"] = edits
            top.destroy()

        def on_cancel():
            result["ok"] = False
            top.destroy()

        btns = tk.Frame(top.body, bg=t.modal_fill)
        btns.pack(side="bottom", fill="x", pady=10)
        confirm = widgets.RoundButton(btns, text="确认，使用此要求",
                                     edition=self.edition, style="primary", height=40,
                                     font=t.sans(TYPE["btn"], bold=True),
                                     command=on_confirm)
        confirm.pack(side="left", padx=6)
        cancel = widgets.RoundButton(btns, text="放弃（不使用画像）",
                                     edition=self.edition, style="secondary", height=40,
                                     font=t.sans(TYPE["btn"], bold=True),
                                     command=on_cancel)
        cancel.pack(side="left", padx=6)
        top.after(10, lambda: canvas.yview_moveto(0))
        top.center_on(self.root)
        top.wait_window()
        return result["ok"], result["edits"]

    # ------------------------------------------------------------ 各步骤
    def _do_profile(self, src):
        """第①步：提取【学校模板】的格式要求（不是从论文提取）。

        画像必须来自学校模板才有意义；若未选模板，则没有“学校要求”可提取，
        走通用规范并明确告知客户，且不弹一个空的“确认模板要求”框。
        """
        tpl = self.template_path.get().strip()
        if tpl and os.path.isfile(tpl):
            out = self._extract_profile(tpl)
            if out:
                self._debug("学校模板画像已保存：" + out)
            else:
                self._debug("学校模板要求提取失败（模板可能无样式/批注）")
        else:
            # 未选模板：无学校要求可提取，按通用规范处理，标记已确认避免后续弹空框
            self.profile_path.set("")
            self._profile_confirmed = True
            self.root.after(0, lambda: messagebox.showinfo(
                "无需提取学校要求",
                "未选择学校模板，将按通用论文格式规范检查 / 修正。\n\n"
                "如需按学校具体要求处理，请点“上一步”回到第①步，先选择学校模板再重做。"))
            self.root.after(0, self._update_profile_box)

    def _do_check(self, src, docx_path):
        profile = self._ensure_profile_ready()
        report = engine.run_check(docx_path, profile_path=profile)
        self._debug(report)
        return report

    def _save_check_report(self, report, idx):
        """主线程：检查完成后让客户选择保存位置，写入 Word 报告并收尾该步骤。

        同一会话内已保存过检查报告时，再次保存前弹“已保存过，是否再次保存”确认。
        """
        base = _base_no_ext(self.thesis_path.get().strip() or "report")
        dst = filedialog.asksaveasfilename(
            title="选择检查报告保存位置",
            initialfile=os.path.basename(base) + "_格式检查报告.docx",
            initialdir=os.path.dirname(base) or None,
            defaultextension=".docx",
            filetypes=[("Word 文档", "*.docx")])
        saved = False
        if dst:
            if self._check_report_saved:
                again = messagebox.askyesno(
                    "已保存过",
                    "检查报告之前已保存过：\n%s\n\n确定要再次保存（覆盖）吗？"
                    % os.path.basename(self._check_report_saved))
                if not again:
                    dst = None
            if dst:
                try:
                    engine.md_to_docx(report, dst)
                    self._check_report_saved = dst
                    saved = True
                except Exception as e:
                    messagebox.showerror("保存失败", str(e))
                    dst = None
        else:
            self._debug("客户未选择保存位置，检查报告未落盘")
        self._on_step_done(idx, "check")
        if saved:
            self._show_check_done(dst)

    def _guard_license(self):
        """付费操作前守卫：本地失效立即拦；联网实时校验（次卡扣次、其他卡只校验）。
        返回 True=放行；False=已拦截。离线放行（不拦不扣）。"""
        if license.is_revoked():
            self.root.after(0, self._handle_revoked)
            return False
        st = license.consume_use()
        if st in ("revoked", "expired"):
            self.root.after(0, self._handle_revoked)
            return False
        return True

    def _do_fix(self, src, docx_path, dst):
        # 授权实时校验（防退款撤销/到期/次卡用尽）+ 次卡联网扣次
        if not self._guard_license():
            return False
        # v1.3.56 试用版：未激活时只能试 N 次，输出带水印；用完弹升级引导。
        # 返回 True=已产出修正；False=被试用拦截（次数用完/客户取消），不进入完成态。
        licensed = trial.is_licensed()
        if not licensed:
            left = trial.trials_left()
            if left <= 0:
                self._show_trial_exhausted()   # worker 线程内调用（内部按调用线程切主线程弹窗）
                return False
            ok = self._ask_trial_confirm(left)  # 主线程弹"修正前提示"
            if not ok:
                return False
            if not trial.consume_trial():
                # v1.3.84：扣减失败（计数文件被篡改锁定/损坏/写入失败）→ 按已用完拦截。
                # 此前返回值被忽略：篡改 trial.json 使其"损坏"后，每次都不扣次数仍放行，
                # 锁定机制形同虚设、可无限试用——安全边界修复。
                self._show_trial_exhausted()
                return False

        # dst 为 None：先产出到临时目录，待修正完成（主线程 _export_fix）再让客户选保存位置。
        if not dst:
            tmp = tempfile.mkdtemp(prefix="tfd_fix_")
            base = os.path.join(tmp, os.path.splitext(os.path.basename(src))[0])
            dst = base + "_已修正.docx"
            rep = base + "_修改报告.docx"
            chk = base + "_检查报告.docx"
        else:
            base_dst = _base_no_ext(dst)
            rep = base_dst + "_修改报告.docx"
            chk = base_dst + "_检查报告.docx"
        profile = self._ensure_profile_ready()
        # 主交付物：修正后的论文 + 修改明细报告（run_fix_headings 内部已写盘）
        report = engine.run_fix_headings(
            docx_path, dst, profile_path=profile,
            report_docx=rep, add_comments=True)
        if not licensed:
            # 试用版：给修正后的论文加水印+只读保护，并让客户知道正式版可编辑无水印
            if watermark.apply_watermark(dst):
                report += ("\n\n> 本预览版带水印且为【只读】文档（编辑需密码，仅作效果预览）；"
                           "激活码解锁后输出可编辑无水印正式版，可一键交稿。\n")
            else:
                self._debug("[试用水印注入失败，已跳过]")
        self._debug(report)
        # 修改明细报告若因引擎内报告环节异常未落盘，置空（主交付物已修正论文不受影响）
        if rep and not os.path.isfile(rep):
            rep = None
        # 次要交付物：修正后的检查报告（生成失败不应阻断主交付物）
        try:
            check_report = engine.run_check(dst, profile_path=profile)
            engine.md_to_docx(check_report, chk)
            self._debug(check_report)
        except Exception as e:
            self._debug("[检查报告生成失败，已跳过] " + str(e))
            chk = None
        # 修正已落到临时文件，结果交给主线程在「修正完成」后导出（先修正、后导出）
        self._fix_out = (dst, chk, rep)
        return True

    def _ask_trial_confirm(self, left):
        """修正前提示（worker 线程调用，主线程弹窗，Event 同步）。"""
        ev = threading.Event()
        box = {}

        def show():
            box["ok"] = messagebox.askyesno(
                "试用版提示",
                "当前为试用版（本机剩余 %d 次），修正后的论文将带水印，且为只读预览文档"
                "（编辑需密码）。\n\n"
                "正式版输出可编辑无水印文档，可一键交稿。\n\n是否继续？" % left,
                parent=self.root)
            ev.set()

        self.root.after(0, show)
        ev.wait(30)
        return bool(box.get("ok", False))

    def _show_trial_exhausted(self):
        """试用次数用完：弹购买引导（小程序码 + 去激活入口）。

        可能在 worker 线程（_do_fix 内）也可能在**主线程**（root.after 排程）被调用：
          * worker 线程 → 用 Event 等主线程把弹窗弹完再返回；
          * 主线程     → **直接弹**，绝不能 ev.wait()：那会阻塞事件循环，
                         排队的 show 永远跑不到，界面要等超时才恢复（Codex 查出的冻结）。
        """
        text = ("本机免费试用（%d 次）已用完。\n\n"
                "正式版激活后：不限次数修正、输出无水印可编辑文档、一键交稿。\n\n"
                "微信扫下方小程序码即可购买激活码，付款后自动发码，立即可用。"
                % trial.TRIAL_LIMIT)
        if threading.current_thread() is threading.main_thread():
            self._show_purchase_qr("免费试用已用完", text)
            return
        ev = threading.Event()

        def show():
            try:
                self._show_purchase_qr("免费试用已用完", text)
            finally:
                ev.set()

        self.root.after(0, show)
        ev.wait(60)

    def _show_upgrade_qr(self):
        """右上「升级正式版 →」入口：直接弹简单小程序码购买引导（老板拍板：不弹复杂激活窗 / 套餐卡片）。
        已购用户点「我已购买 · 去激活」仍走 show_activation 完整激活流程，激活能力不丢。
        """
        self._show_purchase_qr(
            "升级正式版",
            "正式版：不限次数修正、输出无水印可编辑文档、一键交稿。\n\n"
            "微信扫下方小程序码即可购买激活码，付款后自动发码，立即可用。")

    def _show_purchase_qr(self, title, text):
        """购买引导窗：展示小程序码（购买主渠道）+「我已购买 · 去激活」入口。

        必须在主线程调用（_show_trial_exhausted 已用 root.after(0) 调度）。
        小程序码缺失时降级为文字指引，绝不阻断主流程。
        """
        res = {"v": None}
        t = self._theme
        top = ModalShell(self.root, edition=self.edition, title=title, width=600, height=660)
        tk.Label(top.body, text=text, bg=t.modal_fill, fg=t.modal_sub,
                 font=t.sans(TYPE["body"]), justify="left",
                 wraplength=440).pack(padx=0, pady=(0, 10))
        shown = False
        try:
            if os.path.isfile(MINIAPP_QRCODE):
                img = tk.PhotoImage(file=MINIAPP_QRCODE)
                img = img.subsample(max(1, round(img.width() / 150)))
                lbl = tk.Label(top.body, image=img, bg=t.modal_fill)
                lbl.image = img          # 防 GC 回收导致图片不显示
                lbl.pack()
                tk.Label(top.body, text="微信扫一扫，或搜索【%s】小程序，付款后自动发码"
                         % MINIAPP_NAME, bg=t.modal_fill, fg=t.muted,
                         font=t.sans(TYPE["caption"])).pack(pady=(6, 0))
                shown = True
        except Exception:
            shown = False
        if not shown:
            tk.Label(top.body, text="（小程序码未随包提供，请前往官网 reedskill.com 购买）",
                     bg=t.modal_fill, fg=t.muted, font=t.sans(TYPE["caption"])).pack(pady=(4, 0))
        fr = tk.Frame(top.body, bg=t.modal_fill)
        fr.pack(pady=(16, 18))
        widgets.RoundButton(fr, text="我已购买 · 去激活", edition=self.edition,
                           style="primary", height=40,
                           font=t.sans(TYPE["btn"], bold=True),
                           command=lambda: (res.update(v="activate"),
                                           top.destroy())).pack(side="left", padx=6)
        widgets.RoundButton(fr, text="稍后", edition=self.edition,
                           style="secondary", height=40,
                           font=t.sans(TYPE["btn"], bold=True),
                           command=top.destroy).pack(side="left", padx=6)
        top.center_on(self.root)
        top.wait_window()
        if res["v"] == "activate":
            self._open_activation(force=True)

    def _on_trial_blocked(self):
        """试用被拦截（用完/取消）：停留在当前步，给友好提示。"""
        self._errored = False
        # 用户刚在「购买引导」里点「去激活」并激活成功时，worker 仍会返回 False 走到这里；
        # 此时不能再报「试用已用完」，否则与刚写入的「已激活正式版」自相矛盾（Codex N1）。
        try:
            if trial.is_licensed():
                self._update_trial_badge()
                self._set_status("已激活正式版，感谢支持", OKC)
                self._set_bar("idle")
                self._refresh_wizard()
                return
        except Exception:
            pass
        self._set_status("试用次数已用完，请激活后使用", ERRC)
        self._set_bar("idle")
        self._refresh_wizard()

    # ------------------------------------------------------------ 确认页
    def _show_check_done(self, out_md):
        k = self._modal("格式检查完成",
                        "检查报告已保存至：\n%s\n\n如需按学校要求修正论文，请继续第③步。" % out_md,
                        [("open", "打开所在文件夹"), ("ok", "完成")])
        if k == "open":
            self._open_folder(out_md)

    def _show_fix_done(self, dst, chk, rep):
        lines = ["已为您保存以下文件：\n"]
        lines.append("① 修正后论文：%s" % dst)
        lines.append("② 修改报告：%s" % rep)
        lines.append("③ 检查报告：%s" % (chk if chk else "（本次未生成，可点“再处理一篇”重试）"))
        lines.append("\n全程在本机完成，原文件未改动。")
        k = self._modal("修正完成", "\n".join(lines),
                        [("open", "打开所在文件夹"), ("ok", "完成")])
        if k == "open":
            self._open_folder(dst)

    def _export_fix(self):
        """主线程：修正完成后让客户选择保存位置并导出三个文件（先修正、后导出）。

        同一会话内已导出过则先弹“已保存过，是否再次保存（覆盖）”。
        """
        out = getattr(self, "_fix_out", None)
        if not out:
            return
        tmp_dst, chk, rep = out
        if self._fix_saved:
            prev = sorted(self._fix_saved)
            if not messagebox.askyesno(
                    "已保存过",
                    "修正文件之前已保存过：\n%s\n\n确定要再次生成并保存（覆盖）吗？"
                    % "\n".join(os.path.basename(p) for p in prev)):
                return
        base_src = _base_no_ext(self.thesis_path.get().strip() or "论文")
        dst = filedialog.asksaveasfilename(
            title="选择修正后论文的保存位置",
            initialfile=os.path.basename(base_src) + "_已修正.docx",
            initialdir=os.path.dirname(base_src) or None,
            defaultextension=".docx",
            filetypes=[("Word 文档", "*.docx")])
        if not dst:
            messagebox.showinfo("未导出",
                                "未选择保存位置，本次修正结果未导出。\n"
                                "如需导出，请点“上一步”回到第③步重新执行。")
            return
        try:
            base = _base_no_ext(dst)
            shutil.copy(tmp_dst, dst)
            saved = {dst}
            if rep:
                shutil.copy(rep, base + "_修改报告.docx")
                saved.add(base + "_修改报告.docx")
            if chk:
                shutil.copy(chk, base + "_检查报告.docx")
                saved.add(base + "_检查报告.docx")
            self._fix_saved = saved
        except Exception as e:
            messagebox.showerror("导出失败", str(e) + HELP_HINT)
            return
        # 导出完成：清理修正临时产出目录（系统 temp，客户无感，不残留 tfd_fix_*）
        try:
            shutil.rmtree(os.path.dirname(tmp_dst), ignore_errors=True)
        except Exception:
            pass
        # 导出成功 → 进入“完成”子状态（按钮变“完成”，上一步变“再处理一篇”）
        self._fix_phase = "saved"
        self._set_status("已保存，可再处理一篇", OKC)
        self._refresh_wizard()
        self._show_fix_done(dst, (base + "_检查报告.docx") if chk else None,
                           base + "_修改报告.docx")

    def _modal(self, title, text, buttons):
        """通用确认/选择弹窗：ModalShell 壳 + RoundButton，视觉统一走 ui 零件。"""
        result = {"v": None}
        t = self._theme
        top = ModalShell(self.root, edition=self.edition, title=title,
                         width=620, height=340)
        tk.Label(top.body, text=text, bg=t.modal_fill, fg=t.modal_sub,
                 font=t.sans(TYPE["body"]), justify="left", wraplength=520).pack(
            anchor="w", pady=(6, 24))
        fr = tk.Frame(top.body, bg=t.modal_fill)
        fr.pack(fill="x", pady=(0, 6))
        for key, label in buttons:
            st = "primary" if key == "ok" else "secondary"
            b = widgets.RoundButton(fr, text=label, edition=self.edition, style=st,
                                   height=40,
                                   font=t.sans(TYPE["btn"], bold=True),
                                   command=lambda k=key: (result.update(v=k),
                                                          top.destroy()))
            b.pack(side="left", padx=(0, 10))
        top.center_on(self.root)
        top.wait_window()
        return result["v"]

    def _open_folder(self, path):
        folder = os.path.dirname(os.path.abspath(path))
        try:
            if sys.platform.startswith("win"):
                os.startfile(folder)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.Popen(["open", folder])
            else:
                subprocess.Popen(["xdg-open", folder])
        except Exception:
            pass

    # ------------------------------------------------- 桌面快捷方式（v1.3.108 绿色 zip 自建）
    def _desktop_dir(self):
        """真实桌面路径（兼容 OneDrive 重定向）；拿不到时退回 ~/Desktop。"""
        try:
            import win32com.client as _wc
            ws = _wc.Dispatch("WScript.Shell")
            d = ws.SpecialFolders("Desktop")
            if d:
                return d
        except Exception:
            pass
        return os.path.join(os.path.expanduser("~"), "Desktop")

    def _desktop_shortcut_path(self):
        return os.path.join(self._desktop_dir(), APP_SHORTCUT_NAME + ".lnk")

    def _desktop_shortcut_exists(self):
        try:
            return os.path.isfile(self._desktop_shortcut_path())
        except Exception:
            return False

    def _create_desktop_shortcut(self, silent=False):
        """创建桌面快捷方式（指向当前主程序）。成功 True；silent=True 时静默不弹提示。"""
        try:
            if not (sys.platform.startswith("win") and _is_frozen_exe()):
                if not silent:
                    messagebox.showinfo("桌面快捷方式",
                                        "仅 Windows 正式版支持自动创建；\n"
                                        "可手动：右键主程序 → 发送到 → 桌面快捷方式。",
                                        parent=self.root)
                return False
            import win32com.client as _wc
            ws = _wc.Dispatch("WScript.Shell")
            exe = os.path.abspath(sys.executable)
            lnk = self._desktop_shortcut_path()
            sc = ws.CreateShortcut(lnk)
            sc.TargetPath = exe
            sc.WorkingDirectory = os.path.dirname(exe)
            sc.IconLocation = exe + ",0"
            sc.Save()
            if not silent:
                messagebox.showinfo("桌面快捷方式",
                                    "已创建「%s」桌面快捷方式，双击即可打开。" % APP_SHORTCUT_NAME,
                                    parent=self.root)
            return True
        except Exception as e:
            if not silent:
                messagebox.showerror("创建失败",
                                     "自动创建失败：%s\n\n"
                                     "请手动：右键主程序 → 发送到 → 桌面快捷方式。" % e,
                                     parent=self.root)
            return False

    def _maybe_auto_shortcut(self):
        """首次启动自动在桌面创建一次快捷方式（仅 Windows 正式版、桌面尚无该图标时）。
        静默执行，不弹窗；已创建过则不再重复（标记文件存 %APPDATA%\\SHORTCUT_FLAG_TAG）。
        顶栏“桌面图标”链接可随时手动补建 / 重建。"""
        try:
            if not (sys.platform.startswith("win") and _is_frozen_exe()):
                return
            if self._desktop_shortcut_exists():
                return
            flag = os.path.join(os.environ.get("APPDATA", ""), SHORTCUT_FLAG_TAG,
                                "shortcut_auto.txt")
            if os.path.isfile(flag):
                return
            os.makedirs(os.path.dirname(flag), exist_ok=True)
            with open(flag, "w", encoding="utf-8") as f:
                f.write("1")
            self._create_desktop_shortcut(silent=True)
        except Exception:
            pass

    def _update_profile_box(self):
        t = self._theme
        p = self.profile_path.get().strip()
        if p and os.path.isfile(p):
            self.profile_info_var.set("已载入格式画像：" + os.path.basename(p))
            if not self.profile_box.winfo_ismapped():
                anchor = getattr(self, "_profile_after", None)
                if anchor is None:
                    anchor = (self._note if self._note.winfo_manager()
                              else self._tpl_row)
                self.profile_box.pack(fill="x", pady=(10, 0), after=anchor)
        else:
            if self.profile_box.winfo_ismapped():
                self.profile_box.pack_forget()

    def _reset_wizard(self):
        """“再处理一篇”：回到第 1 步并清空选择。"""
        t = self._theme
        self.step_index = 0
        self._errored = False
        self._profile_confirmed = False
        self._profile_abandoned = False
        self._check_report_saved = None
        self._fix_saved = set()
        self._fix_out = None
        self._fix_phase = "idle"
        self.thesis_path.set("")
        self.template_path.set("")
        self.profile_path.set("")
        self._thesis_drop.set_subtitle("")
        self._template_sel.set_text("")
        self._update_profile_box()
        self._set_status("", t.muted)
        self._set_bar("idle")
        self._refresh_wizard()

    # ---------------------------------------------- 内部日志（不展示客户）
    def _debug(self, text):
        self._msgs.append(text)
        if len(self._msgs) > 200:
            self._msgs = self._msgs[-200:]

    def _set_status(self, s, color):
        """左卡底部状态行：设文字 / 配色，并在无文字时整行隐藏（不占位）。"""
        def _apply():
            self.status_var.set(s)
            self.status_lbl.config(fg=color)
            self.status_dot.config(fg=color)
            row = getattr(self, "_status_row", None)
            if row is None:
                return
            if (s or "").strip():
                if not row.winfo_manager():
                    row.pack(fill="x", pady=(6, 0))
            elif row.winfo_manager():
                row.pack_forget()
        try:
            self.root.after(0, _apply)
        except Exception:
            pass

    def _update_trial_badge(self):
        """刷新右上角入口行：试用状态胶囊 + 升级/激活按钮（只改渲染，判定逻辑不变）。"""
        try:
            self._licensed = trial.is_licensed()
        except Exception:
            self._licensed = False
        if self._licensed:
            try:
                s = license.license_summary()
            except Exception:
                s = None
            self._trial_badge.set_kind("formal")
            # 正式徽章：正式版 · 月卡 / 正式版 · 周卡 / 正式版 · 次卡 …
            kind = (s or {}).get("kind") or ""
            if kind in COPY_FORMAL_LABELS:
                label = COPY_FORMAL_LABELS[kind]
            elif kind and kind != "正式版":
                label = COPY_FORMAL_FMT % kind
            else:
                label = COPY_FORMAL_FALLBACK
            self._trial_badge.set_text(label)
            self._activate_btn.set_text("已激活 ✓")
            self._activate_btn.config(state="disabled")
        else:
            left = trial.trials_left()
            self._trial_badge.set_kind("trial")
            self._trial_badge.set_text(COPY_TRIAL_FMT % left if left > 0
                                       else COPY_TRIAL_USED)
            self._activate_btn.set_text(COPY_UPGRADE)
            self._activate_btn.config(state="normal")

    def _show_contact(self):
        """右上「客服微信 / 官方公众号」入口：只展示程序里既有的客服信息，不改口径。"""
        self._modal("联系客服",
                    "客服微信：【%s】 ID：%s\n官方公众号：【%s】 ID：%s\n"
                    "合作邮箱：%s\n\n关注公众号后留言，我们看到即回。"
                    % (WECHAT_NAME, WECHAT_ID, WECHAT_NAME, WECHAT_ID, ABOUT_MAIL),
                    [("ok", "知道了")])

    def _open_activation(self, force: bool = False):
        """主界面右上角激活入口：打开激活窗，成功后刷新状态。

        ``force=True``：来自「试用用完 → 购买引导」弹窗的「我已购买 · 去激活」。
        此时 worker 线程仍在跑（`self.running` 要等 _do_fix 返回才清），走 running 守卫
        会直接弹「正在处理」→ 激活窗永远打不开（Codex 2026-09-12 审查发现的 P1）。
        """
        if self.running and not force:
            messagebox.showinfo("正在处理", "上一步还在处理中，请稍候…", parent=self.root)
            return
        r = show_activation(self.root, show_trial=False)
        if r == "ok":
            self._update_trial_badge()
            try:
                s = license.license_summary()
            except Exception:
                s = None
            self._set_status(("已激活%s · %s" % (s["kind"], s["desc"])) if s
                             else "已激活正式版，感谢支持", OKC)
        elif r == "trial":
            self._update_trial_badge()
            self._set_status("已进入试用模式", MUTED)

    def _handle_revoked(self):
        """授权失效（撤销退款/到期/次卡用尽）：锁定软件，禁止继续修正。"""
        self._licensed = False
        self._update_trial_badge()
        reason = license.get_lock_reason()
        title, body = _revoked_message(reason)
        if reason == "uses_exhausted":
            status = "次数已用完，软件已锁定"
        elif reason == "expired":
            status = "授权已到期，软件已锁定"
        else:
            status = "授权已失效（可能已退款），软件已锁定"
        self._set_status(status, ERRC)
        self._modal(title, body, [("ok", "知道了")])


# ---------------------------------------------------------------------------
_APP_REF = None  # App 实例引用，在 App.__init__ 中赋值


def _on_license_revoked():
    """心跳线程发现 revoked：通过 after(0) 切回主线程处理，避免后台线程碰 Tk。"""
    if _APP_REF is not None:
        try:
            _APP_REF.root.after(0, _APP_REF._handle_revoked)
        except Exception:
            pass


def _revoked_message(reason=None):
    """按失效原因返回 (标题, 正文)。"""
    if reason == "uses_exhausted":
        return ("次数已用完",
                "您的次卡次数已用完，软件已锁定。\n\n"
                "如需继续使用，请续费或联系客服：hi@reedskill.com。")
    if reason == "expired":
        return ("授权已到期",
                "您的授权已到期，软件已锁定。\n\n"
                "如需继续使用，请续费或联系客服：hi@reedskill.com。")
    return ("授权已失效",
            "您的授权已被后台撤销（可能因退款）。\n\n"
            "软件已锁定，无法继续修正论文。\n如需继续使用，请联系客服：hi@reedskill.com。")


def show_activation(root, show_trial=True):
    """激活窗口：在线激活（主） + 离线备用码（兜底）。

    show_trial: 是否显示"先试用"入口。仅启动时首次弹窗为 True；
    从主界面激活入口打开时客户已在试用模式，无需再显示（v1.3.66）。

    返回：
      "ok"     激活成功；
      "trial"  客户选择"先试用"（暂不激活，进入试用模式）；
      "quit"   直接关闭窗口（退出程序）。
    """
    result = {"v": "quit"}
    t = get_theme("student")
    # v1.3.65：窗口尺寸自适应屏幕（笔记本 768 高屏幕时 720 高的窗口底部会被截在屏外）
    try:
        _sw = root.winfo_screenwidth()
        _sh = root.winfo_screenheight()
    except Exception:
        _sw, _sh = 1366, 768
    _base_h = 700 if show_trial else 610
    _w = max(420, min(620, _sw - 60))
    _h = max(420, min(_base_h, _sh - 120))
    top = ModalShell(root, edition="student", title="激活 · 论文格式医生",
                    width=_w, height=_h)

    tk.Label(top.body, text="请输入您购买的激活码以激活。\n"
                            "永久卡：激活后完全离线、永久可用；周卡 / 月卡：期限内可用；\n"
                            "次卡：每次修正需联网扣一次次数，离线时不可使用。",
             bg=t.modal_fill, fg=t.modal_sub, font=t.sans(TYPE["caption"]),
             wraplength=520, justify="center").pack(pady=(0, 10))

    card_var = tk.StringVar()
    tk.Entry(top.body, textvariable=card_var, font=t.sans(TYPE["body"]),
             relief="solid", bd=1, highlightthickness=1,
             highlightbackground=t.select_border, bg=t.surface,
             justify="center").pack(fill="x", pady=(4, 6))

    msg_var = tk.StringVar()
    tk.Label(top.body, textvariable=msg_var, bg=t.modal_fill, fg=t.error_text,
             font=t.sans(TYPE["caption"])).pack(pady=(0, 6))

    def do_activate():
        card = card_var.get().strip()
        if not card:
            msg_var.set("请输入激活码")
            return
        activate_btn.config(state="disabled")
        msg_var.set("正在验证您的授权，请稍候…（首次激活需联网校验，通常需要 20~40 秒）")
        q = queue.Queue()
        start_t = time.time()

        def work():
            # 工作线程只做两件事：跑验证 + 把结果放进队列；绝不直接碰 Tk 界面
            try:
                mc = license.get_machine_code()
                license._log("gui: 开始联网验证 card=%s mc=%s" % (card, mc))
                info = license.activate_online(card, mc)
                ok = info.get("ok", False)
                note = info.get("error") or "验证通过"
                license._log("gui: 联网验证返回 ok=%s note=%s" % (ok, note))
                q.put(("done", ok, note, mc, info))
            except Exception:
                import traceback
                tb = traceback.format_exc()
                license._log("gui: 激活线程异常\n%s" % tb)
                q.put(("done", False, "激活过程出错，请重试或联系客服", "", {}))

        threading.Thread(target=work, daemon=True).start()

        def poll():
            # 主线程轮询队列（线程安全），拿到结果才更新界面
            try:
                _kind, ok, note, mc, info = q.get_nowait()
            except queue.Empty:
                if time.time() - start_t > 120:   # UI 看门狗：物理上不可能无限转圈
                    activate_btn.config(state="normal")
                    msg_var.set("验证超时：请检查网络后重试；或联系客服获取离线激活码")
                    license._log("gui: UI 看门狗触发（120 秒未等到结果）")
                    return
                top.after(300, poll)
                return
            activate_btn.config(state="normal")
            if ok:
                license.save_local_license(
                    card, mc,
                    permanent=(info.get("type") == "lifetime"),
                    lic_type=info.get("type"),
                    expires_at=info.get("expires_at"),
                    product=license.DEFAULT_PRODUCT,
                    uses_total=info.get("uses_total"),
                    uses_used=info.get("uses_used"),
                )
                license._log("gui: 授权已写入本机")
                # 启动后台心跳：联网时若授权被撤销（退款），自动锁死
                license.start_heartbeat(
                    product=license.DEFAULT_PRODUCT, on_revoked=_on_license_revoked)
                result["v"] = "ok"
                top.destroy()
            else:
                msg_var.set(note)

        top.after(300, poll)

    activate_btn = widgets.RoundButton(top.body, text="激活", edition="student",
                                      style="primary", height=40,
                                      font=t.sans(TYPE["btn"], bold=True),
                                      command=do_activate)
    activate_btn.pack(pady=(2, 4))

    # v1.3.58：试用入口放在显眼位置（激活按钮正下方，便于未购买客户先体验）
    # v1.3.66：仅启动时首次弹窗显示；从主界面激活入口打开时客户已在试用模式，无需再显示
    if show_trial:
        widgets.RoundButton(top.body,
                           text="还没有激活码？先试用（免费 %d 次）" % trial.TRIAL_LIMIT,
                           edition="student", style="secondary", height=40,
                           font=t.sans(TYPE["btn"], bold=True),
                           command=lambda: (result.update(v="trial"),
                                           top.destroy())).pack(pady=(4, 2))
        tk.Label(top.body, text="试用版可完整体验一键修正，输出带水印且为只读预览；正式版可编辑无水印。",
                 bg=t.modal_fill, fg=t.muted, font=t.sans(TYPE["caption"]),
                 wraplength=540).pack(pady=(0, 6))

    tk.Label(top.body, text="— 以下为特殊情形使用 —", bg=t.modal_fill, fg=t.muted,
             font=t.sans(TYPE["caption"])).pack(pady=(8, 4))
    mc = license.get_machine_code()
    # 机器码 + 复制按钮并排一行（省高度）
    mc_row = tk.Frame(top.body, bg=t.modal_fill)
    mc_row.pack(pady=(0, 6))
    tk.Label(mc_row, text="本机机器码：" + mc, bg=t.modal_fill, fg=t.muted,
             font=("Consolas", TYPE["mono"])).pack(side="left")
    widgets.RoundButton(mc_row, text="复制", edition="student", style="secondary",
                       height=26, font=t.sans(TYPE["btn_small"], bold=True),
                       command=lambda: top.clipboard_append(mc)).pack(side="left", padx=(8, 0))

    off_var = tk.StringVar()
    tk.Label(top.body, text="离线激活码（网络不通时，联系客服获取）：", bg=t.modal_fill, fg=t.muted,
             font=t.sans(TYPE["caption"])).pack(pady=(2, 2))
    tk.Entry(top.body, textvariable=off_var, font=("Consolas", TYPE["mono"]),
             relief="solid", bd=1, highlightthickness=1,
             highlightbackground=t.select_border, bg=t.surface).pack(pady=(2, 4))

    def do_offline():
        code = off_var.get().strip()
        if not code:
            msg_var.set("请输入离线激活码")
            return
        mc2 = license.get_machine_code()
        if license.verify_offline_code(code, mc2):
            license.save_offline_license(code, mc2)
            result["v"] = "ok"
            top.destroy()
        else:
            msg_var.set("离线激活码无效，请核对后重试")

    offline_btn = widgets.RoundButton(top.body, text="使用离线激活码激活", edition="student",
                                     style="secondary", height=40,
                                     font=t.sans(TYPE["btn"], bold=True),
                                     command=do_offline)
    offline_btn.pack(pady=(2, 4))

    widgets.RoundButton(top.body, text="退出", edition="student", style="secondary",
                       height=34, font=t.sans(TYPE["btn"], bold=True),
                       command=lambda: top.destroy()).pack(pady=(0, 4))

    hl = tk.Frame(top.body, bg=t.modal_fill)
    hl.pack(pady=(4, 2))
    tk.Button(hl, text="关于", bg=t.modal_fill, fg=t.primary, font=t.sans(TYPE["caption"]),
              relief="flat", cursor="hand2",
              command=lambda: show_about(top)).pack(side="left", padx=14)
    tk.Button(hl, text="使用帮助", bg=t.modal_fill, fg=t.primary, font=t.sans(TYPE["caption"]),
              relief="flat", cursor="hand2",
              command=lambda: show_help(top)).pack(side="left", padx=14)
    _site_label(hl, bg=t.modal_fill).pack(side="left", padx=14)

    tk.Label(top.body, text="未签名程序提示：Windows 可能弹出 SmartScreen 拦截，点击「详细信息」→「仍要运行」即可打开（官网激活教程有图文演示）。",
             bg=t.modal_fill, fg=t.muted, font=t.sans(TYPE["caption"]),
             wraplength=560).pack(pady=(2, 6))
    tk.Label(top.body, text="© 2026 论文格式医生 · 公众号【芦苇不熬夜】 ID：reedskill · 合作联系：hi@reedskill.com",
             bg=t.modal_fill, fg=t.muted, font=t.sans(TYPE["caption"]),
             wraplength=560).pack(pady=(8, 10))

    # 按内容实际所需高度自动伸缩（v1.3.67：固定高在字体缩放/DPI 差异下会截掉底部内容），
    # 仅设「不超过屏幕」上限；canvas 高度随之调整，保证版权/链接完整显示。
    try:
        top.update_idletasks()
        _need = top.body.winfo_reqheight() + 150
        _h = int(min(max(_need, _h), _sh - 100))
        top.cv.configure(height=_h)
        top._height = _h
    except Exception:
        pass

    # 居中（兼容启动时 root 被 withdraw 的情况 → 直接屏幕居中，避免依赖父窗坐标）
    top.update_idletasks()
    _sw2, _sh2 = top.winfo_screenwidth(), top.winfo_screenheight()
    top.geometry("%dx%d+%d+%d" % (_w, _h, max(0, (_sw2 - _w) // 2),
                                   max(0, (_sh2 - _h) // 2)))
    top.wait_window()
    return result["v"]


def _site_label(parent, text=OFFICIAL_SITE_TEXT, bg=PAPER):
    """可点击的官网链接标签：点一下用默认浏览器打开官网。"""
    lbl = tk.Label(parent, text=text, bg=bg, fg="#185FA5", font=F_SMALL, cursor="hand2")
    lbl.bind("<Button-1>", lambda e: webbrowser.open(OFFICIAL_SITE))
    lbl.bind("<Enter>", lambda e: lbl.config(fg="#0c447c"))
    lbl.bind("<Leave>", lambda e: lbl.config(fg="#185FA5"))
    return lbl


# ---------------------------------------------------------------------------
# 关于 / 帮助 窗口（v1.3.68：品牌署名 + 客服入口 + 主打论文安全·离线）
# ---------------------------------------------------------------------------
def _scroll_frame(parent, bg=PAPER):
    """可滚动容器：Canvas + 常驻滚动条，滚轮/拖拽/箭头均可用，内部宽度跟随画布。"""
    cv = tk.Canvas(parent, bg=bg, highlightthickness=0)
    sb = ttk.Scrollbar(parent, orient="vertical", command=cv.yview)
    inner = tk.Frame(cv, bg=bg)
    win = cv.create_window((0, 0), window=inner, anchor="nw")

    def _sync(event=None):
        # 先让内部布局稳定（wraplength 重排等），再按实际内容刷新滚动区域；
        # 否则 bbox 高度滞后，滑块会显示成满格"一条"、拖不动（只能点箭头）。
        try:
            cv.update_idletasks()
        except Exception:
            pass
        cv.configure(scrollregion=cv.bbox("all"))

    def _on_cv(event):
        cv.itemconfig(win, width=event.width)
        _sync()

    def _wheel(event):
        if getattr(event, "num", None) == 4:
            cv.yview_scroll(-3, "units")
        elif getattr(event, "num", None) == 5:
            cv.yview_scroll(3, "units")
        else:
            d = getattr(event, "delta", 0) or 0
            step = -3 if d > 0 else 3
            if abs(d) >= 120:
                step = -int(d / 120) * 3
            cv.yview_scroll(step, "units")

    inner.bind("<Configure>", _sync)
    cv.bind("<Configure>", _on_cv)
    cv.configure(yscrollcommand=sb.set)
    cv.grid(row=0, column=0, sticky="nsew")
    sb.grid(row=0, column=1, sticky="ns")
    parent.rowconfigure(0, weight=1)
    parent.columnconfigure(0, weight=1)
    # 兜底：窗口显示、内容布局稳定后多次刷新滚动范围（滑块大小/上下界限由此决定）
    parent.after(150, _sync)
    parent.after(500, _sync)
    for w in (cv, inner):
        w.bind("<MouseWheel>", _wheel)
        w.bind("<Button-4>", _wheel)
        w.bind("<Button-5>", _wheel)
    return inner


def _auto_wrap(lbl, container, padx, extra=0):
    """让 Label 的换行宽度跟随容器宽度，窗口缩放时文字自动重新排版。"""
    def _fit(event):
        w = event.width - 2 * padx - extra
        if w > 80:
            lbl.configure(wraplength=w)
    container.bind("<Configure>", _fit)
    return lbl


def _hint_fit(lbl, container, title_widget, extra_widgets=(), pad=16):
    """卡片头右侧的能力说明：宽度不够时自动折行，绝不压到左边的标题。"""
    def _fit(event):
        try:
            used = title_widget.winfo_reqwidth()
            for w in extra_widgets:
                used += w.winfo_reqwidth()
        except Exception:
            return
        w = event.width - used - pad * 2
        lbl.configure(wraplength=max(140, w))
    container.bind("<Configure>", _fit, add="+")
    return lbl


def _hint_natural_width(lbl):
    """说明文字不折行时的自然宽度（用于判断「放不放得下」）。"""
    try:
        f = tkfont.Font(font=lbl.cget("font"))
        return int(f.measure(lbl.cget("text")))
    except Exception:
        return int(lbl.winfo_reqwidth())


def _hint_fits(container, title_widget, hint, extra_widgets=(), pad=16):
    """卡片头里「标题 + 说明」能不能并排放下（含最小呼吸间距）。"""
    try:
        used = title_widget.winfo_reqwidth()
        for wdg in extra_widgets:
            used += wdg.winfo_reqwidth()
        avail = container.winfo_width()
    except Exception:
        return True
    if avail < 60:
        return True          # 还没布局完，先不收起，等下一次 Configure 再判
    return avail >= used + _hint_natural_width(hint) + 40


def show_about(parent):
    """关于页：名片式居中排版，细线分隔，文字随窗口自适应。"""
    top = tk.Toplevel(parent)
    top.title("关于 · 论文格式医生")
    top.configure(bg=PAPER)
    top.geometry("565x680")
    top.resizable(True, True)
    top.minsize(520, 560)

    inner = _scroll_frame(top)
    padx = 44

    def wlbl(text, fg, font, pady, nowrap=False, **kw):
        """居中文本标签；nowrap=True 固定单行（短信息），否则换行宽度跟随窗口。"""
        opts = dict(bg=PAPER, fg=fg, font=font, justify="center")
        if not nowrap:
            opts["wraplength"] = 420
        lbl = tk.Label(inner, text=text, **opts)
        if not nowrap:
            _auto_wrap(lbl, inner, padx)
        lbl.pack(padx=padx, pady=pady)
        return lbl

    def rule(pady=(0, 14)):
        tk.Frame(inner, bg=LINE, height=1).pack(fill="x", padx=padx, pady=pady)

    # —— 头部：题名 + 版本 ——
    tk.Label(inner, text=COPY_TITLE, bg=PAPER, fg=INK,
             font=F_TITLE).pack(padx=padx, pady=(24, 2))
    tk.Label(inner, text="版本 v%s" % APP_VERSION, bg=PAPER, fg=MUTED,
             font=F_SUBTITLE).pack(padx=padx, pady=(0, 12))
    rule()

    # —— 定位 ——
    wlbl("以学校模板为准绳，为论文格式把脉。", INK, F_HDR, (0, 8), nowrap=True)
    wlbl("标题层级、字体字号、页边距、行距，参考文献与三线表，逐一对照，改至合乎规范。",
         BODY, F_BODY, (0, 14))
    rule()

    # —— 品牌 + 关注 ——
    wlbl("芦苇不熬夜 出品", INK, F_HDR, (2, 8), nowrap=True)
    try:
        if os.path.isfile(QRCODE):
            qr = tk.PhotoImage(file=QRCODE)
            qr = qr.subsample(max(1, round(qr.width() / 150)))  # 缩到约 150px
            ql = tk.Label(inner, image=qr, bg=PAPER)
            ql.image = qr
            ql.pack(pady=(0, 8))
    except Exception:
        pass
    wlbl("微信公众号：【%s】（ID：%s）" % (WECHAT_NAME, WECHAT_ID), BODY, F_BODY, (0, 2), nowrap=True)
    wlbl("联系邮箱：%s" % ABOUT_MAIL, BODY, F_BODY, (0, 6), nowrap=True)
    wlbl("使用中若有疑问，欢迎关注公众号留言，或致信 %s，我们看到即复。" % ABOUT_MAIL,
         MUTED, F_SMALL, (0, 14))
    _site_label(inner, text="官网 reedskill.com · 产品介绍与激活教程").pack(padx=padx, pady=(0, 6))
    rule()

    # —— 关于本软件 ——
    wlbl("关于本软件", ACCENT, F_HDR, (0, 8))
    wlbl("论文安全 · 所有修正均在本机完成，论文不联网、不上传、不被收集。",
         BODY, F_SMALL, (0, 4))
    wlbl("模板驱动 · 格式要求取自学校模板，批注说明为先，样式定义次之。",
         BODY, F_SMALL, (0, 12))
    wlbl("未签名程序 · 首次打开时 Windows 可能弹出 SmartScreen 拦截，点击「详细信息」→「仍要运行」即可（官网激活教程有图文演示）。",
         MUTED, F_SMALL, (0, 12))

    tk.Label(inner, text="© 2026 芦苇不熬夜", bg=PAPER, fg=MUTED,
             font=F_FOOT).pack(padx=padx, pady=(0, 18))


def show_help(parent):
    """使用帮助：文档式左对齐排版，分区细线，文字随窗口自适应。"""
    top = tk.Toplevel(parent)
    top.title("使用帮助 · 论文格式医生")
    top.configure(bg=PAPER)
    top.geometry("805x680")
    top.resizable(True, True)
    top.minsize(680, 560)

    inner = _scroll_frame(top)
    padx = 40
    wrap = 660

    def albl(text, fg, font, nowrap=False, **kw):
        """左对齐文本标签；nowrap=True 固定单行，否则换行宽度跟随窗口。"""
        opts = dict(bg=PAPER, fg=fg, font=font, justify="left", anchor="w")
        if not nowrap:
            opts["wraplength"] = wrap
        lbl = tk.Label(inner, text=text, **opts)
        if not nowrap:
            _auto_wrap(lbl, inner, padx)
        return lbl

    # —— 头部（居中）——
    tk.Label(inner, text="使用帮助", bg=PAPER, fg=INK,
             font=F_TITLE).pack(padx=padx, pady=(22, 4))
    tk.Label(inner, text="将学校模板告知软件，导入论文后依向导循序而行，格式自可妥帖。",
             bg=PAPER, fg=BODY, font=F_BODY, justify="center").pack(padx=padx, pady=(0, 12))
    tk.Frame(inner, bg=LINE, height=1).pack(fill="x", padx=padx, pady=(0, 8))

    def section(title):
        tk.Label(inner, text=title, bg=PAPER, fg=ACCENT, font=F_HDR,
                 anchor="w").pack(fill="x", padx=padx, pady=(14, 4))
        tk.Frame(inner, bg=LINE, height=1).pack(fill="x", padx=padx, pady=(0, 4))

    def item(title, body):
        tl = albl(title, INK, ("KaiTi", TYPE["card_title"], "bold"))
        tl.pack(fill="x", padx=padx, pady=(8, 1))
        bl = albl(body, BODY, F_SUBTITLE)
        bl.pack(fill="x", padx=padx, pady=(0, 4))

    # —— 一、怎么用 ——
    section("一、怎么用")
    steps = [
        ("1. 导入论文",
         "在主界面「文件选择」中，先选择论文文件（.docx / .doc）。"
         "导入后软件进入向导模式，引领你完成后续各步。"),
        ("2. 导入学校模板",
         "仍在「文件选择」处，选择学校下发的格式要求文件（Word 模板或格式规范文档均可）。"
         "模板若带批注更佳：批注常写明具体要求，如「一级标题用黑体小二、居中」；"
         "软件以批注说明为先，故带批注者修正最贴近学校要求。"),
        ("3. 生成模板画像",
         "软件读取模板中的格式要求，生成「模板画像」——字体、字号、行距、页边距、标题层级等要素。"),
        ("4. 逐步确认",
         "软件分步呈现各项设定（标题层级、字体字号、页边距、缩进、行距、参考文献、三线表），"
         "逐一过目、确认即可。"),
        ("5. 生成报告",
         "确认完毕后，软件依模板修正论文，并生成报告，注明改动之处与未能确定之处。"),
    ]
    for title, body in steps:
        item(title, body)

    # —— 二、常见问题 ——
    section("二、常见问题")
    faqs = [
        ("Q1 · 什么是「模板驱动」？没有模板可用吗？",
         "可用。未上传模板时，软件以通用规范（通用毕业论文格式）兜底；"
         "各校要求不一，上传本校模板，修正方更贴合。"),
        ("Q2 · 什么样的模板最好？一定要带批注吗？",
         "非必带，但带批注的学校官方模板效果最佳。软件定格式时，批注说明优先于样式定义："
         "批注有明确要求则遵批注，无批注则读样式定义，样式亦无则以通用规范兜底。"),
        ("Q3 · 试用版与正式版有何区别？卡片有哪些种类？",
         ("试用版：免费试用 %d 次，输出带水印的只读预览（不可直接编辑）。\n" % trial.TRIAL_LIMIT) +
         "正式版：无水印、可编辑、可保存。卡片共 4 种：\n"
         "· 永久卡 — 一次付费终身可用，激活后完全离线；\n"
         "· 周卡 / 月卡 — 期限内不限次数使用，到期后续费即可继续；\n"
         "· 次卡 — 按次计费，每修正一次扣一次次数（需联网扣次）。"),
        ("Q4 · 试用输出为只读，如何修改？",
         "试用文档设有只读保护；正式版输出可编辑文档，更为便捷，建议直接激活正式版。"),
        ("Q5 · 如何激活？激活后还需要联网吗？",
         "在激活窗口输入购买的激活码，验证通过即为正式版，右上角会显示您的卡片种类与剩余期限/次数。\n"
         "· 永久卡 — 激活后完全离线，无需再联网；\n"
         "· 周卡 / 月卡 — 激活后离线可用，到期前会联网校验一次；\n"
         "· 次卡 — 每次修正需联网扣一次次数，离线时暂不可用。"),
         ("Q6 · 论文安全吗？会上传吗？",
          "请放心。所有修正均在本机完成，论文不联网、不上传任何服务器，亦不收集论文内容。\n"
          "联网仅用于「激活校验」与「首次试用登记」，且只传输激活码 / 机器码，绝不含论文内容。\n"
          "永久卡 / 周卡 / 月卡激活后可断网使用；次卡仅在扣减次数时联网。"),
        ("Q7 · 学校模板特殊，或修正结果不尽如人意？",
         "欢迎关注公众号【%s】留言，告知贵校情况，我们协助处理。" % WECHAT_NAME),
        ("Q8 · 打开软件时 Windows 弹出“已保护你的电脑 / 已拦截”提示？",
         "本软件为未签名程序（省去每年数百元代码签名证书费用，让价格更亲民），Windows SmartScreen 会拦截提示，属正常现象，不代表软件有害。\n"
         "处理方式：在拦截框点击【详细信息】→ 再点击【仍要运行】即可正常打开；官网 reedskill.com 的“激活教程”页有图文演示。"),
    ]
    for title, body in faqs:
        item(title, body)

    # —— 三、联系我们 ——
    section("三、联系我们")
    cl = albl(("使用疑问、模板咨询、购买事宜，均可通过以下方式联系：\n"
               "· 微信小程序：微信搜【%s】，可直接购买激活码，付款后自动发码（推荐）\n"
               "· 微信公众号：【%s】（ID：%s），关注后留言\n"
               "· 官网：reedskill.com（产品介绍 · 激活教程 · 购买方式）\n"
               "· 公众号设关键词自动回复（试用 / 激活 / 水印 / 报错），常见问题即时应答\n"
               "· 较复杂的问题，由人工回复\n"
               "· 联系邮箱：%s") % (MINIAPP_NAME, WECHAT_NAME, WECHAT_ID, ABOUT_MAIL),
              BODY, F_SUBTITLE)
    cl.pack(fill="x", padx=padx, pady=(8, 0))

    # 小程序码（购买主渠道）：图片缺失时自动跳过，不影响其余内容显示
    try:
        if os.path.isfile(MINIAPP_QRCODE):
            _mq = tk.PhotoImage(file=MINIAPP_QRCODE)
            _mq = _mq.subsample(max(1, round(_mq.width() / 130)))
            _mrow = tk.Frame(inner, bg=PAPER)
            _mrow.pack(pady=(10, 0))
            _mimg = tk.Label(_mrow, image=_mq, bg=PAPER)
            _mimg.image = _mq
            _mimg.pack()
            tk.Label(_mrow, text="微信扫一扫，或搜索【%s】小程序" % MINIAPP_NAME,
                     bg=PAPER, fg=MUTED, font=F_SMALL).pack(pady=(6, 0))
    except Exception:
        pass

    _site_label(inner, text="前往官网 reedskill.com →").pack(padx=padx, pady=(8, 0))

    tk.Label(inner, text="", bg=PAPER).pack(pady=(0, 18))


def main():
    # Windows 高分屏下让 Tk 字体清晰渲染（必须在创建 Tk 窗口前设置，否则整体发糊）
    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
    root = tk.Tk()
    _init_fonts(root)
    root.withdraw()
    if (not license.check_local_valid()) or license.is_revoked():
        # v1.3.58：未激活（或被后台撤销）时弹激活窗，但提供"先试用"入口（不进主界面则退出）
        r = show_activation(root)
        if r not in ("ok", "trial"):
            root.destroy()
            return
    else:
        # 启动实时校验：本地有效时再联网确认一次（退款撤销/到期/次卡用尽 → 拦；断网放行）
        st = license.validate_now()
        if st in ("revoked", "expired"):
            title, body = _revoked_message(license.get_lock_reason())
            root.deiconify()
            messagebox.showerror(title, body, parent=root)
            root.destroy()
            return
    root.deiconify()
    App(root)
    # 启动后台心跳：联网时若授权被撤销（退款），自动锁死软件
    if license.check_local_valid() and not license.is_revoked():
        license.start_heartbeat(product=license.DEFAULT_PRODUCT, on_revoked=_on_license_revoked)
    root.mainloop()


if __name__ == "__main__":
    main()
