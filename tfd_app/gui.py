# -*- coding: utf-8 -*-
"""
论文格式医生 · 原生 Tkinter 界面（客户安心版 v2）
==============================================
面向客户的简洁流程，不展示任何执行细节；一屏放得下，无需滚动：

  壹 · 选择文件        —— 待处理论文（必选）+ 学校模板（可选）
  贰 · 按步骤操作      —— ① 提取学校模板要求 → ② 论文格式检查 → ③ 按学校要求一键修正
  叁 · 处理状态        —— 只显示友好状态与结果确认，无原始日志

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

from . import trial       # v1.3.56：试用计数（机器码绑定，2 次）——相对导入，PyInstaller 才收集
from . import watermark   # v1.3.56：试用水印（页眉页脚+正文穿插）
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import tkinter.font as tkfont

from . import engine, license


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
# 配色：宣纸 / 墨 / 黛蓝 / 朱砂 —— 学术范 + 文艺感
# ---------------------------------------------------------------------------
PAPER   = "#f6f2ea"   # 宣纸底（保留用户喜欢的底色）
PANEL   = "#fdfbf6"   # 面板（微亮的纸）
INK     = "#222222"   # 主文字（深）—— 标题 / 重点
BODY    = "#444444"   # 正文（灰）—— 描述性文字
MUTED   = "#666666"   # 次要文字 / 页脚（浅灰）
LINE    = "#e0d9c8"   # 细线
ACCENT  = "#46586f"   # 黛蓝（章节条 / 强调）
CINNABAR= "#9e4233"   # 朱砂（主按钮）
CINNABAR_D = "#8a382b"
OKC     = "#5f7d5c"   # 完成（墨绿）
ERRC    = "#a0402f"   # 出错（朱红）
RUN     = ACCENT      # 处理中（黛蓝）

# 字体：标题 / 章节 / 状态用楷体（KaiTi，文艺调性），正文 / 按钮 / 页脚用微软雅黑（清晰）
# 运行时由 _init_fonts() 变成 tkfont.Font 命名对象（name -> Font），随窗口大小整体缩放，
# 任何 widget / ttk style 引用 F_XXX 都会自动级联更新，保证"窗口拉大内容跟着变大"。
_FONT_BASE = {
    "F_TITLE":      ("KaiTi", 20, "bold"),      # 主标题（楷体）
    "F_SUB":        ("Microsoft YaHei", 12),    # 副标题
    "F_HDR":        ("KaiTi", 13, "bold"),      # 章节标题（楷体）
    "F_BODY":       ("Microsoft YaHei", 12),    # 正文 / 步骤说明
    "F_SMALL":      ("Microsoft YaHei", 10),    # 底部提示 / 次要
    "F_SMALL_B":    ("Microsoft YaHei", 10, "bold"),  # 提示框标题（加粗，随缩放联动）
    "F_BTN":        ("Microsoft YaHei", 12, "bold"),  # 按钮
    "F_STAT":       ("KaiTi", 12),              # 状态文字（楷体）
    "F_SUBTITLE":   ("Microsoft YaHei", 11),    # 元信息
    "F_CARD_HDR":   ("KaiTi", 13, "bold"),      # 卡片标题（楷体）
    "F_FOOT":       ("Microsoft YaHei", 10),    # 状态栏 / 页脚 / 说明（次要，最小层级）
    "F_MONO":       ("Consolas", 9),            # 机器码 / 离线码（等宽）
    "F_DIALOG_TITLE": ("KaiTi", 14, "bold"),    # 弹窗标题（楷体）
    "F_ICON":       ("KaiTi", 12, "bold"),      # 印章图标（论 / 模，楷体朱砂）
}
APP_VERSION = "1.3.116"   # 与 VERSION 文件保持同步（状态栏显示用）

# v1.3.108：绿色 zip 版由软件自建桌面快捷方式（win32com 已内置，客户零依赖、零黑框）。
APP_SHORTCUT_NAME = "论文格式医生"       # 桌面快捷方式显示名
SHORTCUT_FLAG_TAG = "ThesisFormatDoctor"  # %APPDATA% 下询问标记目录（区分学生/导师版）


def _is_frozen_exe():
    """打包后的主程序才显示“桌面图标”入口 / 首启询问；
    开发态（python.exe/pythonw.exe）与 mac/Linux 一律不建 Windows 快捷方式。"""
    base = os.path.basename(sys.executable or "").lower()
    return base.endswith(".exe") and not (base.startswith("python")
                                          or base.startswith("pythonw"))


def _btn_display_width(text, pad=2):
    """按钮文案的「显示宽度」：全角字符按 2、半角按 1 累加，再加左右余量。

    ttk.Button 的 width 单位是「英文字符宽」，而中/日/韩全角字符渲染宽度约为
    英文的 2 倍。直接拿 len(text) 当宽度会让含中文的文案被截断（v1.3.98 的 bug：
    "上一步" 3 字只分到 4 字符宽 → 只显示得出 2 个汉字）。
    """
    w = 0
    for ch in text:
        # 常用全角区间：CJK 统一表意文字、假名、谚文、全角 ASCII/标点、CJK 符号
        o = ord(ch)
        if (0x1100 <= o <= 0x115F or 0x2E80 <= o <= 0xA4CF
                or 0xAC00 <= o <= 0xD7A3 or 0xF900 <= o <= 0xFAFF
                or 0xFE30 <= o <= 0xFE6F or 0xFF00 <= o <= 0xFF60
                or 0xFFE0 <= o <= 0xFFE6 or 0x20000 <= o <= 0x3FFFD):
            w += 2
        else:
            w += 1
    return max(6, w + pad)
_FONTS = {}      # name -> (Font, base_size)
_CUR_SCALE = 1.0 # 当前窗口缩放比例（宽度 / 基准宽度，钳制 0.8~1.0：只缩小不放大）
BASE_W = 900     # 设计基准宽度（px），与主窗口默认 900x640 对应

def _init_fonts(root):
    """在 root 创建后调用：把 F_XXX 全局名替换为可缩放的 Font 命名对象。"""
    import platform as _platform
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
    for name, spec in _FONT_BASE.items():
        kw = {"family": _resolve(spec[0]), "size": spec[1]}
        if len(spec) > 2:
            kw["weight"] = spec[2]
        f = tkfont.Font(root=root, **kw)
        _FONTS[name] = (f, spec[1])
        globals()[name] = f

def _apply_scale(factor):
    """按宽度因子缩放全部字体字号（下限 9pt 保证可读；上限 1.0：窗口拉大不放大字体）。"""
    global _CUR_SCALE
    factor = max(0.8, min(1.0, factor))
    _CUR_SCALE = factor
    for f, base in _FONTS.values():
        f.configure(size=max(9, int(round(base * factor))))

def _geo(w, h):
    """弹窗固定尺寸按当前缩放比例换算，避免字体变大后内容溢出。"""
    s = _CUR_SCALE
    return "%dx%d" % (max(420, int(w * s)), max(380, int(h * s)))

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

# v1.3.116：手动填写格式（与学校模板互斥）——候选常量
_MANUAL_ZH_FONTS = ["宋体", "黑体", "仿宋", "楷体", "微软雅黑"]
_MANUAL_EN_FONTS = ["Times New Roman", "Arial", "Calibri"]
_MANUAL_SIZES = ["初号", "小初", "一号", "小一", "二号", "小二", "三号", "小三",
                 "四号", "小四", "五号", "小五", "六号", "小六", "七号", "八号"]
_MANUAL_ALIGNS = ["居中", "左对齐", "右对齐", "两端对齐"]
_MANUAL_LINE_TYPES = ["单倍行距", "1.5倍行距", "2倍行距", "固定值(磅)", "最小值(磅)"]


def _build_manual_form():
    """手动填写格式表单的字段定义：[(字段显示名, key, 控件类型, 灰字说明)]。

    控件类型：zh_font/en_font=可手输下拉；size=纯下拉(不可手输)；align=纯下拉；
    bold=勾选；num=数字输入；line_type=纯下拉；line_val=数字输入。
    key=None 表示这是一个分区标题行（不生成输入控件）。

    分区标题仅做层级分组（① 一级标题 / ② 二级标题 / ③ 三级标题 / 正文 / 摘要·标题
    / 摘要·正文 / 参考文献·标题 / 参考文献·条目 / 页面边距）；每个字段的显示名写「具体
    类目」（中文字体/英文字体/字号…），不再重复写分区标题，层级一眼看清。
    """
    F = []
    def sec(n):
        F.append((n, None, None, None))
    # 标题层级：固定 3 级，每级字段相同，仅分区标题区分级别
    _HEAD = (
        ("zh_font", "中文字体", "中文如 黑体"),
        ("en_font", "英文字体", "英文如 Times New Roman"),
        ("size", "字号", "如 小二"),
        ("align", "对齐方式", ""),
        ("bold", "加粗", "是否加粗"),
        ("before", "段前间距", "单位：磅"),
        ("after", "段后间距", "单位：磅"),
        ("line_type", "行距类型", ""),
        ("line_val", "行距值", "固定/最小值时填磅"),
    )
    for lk, lbl in (("1", "① 一级标题"), ("2", "② 二级标题"), ("3", "③ 三级标题")):
        sec(lbl)
        for k, flbl, hint in _HEAD:
            F.append((flbl, "%s_%s" % (lk, k), k, hint))
    # 正文
    sec("正文")
    for k, flbl, hint in (
        ("zh_font", "中文字体", "中文如 宋体"),
        ("en_font", "英文字体", "英文如 Times New Roman"),
        ("size", "字号", "如 小四"),
        ("align", "对齐方式", ""),
        ("line_type", "行距类型", ""),
        ("line_val", "行距值", "固定/最小值时填磅"),
        ("indent", "首行缩进", "单位：字符"),
    ):
        F.append((flbl, "body_%s" % k, k, hint))
    # 摘要
    sec("摘要·标题")
    for k, flbl, hint in (
        ("zh_font", "中文字体", "中文如 黑体"),
        ("en_font", "英文字体", "英文如 Times New Roman"),
        ("size", "字号", "如 小二"),
    ):
        F.append((flbl, "abs_t_%s" % k, k, hint))
    sec("摘要·正文")
    for k, flbl, hint in (
        ("zh_font", "中文字体", "中文如 宋体"),
        ("en_font", "英文字体", "英文如 Times New Roman"),
        ("size", "字号", "如 小四"),
        ("line_type", "行距类型", ""),
        ("line_val", "行距值", "固定/最小值时填磅"),
    ):
        F.append((flbl, "abs_%s" % k, k, hint))
    # 参考文献
    sec("参考文献·标题（引用格式默认 GB/T 7714 顺序编码制）")
    for k, flbl, hint in (
        ("zh_font", "中文字体", "中文如 黑体"),
        ("en_font", "英文字体", "英文如 Times New Roman"),
        ("size", "字号", "如 小四"),
    ):
        F.append((flbl, "ref_t_%s" % k, k, hint))
    sec("参考文献·条目")
    for k, flbl, hint in (
        ("zh_font", "中文字体", "中文如 宋体"),
        ("en_font", "英文字体", "英文如 Times New Roman"),
        ("size", "字号", "如 小四"),
        ("line_type", "行距类型", ""),
        ("line_val", "行距值", "固定/最小值时填磅"),
    ):
        F.append((flbl, "ref_i_%s" % k, k, hint))
    # 参考文献条目悬挂缩进：按编号位数分三档（字符），与引擎 _ref_entry_spec_for_text 对齐
    for dig, flbl in ((1, "1-9 条·悬挂缩进"), (2, "10-99 条·悬挂缩进"), (3, "100-999 条·悬挂缩进")):
        F.append((flbl, "ref_i_hanging_%d" % dig, "num", "单位：字符"))
    # 页面边距
    sec("页面边距（留空用通用规范 2.5 厘米）")
    for k, lbl, hint in (("top", "上", "厘米"), ("bottom", "下", "厘米"),
                         ("left", "左", "厘米"), ("right", "右", "厘米")):
        F.append((lbl, "page_%s" % k, "num", hint))
    return F


_MANUAL_FORM = _build_manual_form()


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
        # 2) 最小宽度 1040：50:50 等分后每栏仍有约 500px，按钮/文件名不被挤变形；
        # 3) 最小高度 660：保证底部按钮区完整可见；
        # 4) Windows 且屏幕 ≥1440×900 时启动即最大化（小屏跳过，保留窗口控制权）。
        _sw = self.root.winfo_screenwidth()
        _sh = self.root.winfo_screenheight()
        self.root.geometry("%dx%d" % (min(1180, int(_sw * 0.86)),
                                      min(880, int(_sh * 0.88))))
        self.root.minsize(1040, 660)
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

        # 随窗口缩放自适应：监听尺寸变化，按宽度比例整体缩放字体
        self._last_scale = None
        self.root.bind("<Configure>", self._on_resize)

        self.thesis_path = tk.StringVar()
        self.template_path = tk.StringVar()
        self.profile_path = tk.StringVar()
        self.status_var = tk.StringVar(value="请按步骤操作")
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

        # v1.1.10：第①步输入方式互斥（学校模板 / 手动填写格式）——二选一，不可混用
        self.input_mode = tk.StringVar(value="template")
        self.manual_profile_path = ""      # 手动填写生成的画像 json 路径
        self._manual_form_values = {}      # 已填写表单原始值（用于预填 / 记住）
        self._manual_remember = False      # 是否记住手动配置
        self._manual_filled = False        # 是否已填写手动格式
        self._manual_frame = None          # 手动填写区（按钮 + 状态）
        self._manual_btn = None            # 手动填写按钮
        self._manual_status = None         # 已填写状态标签

        self._build_style()
        self._build_widgets()
        self._restore_manual_config()   # 已记住手动配置则自动载入

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
        factor = w / BASE_W
        if self._last_scale is not None and abs(factor - self._last_scale) < 0.02:
            return
        self._last_scale = factor
        _apply_scale(factor)

    # ------------------------------------------------------------- style
    def _build_style(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("TButton", font=F_BODY, padding=(12, 6),
                        background="#eae3d3", foreground=INK)
        style.map("TButton", background=[("active", "#ded6c2"), ("disabled", "#f0ece1")])
        style.configure("Action.TButton", font=F_BTN, padding=(12, 10),
                        background="#eae3d3", foreground=INK)
        style.map("Action.TButton",
                  background=[("active", "#ded6c2"), ("disabled", "#f0ece1")])
        style.configure("Primary.TButton", font=F_BTN, padding=(12, 10),
                        background=CINNABAR, foreground="white")
        style.map("Primary.TButton",
                  background=[("active", CINNABAR_D), ("disabled", "#c8a79f")],
                  foreground=[("disabled", "#f5e9e6")])
        style.configure("Ghost.TButton", font=F_SMALL, padding=(8, 5),
                        background=PANEL, foreground=ACCENT)
        style.map("Ghost.TButton",
                  background=[("active", "#efe9db"), ("disabled", "#f4f0e6")])

    # ------------------------------------------------------------- layout
    def _build_widgets(self):
        self.root.configure(bg=PAPER)

        # 右上角文字链接：桌面图标 ｜ 关于 ｜ 帮助（黛蓝楷体，文艺学术风）
        topbar = tk.Frame(self.root, bg=PAPER)
        topbar.pack(fill="x", padx=20, pady=(10, 0))
        tk.Label(topbar, text="", bg=PAPER).pack(side="left", expand=True)

        def _link(parent, text, cmd):
            lbl = tk.Label(parent, text=text, bg=PAPER, fg=ACCENT,
                           font=("KaiTi", 13), cursor="hand2")
            lbl.bind("<Button-1>", lambda e: cmd())
            lbl.bind("<Enter>", lambda e: lbl.config(fg="#33465c"))
            lbl.bind("<Leave>", lambda e: lbl.config(fg=ACCENT))
            return lbl

        _link(topbar, "帮 助", lambda: show_help(self.root)).pack(side="right")
        tk.Label(topbar, text="｜", bg=PAPER, fg="#c9c0ae",
                 font=("Microsoft YaHei", 12)).pack(side="right", padx=(0, 8))
        _link(topbar, "关 于", lambda: show_about(self.root)).pack(side="right", padx=(0, 8))
        # v1.3.108：绿色 zip 版一键补建桌面快捷方式（win32com 已内置 exe，零外部依赖、零黑框）
        if sys.platform.startswith("win") and _is_frozen_exe():
            tk.Label(topbar, text="｜", bg=PAPER, fg="#c9c0ae",
                     font=("Microsoft YaHei", 12)).pack(side="right", padx=(0, 8))
            _link(topbar, "桌面图标", self._create_desktop_shortcut).pack(side="right", padx=(0, 8))

        # 顶部标题区
        header = tk.Frame(self.root, bg=PAPER)
        header.pack(fill="x", padx=20, pady=(16, 8))
        tk.Label(header, text="论 文 格 式 医 生", bg=PAPER, fg=INK,
                 font=F_TITLE).pack(anchor="center")
        tk.Label(header, text="THESIS FORMAT DOCTOR · 高校论文格式规范引擎",
                 bg=PAPER, fg="#8b8378", font=F_SUBTITLE).pack(anchor="center", pady=(5, 0))
        tk.Frame(self.root, bg=CINNABAR, height=2).pack(fill="x", padx=20)

        # 主体两栏（文件选择 | 处理步骤）
        # v1.3.90：左右严格 50:50 等分（与导师版一致）。
        # 关键：仅靠 weight 做不到等宽——grid 先满足各列"固有请求宽度"，导入论文后
        # 左栏出现文件名/路径、请求宽度变大就会多吃，右栏被挤窄，两栏比例跟着跳动。
        # 必须用 uniform 把两列归为同一尺寸组，tk 才会强制两列等宽；minsize 是小窗口兜底。
        main = tk.Frame(self.root, bg=PAPER)
        main.pack(fill="both", expand=True, padx=20, pady=14)
        main.columnconfigure(0, weight=1, uniform="half", minsize=470)
        main.columnconfigure(1, weight=1, uniform="half", minsize=470)
        main.rowconfigure(0, weight=1)
        left = tk.Frame(main, bg=PAPER)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        right = tk.Frame(main, bg=PAPER)
        right.grid(row=0, column=1, sticky="nsew")
        self._build_left(left)
        self._build_right(right)

        # 页脚（两行版权，贴底）
        footer = tk.Frame(self.root, bg=PAPER)
        footer.pack(side="bottom", fill="x", pady=(0, 18))
        tk.Label(footer, text="论文格式医生 · 桌面版 — 本机处理，文件不会上传任何服务器",
                 bg=PAPER, fg=MUTED, font=F_FOOT).pack()
        f2 = tk.Frame(footer, bg=PAPER)
        f2.pack(pady=(3, 0))
        tk.Label(f2,
                 text="© 2026 论文格式医生 · 公众号【芦苇不熬夜】 ID：reedskill · 合作联系：hi@reedskill.com",
                 bg=PAPER, fg=MUTED, font=F_FOOT).pack(side="left")
        _site_label(f2).pack(side="left", padx=(6, 0))

        # 状态栏（顶部细线 + 单行，绝不与其他文字重叠）
        tk.Frame(self.root, bg=LINE, height=1).pack(fill="x", side="bottom", padx=20)
        statusbar = tk.Frame(self.root, bg=PANEL)
        statusbar.pack(side="bottom", fill="x", padx=20, pady=(6, 6))
        self.bar_dot = tk.Label(statusbar, text="●", bg=PANEL, fg=OKC, font=F_FOOT)
        self.bar_dot.pack(side="left", padx=(0, 6))
        self.bar_left = tk.Label(statusbar, text="论文格式医生 · 桌面版",
                                 bg=PANEL, fg=MUTED, font=F_FOOT)
        self.bar_left.pack(side="left")
        self.bar_right = tk.Label(statusbar, text="请按步骤操作",
                                  bg=PANEL, fg=MUTED, font=F_FOOT)
        self.bar_right.pack(side="right")

        self._set_bar("idle")

    # ---------------------------------------------------------- left panel
    def _build_left(self, parent):
        # 文件选择卡片
        card = tk.Frame(parent, bg=PANEL, highlightthickness=1, highlightbackground=LINE)
        card.pack(fill="both", expand=True)

        hdr = tk.Frame(card, bg=PANEL)
        hdr.pack(fill="x", padx=14, pady=(12, 6))
        tk.Frame(hdr, bg=ACCENT, width=4, height=15).pack(side="left", padx=(0, 7))
        tk.Label(hdr, text="文件选择", bg=PANEL, fg=INK, font=F_CARD_HDR).pack(side="left")
        tk.Label(hdr, text="本机离线", bg="#e8f0f6", fg=ACCENT, font=F_FOOT,
                 padx=8, pady=2).pack(side="right")

        self._thesis_box, self._thesis_name = self._file_row(
            card, "论", "待处理论文", "必选", CINNABAR,
            "Word 文档 .docx / .doc / .wps", "选择…", self._pick_input)
        self._template_box, self._template_name = self._file_row(
            card, "模", "学校模板", "可选", MUTED,
            "用于按学校要求检查 / 修正，更贴合要求", "选择…", self._pick_template)

        # v1.1.10：第①步输入方式互斥（学校模板 / 手动填写格式）——二选一，不可混用
        self._mode_frame = tk.Frame(card, bg=PANEL)
        self._mode_frame.pack(fill="x", padx=14, pady=(2, 2))
        tk.Label(self._mode_frame, text="格式来源", bg=PANEL, fg=INK, font=F_SMALL).pack(side="left", padx=(0, 6))
        tk.Radiobutton(self._mode_frame, text="使用学校模板", variable=self.input_mode,
                       value="template", bg=PANEL, fg=ACCENT, font=F_SMALL_B,
                       activebackground=PANEL, command=self._set_input_mode).pack(side="left", padx=(2, 8))
        tk.Radiobutton(self._mode_frame, text="手动填写格式", variable=self.input_mode,
                       value="manual", bg=PANEL, fg=ACCENT, font=F_SMALL_B,
                       activebackground=PANEL, command=self._set_input_mode).pack(side="left", padx=(2, 0))
        # 优先级声明（压一行）
        tk.Label(self._mode_frame, text="格式优先级：模板批注 ＞ 样式 ＞ 通用规范；无模板时手动填写",
                 bg=PANEL, fg="#9a9486", font=F_FOOT).pack(side="left", padx=(10, 0))

        # 手动填写区（默认隐藏；切到“手动填写格式”时显示，同时隐藏模板区）
        self._manual_frame = tk.Frame(card, bg="#fdf3e7", highlightthickness=1, highlightbackground="#e6c794")
        self._manual_btn = tk.Button(self._manual_frame, text="手动填写格式要求 ›",
                                     bg="#e3a93b", fg="#5a3d12", relief="flat",
                                     activebackground="#f0bf5e", activeforeground="#5a3d12",
                                     font=F_SMALL, padx=10, pady=4, cursor="hand2",
                                     command=self._open_manual_dialog)
        self._manual_btn.pack(side="left", padx=10, pady=8)
        self._manual_status = tk.Label(self._manual_frame, text="", bg="#fdf3e7", fg="#7a4e0e", font=F_SMALL)
        self._manual_status.pack(side="left", padx=(0, 10), pady=8)
        self._manual_frame.pack_forget()

        # v1.3.46：模板驱动说明常驻提示（防止客户上传无批注模板造成误解，减少纠纷）
        # v1.3.50：升级为标准 Info 提示框——圆角浅色容器 + 加粗标题 + 说明文字行距 1.6
        def _round_pts(x1, y1, x2, y2, r):
            return [x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r, x2, y2 - r, x2, y2,
                    x2 - r, y2, x1 + r, y2, x1, y2, x1, y2 - r, x1, y1 + r, x1, y1]

        tpl_canvas = tk.Canvas(card, bg=PANEL, highlightthickness=0)
        tpl_canvas.pack(fill="x", padx=14, pady=(4, 2))
        self._tpl_info = tpl_canvas   # v1.1.10：手动模式时与模板区一起隐藏
        tpl_body = tk.Frame(tpl_canvas, bg="#fdf3e7")
        tk.Label(tpl_body, text="模板驱动：以学校模板【批注】写明的格式要求为准（批注优先于样式定义）",
                 bg="#fdf3e7", fg="#7a4e0e", font=F_SMALL_B,
                 justify="left", anchor="w").pack(fill="x", padx=14, pady=(9, 0))
        # 说明文字：Text 的 spacing2 即"段内行距"，实现约 1.6 倍行高；relief=flat 无边框
        # 注意：tk.Text 的 pady 只接受单值（不支持 (0,9) 元组），下边距放 pack 上
        tpl_note_txt = tk.Text(tpl_body, wrap="word", bg="#fdf3e7", fg="#8a5a1a",
                               font=F_FOOT, relief="flat", bd=0, height=3,
                               spacing1=5, spacing2=5, spacing3=5,
                               padx=14, highlightthickness=0, cursor="arrow")
        tpl_note_txt.insert("1.0", "请优先使用学校官方模板（通常批注中写明了格式要求）；"
                                   "若模板无批注，将按模板样式定义 / 通用规范处理，"
                                   "可能与学校要求有出入。")
        tpl_note_txt.config(state="disabled")
        tpl_note_txt.pack(fill="x", pady=(0, 9))
        _tpl_rect = tpl_canvas.create_polygon([0, 0, 20, 20], smooth=True,
                                              fill="#fdf3e7", outline="#e6c794")
        _tpl_win = tpl_canvas.create_window(1, 1, window=tpl_body, anchor="nw")

        def _resize_tpl(_e=None):
            w = tpl_canvas.winfo_width()
            if w <= 2:
                return
            tpl_canvas.itemconfig(_tpl_win, width=w - 4)
            h = tpl_body.winfo_reqheight()
            tpl_canvas.coords(_tpl_rect, *_round_pts(1, 1, w - 2, h + 2, 10))
            tpl_canvas.configure(height=h + 4)

        tpl_canvas.bind("<Configure>", _resize_tpl)
        tpl_body.bind("<Configure>", _resize_tpl)
        tpl_canvas.after(20, _resize_tpl)

        # 底部选择状态
        # 选择状态：论文 / 模板 上下两行、左对齐；选中后打对号
        self._thesis_row = tk.Frame(card, bg=PANEL)
        self._thesis_row.pack(fill="x", padx=14, pady=(10, 1))
        self._thesis_dot = tk.Label(self._thesis_row, text="○", bg=PANEL, fg=MUTED, font=F_FOOT)
        self._thesis_dot.pack(side="left", padx=(0, 6))
        self._thesis_lbl = tk.Label(self._thesis_row, text="未选择论文", bg=PANEL, fg=MUTED,
                                    font=F_FOOT)
        self._thesis_lbl.pack(side="left")
        self._tpl_row = tk.Frame(card, bg=PANEL)
        self._tpl_row.pack(fill="x", padx=14, pady=(1, 12))
        self._tpl_dot = tk.Label(self._tpl_row, text="○", bg=PANEL, fg=MUTED, font=F_FOOT)
        self._tpl_dot.pack(side="left", padx=(0, 6))
        self._tpl_lbl = tk.Label(self._tpl_row, text="模板未选（可选）", bg=PANEL, fg=MUTED,
                                 font=F_FOOT)
        self._tpl_lbl.pack(side="left")

        # 画像状态（提取后显示）
        self.profile_box = tk.Frame(card, bg="#f0f3ec",
                                    highlightthickness=1, highlightbackground="#b7c6ae")
        tk.Label(self.profile_box, text="●", bg="#f0f3ec", fg=OKC,
                 font=F_FOOT).pack(side="left", padx=(10, 4), pady=6)
        self.profile_info_var = tk.StringVar(value="已载入模板格式")
        tk.Label(self.profile_box, textvariable=self.profile_info_var, bg="#f0f3ec",
                 fg="#4c5f49", font=F_FOOT).pack(side="left", fill="x", expand=True)
        ttk.Button(self.profile_box, text="清除", width=6, style="Ghost.TButton",
                   command=self._clear_profile).pack(side="right", padx=8, pady=3)
        self._update_profile_box()

        # 卡片底部章节小字（文艺学术点缀）
        tk.Label(card, text="壹 · 选择", bg=PANEL, fg="#b8b0a0",
                 font=F_FOOT).pack(side="bottom", pady=(0, 8))

    def _file_row(self, parent, icon, title, mark, mark_color, desc, btn_text, cmd):
        box = tk.Frame(parent, bg="#ffffff", highlightthickness=1, highlightbackground="#e3dccb")
        box.pack(fill="x", padx=14, pady=6)
        row = tk.Frame(box, bg="#ffffff")
        row.pack(fill="x", padx=10, pady=8)
        ic = tk.Label(row, text=icon, bg="#ffffff", fg=CINNABAR,
                      font=F_ICON, padx=7, pady=4,
                      highlightthickness=1, highlightbackground=CINNABAR)
        ic.pack(side="left", padx=(0, 8))
        txt = tk.Frame(row, bg="#ffffff")
        txt.pack(side="left", fill="x", expand=True)
        tl = tk.Frame(txt, bg="#ffffff")
        tl.pack(fill="x")
        tk.Label(tl, text=title, bg="#ffffff", fg=INK, font=F_SUBTITLE).pack(side="left")
        tk.Label(tl, text=" " + mark, bg="#ffffff", fg=mark_color, font=F_FOOT).pack(side="left")
        name_lbl = tk.Label(txt, text=desc, bg="#ffffff", fg=MUTED, font=F_FOOT)
        name_lbl.pack(anchor="w")
        ttk.Button(row, text=btn_text, style="Ghost.TButton", command=cmd).pack(side="right")
        return box, name_lbl

    # --------------------------------------------------------- right panel
    def _build_right(self, parent):
        card = tk.Frame(parent, bg=PANEL, highlightthickness=1, highlightbackground=LINE)
        card.pack(fill="both", expand=True)

        hdr = tk.Frame(card, bg=PANEL)
        hdr.pack(fill="x", padx=14, pady=(12, 4))
        tk.Frame(hdr, bg=ACCENT, width=4, height=15).pack(side="left", padx=(0, 7))
        tk.Label(hdr, text="处理步骤", bg=PANEL, fg=INK, font=F_CARD_HDR).pack(side="left")
        self._step_counter = tk.Label(hdr, text="1 / 3", bg=PANEL, fg=MUTED, font=F_FOOT)
        self._step_counter.pack(side="right")

        self._build_timeline(card)

        self.progress = ttk.Progressbar(card, mode="indeterminate")
        self.progress.pack(fill="x", padx=14, pady=(4, 2))
        self.progress.pack_forget()

        # v1.3.100：把 status 行包成子 Frame——v1.3.99 错误地把 trial_btn 长文案
        # 直接 pack 到 card 上，配合 status_lbl fill="x" expand=True 时，会把
        # side="right" 的 trial_btn 挤到 btn_row 下方（实测 y=383 在 btn_row 之下），
        # 导致最后一步布局错位。子 Frame 内 status_lbl fill x expand 只影响本行。
        status_row = tk.Frame(card, bg=PANEL)
        status_row.pack(fill="x")
        self.status_dot = tk.Label(status_row, text="●", bg=PANEL, fg=MUTED, font=F_BODY)
        self.status_dot.pack(side="left", padx=(14, 6), pady=(8, 4))
        self.status_lbl = tk.Label(status_row, textvariable=self.status_var, bg=PANEL, fg=INK,
                                   font=F_STAT)
        # fill="x" expand=True 让 status 文案吃掉中间空间，右边留给 trial_btn
        self.status_lbl.pack(side="left", fill="x", expand=True, pady=(8, 4))

        # v1.3.58：激活/试用状态与入口（未激活=试用版，可随时点激活）
        self._licensed = trial.is_licensed()
        # v1.3.100：trial_btn 不预设固定 width——v1.3.99 实测固定 width 在
        # 中文长文案（"正式版 · 次卡（剩余 3/10 次）✓"）下会被 ttk style 撑爆；
        # 现在 trial_btn 已移入 status_row 子 Frame，按内容自适应不会污染外层 card
        # 的 pack 顺序。badge 由 license.license_summary() 控制 ≤8 字，试用版
        # 文案（"试用版（剩 X 次）· 激活"）由 _update_trial_badge() 写入，
        # 实际宽度按渲染自适。
        self._trial_btn = ttk.Button(status_row, text="", style="Ghost.TButton",
                                     command=self._open_activation)
        self._trial_btn.pack(side="right", padx=(6, 14), pady=(6, 2))
        self._update_trial_badge()

        btn_row = tk.Frame(card, bg=PANEL)
        btn_row.pack(fill="x", padx=14, pady=(10, 14))
        self._prev_btn = ttk.Button(btn_row, text="上一步", style="TButton",
                                    command=self._go_prev)
        self._prev_btn.pack(side="left")
        self._next_btn = ttk.Button(btn_row, text="下一步", style="Primary.TButton",
                                    command=self._run_step)
        self._next_btn.pack(side="right")
        # v1.3.99：按钮宽度按「显示宽度」计算——ttk 的 width 单位是英文字符宽，
        # 而中日韩全角字符渲染宽度约为英文的 2 倍。v1.3.98 用 len(text) 直接设宽度，
        # 导致"上一步"(3字→width 4) 只装得下 2 个汉字，显示成"上一…"；
        # "保存修正后论文"(7字→width 7) 仍被截断。改为逐字累加全角算 2、半角算 1，
        # 再左右各留 1 字符余量，确保任何文案都完整显示。
        self._set_btn_text = lambda btn, text, **kw: btn.config(
            text=text, width=_btn_display_width(text), **kw)

        # 卡片底部章节小字（文艺学术点缀）
        tk.Label(card, text="贰 · 处理", bg=PANEL, fg="#b8b0a0",
                 font=F_FOOT).pack(side="bottom", pady=(0, 8))

        self._refresh_wizard()

    # ----------------------------------------------------- 步骤时间线（向导）
    def _build_timeline(self, parent):
        tl = tk.Frame(parent, bg=PANEL)
        tl.pack(fill="x", padx=14, pady=(4, 2))
        self._step_circle = []
        self._step_title = []
        self._step_line = []
        n = len(self.step_defs)
        for i, (mode, label) in enumerate(self.step_defs):
            row = tk.Frame(tl, bg=PANEL)
            row.pack(fill="x", pady=2)
            col = tk.Frame(row, bg=PANEL)
            col.pack(side="left", padx=(0, 10))
            circ = tk.Label(col, text=str(i + 1), bg="#ffffff", fg="#8b8378",
                            font=F_STAT, width=2, height=1, relief="flat",
                            highlightthickness=1, highlightbackground="#c9c1ae")
            circ.pack()
            line = None
            if i < n - 1:
                line = tk.Frame(col, width=2, height=22, bg="#e3dccb")
                line.pack()
            self._step_circle.append(circ)
            self._step_line.append(line)
            txt = tk.Frame(row, bg=PANEL)
            txt.pack(side="left", fill="x", expand=True)
            title = tk.Label(txt, text=label, bg=PANEL, fg=MUTED, font=F_SUBTITLE)
            title.pack(anchor="w")
            tk.Label(txt, text=self.step_desc[i], bg=PANEL, fg=MUTED,
                     font=F_SMALL).pack(anchor="w")
            self._step_title.append(title)

    def _refresh_wizard(self):
        """根据 step_index / _errored / _fix_phase 重绘时间线、计数与按钮三态。"""
        n = len(self.step_defs)
        for i, (mode, label) in enumerate(self.step_defs):
            circ = self._step_circle[i]
            title = self._step_title[i]
            line = self._step_line[i]
            # 第三步（最后一步）修正完成（fixed/saved）也视为完成态
            is_done = i < self.step_index or (
                i == n - 1 and self._fix_phase in ("fixed", "saved"))
            if is_done:
                circ.config(bg="#ffffff", fg=OKC, highlightbackground=OKC, text="✔")
                title.config(fg=INK)
                if line: line.config(bg=OKC)
            elif i == self.step_index and self._errored:
                circ.config(bg="#ffffff", fg=ERRC, highlightbackground=ERRC, text="✕")
                title.config(fg=INK)
                if line: line.config(bg="#e3dccb")
            elif i == self.step_index:
                circ.config(bg="#ffffff", fg=ACCENT, highlightbackground=ACCENT,
                           text=str(i + 1))
                title.config(fg=INK)
                if line: line.config(bg="#e3dccb")
            else:
                circ.config(bg="#ffffff", fg="#8b8378", highlightbackground="#c9c1ae",
                            text=str(i + 1))
                title.config(fg=MUTED)
                if line: line.config(bg="#e3dccb")
        self._step_counter.config(text="%d / %d" % (min(self.step_index + 1, n), n))
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
            # 第①/②步按钮文案随步骤精细变化（不再一律“下一步”）
            if self.step_index == 0:
                if self.input_mode.get() == "manual":
                    _next = ("① 使用手动格式" if self._manual_filled else "① 填写格式要求")
                else:
                    _next = ("① 提取模板要求" if self.template_path.get().strip()
                             else "① 按通用规范继续")
            elif self.step_index == 1:
                _next = "② 开始检查"
            else:
                _next = "下一步"
            self._set_btn_text(self._next_btn, _next, command=self._run_step,
                                  state="disabled" if self.running else "normal")
            self._set_btn_text(self._prev_btn, "上一步",
                                  state="disabled" if (self.running or self.step_index == 0) else "normal",
                                  command=self._go_prev)

    def _set_bar(self, state, hint=None):
        """底部状态栏：idle / running / done / error。"""
        cmap = {"idle": OKC, "running": RUN, "done": OKC, "error": ERRC}
        tmap = {
            "idle": ("就绪 · 论文格式医生 v%s · 本机处理" % APP_VERSION, "请按步骤操作"),
            "running": ("处理中…", hint or "正在处理"),
            "done": ("已完成", "可再处理一篇"),
            "error": ("出错", "请重试或联系客服"),
        }
        left, right = tmap.get(state, tmap["idle"])
        c = cmap.get(state, MUTED)
        self.bar_left.config(text=left, fg=c)
        self.bar_dot.config(fg=c)
        self.bar_right.config(text=right)

    def _on_step_done(self, idx, mode):
        self._step_circle[idx].config(bg="#ffffff", fg=OKC, highlightbackground=OKC, text="✔")
        self._step_title[idx].config(fg=INK)
        if self._step_line[idx]:
            self._step_line[idx].config(bg=OKC)
        if mode == "fix":
            # 第三步：修正完成→进入“保存修正后论文”子状态（先修正、后导出）。
            # 不前进到“再处理一篇”，按钮由 _fix_phase 驱动为「保存修正后论文」。
            self._fix_phase = "fixed"
            self._set_status("修正完成，请点击「保存修正后论文」", OKC)
            self._set_bar("done")
        else:
            self.step_index = idx + 1
            self._set_status("已完成", OKC)
            self._set_bar("done")
        self._refresh_wizard()

    def _on_step_error(self, idx, mode, err):
        self._errored = True
        if mode == "fix":
            self._fix_phase = "idle"
        self._set_status("未能完成，请查看提示", ERRC)
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
            title="选择待处理论文",
            filetypes=[("Word 文档", "*.docx *.doc *.wps"), ("所有文件", "*.*")])
        if p:
            self.thesis_path.set(p)
            self._thesis_name.config(text=os.path.basename(p), fg=INK)
            self._thesis_dot.config(text="✓", fg=OKC)
            self._thesis_lbl.config(text="论文已选择", fg=INK)

    def _pick_template(self):
        p = filedialog.askopenfilename(
            title="选择学校模板",
            filetypes=[("Word 文档", "*.docx *.doc *.wps"), ("所有文件", "*.*")])
        if p:
            self.template_path.set(p)
            self._template_name.config(text=os.path.basename(p), fg=INK)
            self._tpl_dot.config(text="✓", fg=OKC)
            self._tpl_lbl.config(text="模板已选择", fg=INK)
            # 换了新模板：重新允许提取画像（清除"放弃"标记）
            self._profile_abandoned = False
            self._profile_confirmed = False

    def _clear_profile(self):
        self.profile_path.set("")
        self._update_profile_box()

    # --------------------------------------------------- 手动填写格式（与模板互斥）
    def _set_input_mode(self, *_a):
        """第①步输入方式互斥：学校模板 ↔ 手动填写格式，二选一。"""
        if self.input_mode.get() == "manual":
            # 用 winfo_manager()=="pack" 守卫：既能覆盖“启动期已 pack 但未 map”的情况
            # （避免两区域同时显示），又不会对“从未 pack”的控件（如导师版默认隐藏的手动区）报错
            if self._template_box.winfo_manager() == "pack":
                self._template_box.pack_forget()
            if self._tpl_info.winfo_manager() == "pack":
                self._tpl_info.pack_forget()
            if self._manual_frame.winfo_manager() != "pack":
                self._manual_frame.pack(fill="x", padx=14, pady=(2, 6), after=self._mode_frame)
        else:
            if self._manual_frame.winfo_manager() == "pack":
                self._manual_frame.pack_forget()
            # 关键：用 after= 锚定原始位置，否则重新 pack 会落到卡片底部、打乱上方“格式来源”框顺序
            if self._template_box.winfo_manager() != "pack":
                self._template_box.pack(fill="x", padx=14, pady=6, after=self._thesis_box)
            if self._tpl_info.winfo_manager() != "pack":
                self._tpl_info.pack(fill="x", padx=14, pady=(4, 2), after=self._mode_frame)
            # 切回模板：清空手动填写状态（已记住配置保留，下次自动载入）
            self._manual_filled = False
            self.manual_profile_path = ""
            # 清空真正被引擎读取的画像路径，避免沿用刚才的手动画像
            self.profile_path.set("")
        self._refresh_manual_status()
        self._refresh_wizard()

    def _refresh_manual_status(self):
        if self._manual_filled:
            self._manual_btn.config(text="重新填写格式 ›", bg="#f0e7d2", fg="#6f5526",
                                    activebackground="#e3dccb", activeforeground="#6f5526")
            self._manual_status.config(text="✓ 已填写手动格式（点此修改）")
        else:
            self._manual_btn.config(text="手动填写格式要求 ›", bg="#e3a93b", fg="#5a3d12",
                                    activebackground="#f0bf5e", activeforeground="#5a3d12")
            self._manual_status.config(text="")

    def _open_manual_dialog(self):
        """打开手动填写格式大弹窗；确认后构建画像并写临时 json，标记已填写。"""
        ok, values = self._manual_profile_dialog(self._manual_form_values)
        if not ok:
            return
        remember = values.pop("_remember", False)
        self._manual_form_values = values
        self._manual_remember = remember
        if remember:
            self._save_manual_config(values)
        else:
            self._delete_manual_config()
        profile = self._build_manual_profile(values)
        try:
            out = os.path.join(tempfile.gettempdir(),
                               "tfd_manual_profile_%s.json"
                               % hashlib.sha1(json.dumps(values, sort_keys=True, ensure_ascii=False).encode("utf-8")).hexdigest()[:10])
            with open(out, "w", encoding="utf-8") as f:
                json.dump(profile, f, ensure_ascii=False, indent=2)
            self.manual_profile_path = out
            self._manual_filled = True
            self.root.after(0, lambda: self.profile_path.set(out))
            self._profile_confirmed = True
            self.root.after(0, self._update_profile_box)
            self.root.after(0, self._refresh_wizard)
            self.root.after(0, self._refresh_manual_status)
        except Exception as e:
            self._debug("[手动画像构建失败] " + str(e))
            messagebox.showerror("构建失败", "手动格式画像生成失败：%s" % e)

    def _manual_profile_dialog(self, init):
        """手动填写格式弹窗。返回 (ok, values)；values 含 _remember 键。

        init：已记住/上次填写的表单值，用于预填。
        """
        result = {"ok": False, "values": {}}
        top = tk.Toplevel(self.root)
        top.title("手动填写格式要求")
        top.configure(bg=PAPER)
        top.transient(self.root)
        top.grab_set()
        top.geometry(_geo(820, 760) + "+%d+%d" % (self.root.winfo_rootx() + 40,
                                                  self.root.winfo_rooty() + 20))

        tk.Label(top, text="没有学校模板？手动录入格式要求", bg=PAPER, fg=INK,
                 font=F_DIALOG_TITLE).pack(pady=(12, 2))
        tk.Label(top, text="留空的项按通用规范处理；字体可下拉选择，也可直接输入特殊字体。",
                 bg=PAPER, fg=MUTED, font=F_SMALL).pack(pady=(0, 4))

        # 滚动区
        body = tk.Frame(top, bg=PAPER)
        body.pack(fill="both", expand=True, padx=16, pady=2)
        canvas = tk.Canvas(body, bg=PAPER, highlightthickness=0)
        vbar = ttk.Scrollbar(body, orient="vertical", command=canvas.yview)
        form = tk.Frame(canvas, bg=PAPER)
        form.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=form, anchor="nw")
        canvas.configure(yscrollcommand=vbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        vbar.pack(side="right", fill="y")

        def _wheel(e):
            d = getattr(e, "delta", 0) or 0
            canvas.yview_scroll(-3 if d > 0 else 3, "units")
        for w in (canvas, form):
            w.bind("<MouseWheel>", _wheel)
            w.bind("<Button-4>", _wheel)
            w.bind("<Button-5>", _wheel)

        widgets = {}
        row = 0
        for label, key, kind, hint in _MANUAL_FORM:
            if key is None:  # 分区标题
                tk.Label(form, text=label, bg="#eef1ea", fg="#3f5a35",
                         font=F_SMALL_B, anchor="w").grid(
                    row=row, column=0, columnspan=3, sticky="ew", padx=4, pady=(8, 3))
                row += 1
                continue
            tk.Label(form, text=label, bg=PAPER, fg=INK, font=F_SMALL).grid(
                row=row, column=0, sticky="e", padx=(0, 8), pady=3)
            cur = (init or {}).get(key, "")
            if kind in ("zh_font", "en_font"):
                var = tk.StringVar(value=cur)
                cb = ttk.Combobox(form, textvariable=var, width=22, font=F_SMALL,
                                  values=_MANUAL_ZH_FONTS if kind == "zh_font" else _MANUAL_EN_FONTS)
                cb.grid(row=row, column=1, sticky="w", pady=3)
                widgets[key] = (var, "combo_edit")
            elif kind == "size":
                var = tk.StringVar(value=cur)
                cb = ttk.Combobox(form, textvariable=var, width=14, font=F_SMALL,
                                  values=_MANUAL_SIZES, state="readonly")
                cb.grid(row=row, column=1, sticky="w", pady=3)
                widgets[key] = (var, "combo")
            elif kind == "align":
                var = tk.StringVar(value=cur)
                cb = ttk.Combobox(form, textvariable=var, width=14, font=F_SMALL,
                                  values=_MANUAL_ALIGNS, state="readonly")
                cb.grid(row=row, column=1, sticky="w", pady=3)
                widgets[key] = (var, "combo")
            elif kind == "line_type":
                var = tk.StringVar(value=cur)
                cb = ttk.Combobox(form, textvariable=var, width=14, font=F_SMALL,
                                  values=_MANUAL_LINE_TYPES, state="readonly")
                cb.grid(row=row, column=1, sticky="w", pady=3)
                widgets[key] = (var, "combo")
            elif kind == "bold":
                var = tk.BooleanVar(value=bool(cur))
                cb = tk.Checkbutton(form, variable=var, bg=PAPER,
                                    activebackground=PAPER)
                cb.grid(row=row, column=1, sticky="w", pady=3)
                widgets[key] = (var, "bool")
            else:  # num
                var = tk.StringVar(value=cur)
                ent = tk.Entry(form, textvariable=var, width=14, font=F_SMALL,
                               relief="solid", bd=1)
                ent.grid(row=row, column=1, sticky="w", pady=3)
                widgets[key] = (var, "text")
            if hint:
                tk.Label(form, text=hint, bg=PAPER, fg="#9a9486",
                         font=F_FOOT).grid(row=row, column=2, sticky="w", padx=(6, 0), pady=3)
            row += 1

        top.after(10, lambda: canvas.yview_moveto(0))

        def on_confirm():
            vals = {}
            for key, (var, typ) in widgets.items():
                if typ == "bool":
                    if var.get():
                        vals[key] = True
                    continue
                text = (var.get() or "").strip()
                if text:
                    vals[key] = text
            result["values"] = vals
            result["ok"] = True
            top.destroy()

        def on_cancel():
            result["ok"] = False
            top.destroy()

        # 底部：记住勾选 + 按钮（固定贴底）
        btns = tk.Frame(top, bg=PAPER)
        btns.pack(side="bottom", fill="x", pady=10, padx=16)
        rem_var = tk.BooleanVar(value=self._manual_remember)
        tk.Checkbutton(btns, text="记住此手动配置", variable=rem_var, bg=PAPER,
                       activebackground=PAPER, font=F_SMALL).pack(side="left", padx=4)
        # 把 rem_var 纳入 result：确认后补写 _remember
        def _wrap_confirm():
            on_confirm()
            if result["ok"]:
                result["values"]["_remember"] = rem_var.get()
        ttk.Button(btns, text="确定，使用此格式", style="Primary.TButton",
                   command=_wrap_confirm).pack(side="right", padx=6)
        ttk.Button(btns, text="取消", command=on_cancel).pack(side="right", padx=6)

        top.wait_window()
        return result["ok"], result.get("values", {})

    def _build_manual_profile(self, v):
        """把手动表单值转成引擎消费的 profile（仅填了的字段写入，空项留空→引擎用通用规范兜底）。"""
        def _m_num(x):
            try:
                return float(x)
            except (TypeError, ValueError):
                return None
        def _m_sz(name):
            if not name:
                return (None, None)
            pt = _CN_SIZE_PT.get(name)
            if pt is None:
                try:
                    pt = float(name)
                except (TypeError, ValueError):
                    return (None, None)
            return (int(pt * 2), name)
        def _m_line(rule, val):
            if not rule:
                return (None, None)
            if rule.startswith("单倍"):
                return ("single", 240)
            if rule.startswith("1.5"):
                return ("auto", 360)
            if rule.startswith("2倍"):
                return ("auto", 480)
            p = _m_num(val)
            if rule.startswith("固定"):
                return ("exact", p) if p is not None else (None, None)
            if rule.startswith("最小"):
                return ("atLeast", p) if p is not None else (None, None)
            return (None, None)
        def _m_align(disp):
            return ALIGN_CODE.get(disp)

        levels = {}
        for lk in ("1", "2", "3"):
            spec = {}
            if v.get("%s_zh_font" % lk):
                spec["zh_font"] = v["%s_zh_font" % lk]
            if v.get("%s_en_font" % lk):
                spec["en_font"] = v["%s_en_font" % lk]
            sz, size = _m_sz(v.get("%s_size" % lk))
            if sz:
                spec["sz"] = sz
                spec["size"] = size
            al = _m_align(v.get("%s_align" % lk))
            if al:
                spec["align"] = al
            if v.get("%s_bold" % lk):
                spec["bold"] = True
            bp = _m_num(v.get("%s_before" % lk))
            if bp is not None:
                spec["before_pt"] = bp
            ap = _m_num(v.get("%s_after" % lk))
            if ap is not None:
                spec["after_pt"] = ap
            lr, lv = _m_line(v.get("%s_line_type" % lk), v.get("%s_line_val" % lk))
            if lr:
                spec["line_rule"] = lr
                spec["line_val"] = lv
            if spec:
                levels[lk] = spec
        # 正文
        body = {}
        if v.get("body_zh_font"):
            body["zh_font"] = v["body_zh_font"]
        if v.get("body_en_font"):
            body["en_font"] = v["body_en_font"]
        sz, size = _m_sz(v.get("body_size"))
        if sz:
            body["sz"] = sz
            body["size"] = size
        al = _m_align(v.get("body_align"))
        if al:
            body["align"] = al
        lr, lv = _m_line(v.get("body_line_type"), v.get("body_line_val"))
        if lr:
            body["line_rule"] = lr
            body["line_val"] = lv
        ic = _m_num(v.get("body_indent"))
        if ic is not None:
            body["indent_chars"] = ic
            body["indent_type"] = "first"
        if body:
            levels["body"] = body
        # 摘要标题 / 摘要正文
        at = {}
        if v.get("abs_t_zh_font"):
            at["zh_font"] = v["abs_t_zh_font"]
        if v.get("abs_t_en_font"):
            at["en_font"] = v["abs_t_en_font"]
        sz, size = _m_sz(v.get("abs_t_size"))
        if sz:
            at["sz"] = sz
            at["size"] = size
        if at:
            levels["abstract_title"] = at
        ab = {}
        if v.get("abs_zh_font"):
            ab["zh_font"] = v["abs_zh_font"]
        if v.get("abs_en_font"):
            ab["en_font"] = v["abs_en_font"]
        sz, size = _m_sz(v.get("abs_size"))
        if sz:
            ab["sz"] = sz
            ab["size"] = size
        lr, lv = _m_line(v.get("abs_line_type"), v.get("abs_line_val"))
        if lr:
            ab["line_rule"] = lr
            ab["line_val"] = lv
        if ab:
            levels["abstract"] = ab
        # 参考文献标题 / 条目
        rt = {}
        if v.get("ref_t_zh_font"):
            rt["zh_font"] = v["ref_t_zh_font"]
        if v.get("ref_t_en_font"):
            rt["en_font"] = v["ref_t_en_font"]
        sz, size = _m_sz(v.get("ref_t_size"))
        if sz:
            rt["sz"] = sz
            rt["size"] = size
        if rt:
            levels["reference_heading"] = rt
        ri = {}
        if v.get("ref_i_zh_font"):
            ri["zh_font"] = v["ref_i_zh_font"]
        if v.get("ref_i_en_font"):
            ri["en_font"] = v["ref_i_en_font"]
        sz, size = _m_sz(v.get("ref_i_size"))
        if sz:
            ri["sz"] = sz
            ri["size"] = size
        lr, lv = _m_line(v.get("ref_line_type"), v.get("ref_line_val"))
        if lr:
            ri["line_rule"] = lr
            ri["line_val"] = lv
        # 参考文献条目悬挂缩进：按编号位数三档（字符），与引擎 _ref_entry_spec_for_text 对齐
        # 同时修掉旧版 ref_hanging 缺 ref_i_ 前缀导致读不到的隐藏 bug
        _tiers = []
        for _dig, _key in ((1, "ref_i_hanging_1"), (2, "ref_i_hanging_2"), (3, "ref_i_hanging_3")):
            _hc = _m_num(v.get(_key))
            if _hc is not None:
                _tiers.append({"digits": _dig, "chars": _hc})
        if _tiers:
            ri["hanging_tiers"] = _tiers
            ri["indent_type"] = "hanging"
        # 参考文献条目格式必须写入 levels["reference"]（引擎参考文献条目循环读的就是这个 key，
        # 不是 ref_item）；GB/T 7714 引用格式标记一并带上，便于引擎自动重排
        if ri:
            ri["style"] = "gb7714"
            ri["format"] = "sequential"
            levels["reference"] = ri
        else:
            levels["reference"] = {"style": "gb7714", "format": "sequential"}
        # 标题样式映射（通用 Heading1/2/3，让引擎识别标题段落）
        heading_styles = {str(i): {"styleId": "Heading%d" % i, "name": "标题 %d" % i}
                          for i in (1, 2, 3)}
        # 页面边距（cm）：写入 *_cm 键、保留厘米值，引擎 _fix_sect_margins 只读 top_cm 等
        page = {}
        for ck, cmkey in (("page_top", "top_cm"), ("page_bottom", "bottom_cm"),
                          ("page_left", "left_cm"), ("page_right", "right_cm")):
            cm = _m_num(v.get(ck))
            if cm is not None:
                page[cmkey] = cm
        profile = {
            "source": "手动填写格式",
            "manual": True,
            "comment_count": 0,
            "headingStyles": heading_styles,
            "levels": levels,
        }
        if page:
            profile["spec"] = {"page": page}
        return profile

    def _manual_config_path(self):
        return os.path.join(os.path.expanduser("~"), ".tfd_license", "manual_profile.json")

    def _save_manual_config(self, values):
        try:
            p = self._manual_config_path()
            os.makedirs(os.path.dirname(p), exist_ok=True)
            with open(p, "w", encoding="utf-8") as f:
                json.dump({"values": values, "remember": True}, f, ensure_ascii=False)
        except Exception as e:
            self._debug("[手动配置保存失败] " + str(e))

    def _load_manual_config(self):
        try:
            p = self._manual_config_path()
            if not os.path.isfile(p):
                return None
            return json.load(open(p, encoding="utf-8"))
        except Exception:
            return None

    def _delete_manual_config(self):
        try:
            p = self._manual_config_path()
            if os.path.isfile(p):
                os.remove(p)
        except Exception:
            pass

    def _restore_manual_config(self):
        """启动时：若已记住手动配置，自动切到手动模式并构建画像。"""
        cfg = self._load_manual_config()
        if not cfg:
            return
        self._manual_form_values = cfg.get("values", {})
        self._manual_remember = True
        self.input_mode.set("manual")
        self._set_input_mode()
        profile = self._build_manual_profile(self._manual_form_values)
        try:
            out = os.path.join(tempfile.gettempdir(), "tfd_manual_profile_autoload.json")
            with open(out, "w", encoding="utf-8") as f:
                json.dump(profile, f, ensure_ascii=False, indent=2)
            self.manual_profile_path = out
            self._manual_filled = True
            self.profile_path.set(out)
            self._profile_confirmed = True
        except Exception as e:
            self._debug("[手动配置自动载入失败] " + str(e))
        self._update_profile_box()

    # ---------------------------------------------------------------- run（向导）
    def _go_prev(self):
        """上一步：回退一个步骤，该步及其后的进度重置为待办，可重新执行。

        第三步（修正）内的子状态：先退回“未修正”，停留在第三步便于重新修正，
        而非跳回检查步骤。
        """
        if self.running or self.step_index <= 0:
            return
        if self.step_index == len(self.step_defs) - 1 and self._fix_phase != "idle":
            self._fix_phase = "idle"
            self._errored = False
            self._set_status("请按步骤操作", MUTED)
            self._set_bar("idle")
            self._refresh_wizard()
            return
        self.step_index -= 1
        self._errored = False
        # v1.3.49：回退后允许重新确认画像（清除"放弃"标记）
        self._profile_abandoned = False
        self._set_status("请按步骤操作", MUTED)
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
            messagebox.showerror("缺少输入", "请先选择“待处理论文”。")
            return

        # v1.3.116：第①步手动模式但未填写格式→拦截，提示先填表单
        if idx == 0 and self.input_mode.get() == "manual" and not self._manual_filled:
            messagebox.showinfo("请先填写格式要求",
                                "您选择了「手动填写格式」，但尚未填写格式要求。\n\n"
                                "请点击「手动填写格式要求 ›」按钮，录入标题 / 正文 / 页边距等格式，再继续。")
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
        self._set_status(status, RUN)
        self._set_bar("running", self.step_defs[idx][1])
        # 点亮当前步
        self._step_circle[idx].config(bg="#ffffff", fg=ACCENT,
                                      highlightbackground=ACCENT, text=str(idx + 1))
        self._step_title[idx].config(fg=INK)
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
                self.progress.pack(fill="x", padx=14, pady=(4, 2))
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
        # 手动模式：画像已由表单构建，跳过模板确认弹窗
        if self.input_mode.get() == "manual" and self._manual_filled and self.manual_profile_path:
            return self.manual_profile_path
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
        top = tk.Toplevel(self.root)
        top.title("学校模板要求 · 请确认")
        top.configure(bg=PAPER)
        top.transient(self.root)
        top.grab_set()
        top.geometry(_geo(620, 620) + "+%d+%d" % (self.root.winfo_rootx() + 90,
                                                  self.root.winfo_rooty() + 30))

        tk.Label(top, text="已提取出学校模板的格式要求", bg=PAPER, fg=INK,
                 font=F_DIALOG_TITLE).pack(pady=(14, 2))
        tk.Label(top, text="请核对是否与学校规定一致；如有不准，可直接修改后确认。",
                 bg=PAPER, fg=MUTED, font=F_SMALL).pack(pady=(0, 6))

        # v1.3.46：批注来源提示——有批注=绿色"以批注为准"；无批注=红色警告（防误解、减纠纷）
        _n_cmt = int((profile or {}).get("comment_count", 0) or 0)
        if _n_cmt > 0:
            src_note = tk.Frame(top, bg="#eef4ea", highlightthickness=1,
                                highlightbackground="#b9cdaa")
            src_txt = "已读取学校模板批注 %d 条 —— 以下要求以【批注】为准（最权威）。" % _n_cmt
            src_fg = "#3f6b35"
        else:
            src_note = tk.Frame(top, bg="#fbeae8", highlightthickness=1,
                                highlightbackground="#e0b4ae")
            src_txt = ("该模板【未检测到批注】。学校批注是最权威的格式要求；"
                       "无批注时以下要求来自模板样式定义 / 通用规范，"
                       "可能与学校规定有出入，请仔细核对后再确认。")
            src_fg = "#9e3b30"
        src_note.pack(fill="x", padx=16, pady=(0, 6))
        tk.Label(src_note, text=src_txt, bg=src_note.cget("bg"), fg=src_fg,
                 font=F_SMALL, justify="left", anchor="w",
                 wraplength=560).pack(fill="x", padx=10, pady=6)

        sum_f = tk.Frame(top, bg=PANEL, highlightthickness=1, highlightbackground=LINE)
        sum_f.pack(fill="x", padx=16, pady=3)
        sum_txt = tk.Text(sum_f, height=6, wrap="word", bg="#fbf8f1", fg=INK,
                          font=F_SMALL, relief="flat", padx=10, pady=6)
        sum_txt.insert("1.0", "\n".join(_profile_summary(profile)))
        sum_txt.config(state="disabled")
        sum_txt.pack(fill="x")

        tk.Label(top, text="如需修正，直接修改下列项目（留空表示保持提取结果）",
                 bg=PAPER, fg=ACCENT, font=F_SMALL).pack(anchor="w", padx=18, pady=(8, 2))

        # 表单区：canvas 与滚动条同在一个 frame 内，滚动条贴右侧整个高度（不沉到右下角）
        form_area = tk.Frame(top, bg=PAPER)
        form_area.pack(fill="both", expand=True, padx=(18, 0), pady=3)
        canvas = tk.Canvas(form_area, bg=PAPER, highlightthickness=0)
        vbar = ttk.Scrollbar(form_area, orient="vertical", command=canvas.yview)
        form = tk.Frame(canvas, bg=PAPER)
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
            tk.Label(form, text=label, bg=PAPER, fg=INK, font=F_SMALL).grid(
                row=i, column=0, sticky="e", padx=(0, 10), pady=3)
            cur = _field_value(profile, candidates)
            if kind == "align":
                var = tk.StringVar(value=ALIGN_DISPLAY.get(cur, cur))
                cb = ttk.Combobox(form, textvariable=var, width=20, font=F_SMALL,
                                  values=list(ALIGN_DISPLAY.values()), state="readonly")
                cb.grid(row=i, column=1, sticky="w", pady=3)
                entries[i] = (candidates, var, "align", var.get())
            elif kind == "ref":
                has_ref = bool(cur)
                var = tk.StringVar(value="已提取" if has_ref else "未提取（按通用规范检查）")
                tk.Entry(form, textvariable=var, width=28, font=F_SMALL,
                         relief="flat", bd=0, bg=PAPER, state="disabled",
                         disabledforeground=OKC if has_ref else MUTED).grid(
                    row=i, column=1, sticky="w", pady=3)
                entries[i] = (candidates, var, "ref", var.get())
            else:
                var = tk.StringVar(value=cur)
                tk.Entry(form, textvariable=var, width=22, font=F_SMALL,
                         relief="solid", bd=1).grid(row=i, column=1, sticky="w", pady=3)
                entries[i] = (candidates, var, "text", var.get())

        _NUM_FIELDS = ("indent_chars", "line_val", "top_cm", "bottom_cm",
                       "left_cm", "right_cm", "before_pt", "after_pt", "hanging_cm")

        def on_confirm():
            edits = []
            for (candidates, var, kind, init_text) in entries.values():
                if kind == "ref":
                    continue  # 参考文献格式为只读提示，不参与修改
                text = var.get().strip()
                # v1.3.43：只写回【真正被修改】的字段——此前把"所有非空字段"全部写回，
                # 导致未修改的 sz（显示为磅）被当成半磅写回、数值字段变字符串，
                # 修正引擎 %d 崩溃/字号错乱（用户实测"一键修正没改论文"的根因）。
                if text == init_text or not text:
                    continue
                if kind == "align":
                    code = ALIGN_CODE.get(text, text)
                    edits.append((candidates[0], code))
                    continue
                k = candidates[0][-1]
                if k == "sz":
                    # 确认页显示为磅（_field_value ÷2），写回需还原为半磅（×2）。
                    # v1.3.49：支持输入中文字号名（如"小四"→12磅）或数字磅值。
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
                    # v1.3.47：客户填了首行缩进 → 同时落定 indent_type="first"，
                    # 否则画像无 indent_type 时引擎不会套用缩进（用户实测踩坑）
                    if k == "indent_chars":
                        edits.append((candidates[0][:-1] + ("indent_type",), "first"))
                    # v1.3.48：客户填了行距(磅) → 同时落定 line_rule="exact"（固定值行距），
                    # 否则画像无 line_rule 时引擎不设行距（实测发现同款坑）
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

        # 按钮区固定底部（表单区自动占据剩余空间并可滚动，滚动条置顶）
        btns = tk.Frame(top, bg=PAPER)
        btns.pack(side="bottom", fill="x", pady=10)
        # 重新 pack 表单区：让出底部给按钮，保证滚动到底/任何位置按钮都贴底可见
        form_area.pack_forget()
        form_area.pack(fill="both", expand=True, padx=(18, 0), pady=3)
        ttk.Button(btns, text="确认，使用此要求", style="Primary.TButton",
                   command=on_confirm).pack(side="left", padx=6)
        ttk.Button(btns, text="放弃（不使用画像）", command=on_cancel).pack(side="left", padx=6)
        top.after(10, lambda: canvas.yview_moveto(0))

        top.wait_window()
        return result["ok"], result["edits"]

    # ------------------------------------------------------------ 各步骤
    def _do_profile(self, src):
        """第①步：提取【学校模板】的格式要求（不是从论文提取）。

        画像必须来自学校模板才有意义；若未选模板，则没有“学校要求”可提取，
        走通用规范并明确告知客户，且不弹一个空的“确认模板要求”框。
        """
        # v1.3.116：手动模式——画像已由表单构建，跳过模板提取（互斥）
        if self.input_mode.get() == "manual":
            if self._manual_filled and self.manual_profile_path:
                self.profile_path.set(self.manual_profile_path)
                self._profile_confirmed = True
                self.root.after(0, self._update_profile_box)
                self.root.after(0, self._refresh_wizard)
            else:
                self.root.after(0, lambda: messagebox.showinfo(
                    "请先填写格式要求",
                    "您选择了「手动填写格式」，但尚未填写格式要求。\n\n"
                    "请点击「手动填写格式要求 ›」按钮，录入标题 / 正文 / 页边距等格式，再继续。"))
            return
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
                self._show_trial_exhausted()   # 主线程弹升级引导
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
        # 参考文献 GB/T 7714 重排（用户选定的优化项）：标题/正文修正之后独立 pass 执行；
        # 失败不影响主交付物（已修正论文）。
        try:
            _ref_out = os.path.splitext(dst)[0] + "_ref.docx"
            engine.run_reformat_refs(dst, _ref_out)
            if os.path.isfile(_ref_out):
                try:
                    os.replace(_ref_out, dst)
                except Exception:
                    if os.path.isfile(_ref_out):
                        os.remove(_ref_out)
                report = (report or "") + "\n\n> 参考文献已按 GB/T 7714 自动重排。"
        except Exception as e:
            self._debug("[参考文献重排跳过] " + str(e))
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
        """试用次数用完：主线程弹升级引导（公众号引导）。"""
        ev = threading.Event()
        box = {}

        def show():
            box["v"] = self._modal(
                "试用次数已用完",
                "本机试用已满 %d 次。\n\n"
                "正式版激活后：不限次数修正、输出无水印文档、一键交稿。\n\n"
                "获取激活码：请关注公众号【芦苇不熬夜】（ID：reedskill）或联系客服。\n"
                "激活教程与购买方式详见官网 reedskill.com。\n"
                "激活码购买与激活问题，公众号留言即可。" % trial.TRIAL_LIMIT,
                [("ok", "知道了")])
            ev.set()

        self.root.after(0, show)
        ev.wait(30)

    def _on_trial_blocked(self):
        """试用被拦截（用完/取消）：停留在当前步，给友好提示。"""
        self._errored = False
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
        result = {"v": None}
        top = tk.Toplevel(self.root)
        top.title(title)
        top.configure(bg=PAPER)
        top.transient(self.root)
        top.grab_set()
        top.geometry("+%d+%d" % (self.root.winfo_rootx() + 140,
                                 self.root.winfo_rooty() + 120))
        tk.Label(top, text=text, bg=PAPER, fg=BODY, font=F_BODY, justify="left",
                 wraplength=480).pack(padx=24, pady=(20, 14))
        fr = tk.Frame(top, bg=PAPER)
        fr.pack(pady=(0, 18))
        for key, label in buttons:
            style = "Primary.TButton" if key == "ok" else "TButton"
            b = ttk.Button(fr, text=label, style=style,
                           command=lambda k=key: (result.update(v=k), top.destroy()))
            b.pack(side="left", padx=6)
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
        p = self.profile_path.get().strip()
        if p and os.path.isfile(p):
            # v1.3.116：手动模式下显示为"已填写手动格式"，与模板来源区分
            if self.input_mode.get() == "manual" and self._manual_filled:
                self.profile_info_var.set("已填写手动格式")
            else:
                self.profile_info_var.set("已载入模板格式")
            if not self.profile_box.winfo_ismapped():
                # 跟随当前可见的来源区：手动模式跟手动区，模板模式跟模板区（避免 after 未 pack 控件报错）
                _anchor = self._manual_frame if (self.input_mode.get() == "manual"
                                                 and self._manual_frame.winfo_ismapped()) else self._template_box
                self.profile_box.pack(fill="x", padx=14, pady=(4, 8), after=_anchor)
        else:
            if self.profile_box.winfo_ismapped():
                self.profile_box.pack_forget()

    def _reset_wizard(self):
        """“再处理一篇”：回到第 1 步并清空选择。"""
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
        self._manual_filled = False
        self.manual_profile_path = ""
        self._set_input_mode()
        self._thesis_name.config(text="Word 文档 .docx / .doc / .wps", fg=MUTED)
        self._template_name.config(text="用于按学校要求检查 / 修正，更贴合要求", fg=MUTED)
        self._thesis_dot.config(text="○", fg=MUTED)
        self._thesis_lbl.config(text="未选择论文", fg=MUTED)
        self._tpl_dot.config(text="○", fg=MUTED)
        self._tpl_lbl.config(text="模板未选（可选）", fg=MUTED)
        self._update_profile_box()
        self._set_status("请按步骤操作", MUTED)
        self._set_bar("idle")
        self._refresh_wizard()

    # ---------------------------------------------- 内部日志（不展示客户）
    def _debug(self, text):
        self._msgs.append(text)
        if len(self._msgs) > 200:
            self._msgs = self._msgs[-200:]

    def _set_status(self, s, color):
        self.root.after(0, lambda: (self.status_var.set(s),
                                    self.status_lbl.config(fg=color),
                                    self.status_dot.config(fg=color)))

    def _update_trial_badge(self):
        """刷新右下角激活/试用标识。"""
        try:
            self._licensed = trial.is_licensed()
        except Exception:
            self._licensed = False
        if self._licensed:
            # v1.3.99：区分卡种展示——周卡/月卡/次卡不再是"永久"，要让用户看得见期限与余量
            try:
                s = license.license_summary()
            except Exception:
                s = None
            # v1.3.100：badge 已收紧为 ≤ 8 字（license.py），无需担心撑爆
            self._trial_btn.config(
                text=("%s ✓" % s["badge"]) if s else "正式版 ✓", state="disabled")
        else:
            left = trial.trials_left()
            if left > 0:
                self._trial_btn.config(text="试用版（剩 %d 次）· 激活" % left, state="normal")
            else:
                self._trial_btn.config(text="试用已用完 · 激活", state="normal")

    def _open_activation(self):
        """主界面右上角激活入口：打开激活窗，成功后刷新状态。"""
        if self.running:
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
    top = tk.Toplevel(root)
    top.title("激活 · 论文格式医生")
    top.configure(bg=PAPER)
    top.resizable(False, False)
    # v1.3.65：窗口尺寸自适应屏幕（笔记本 768 高屏幕时 720 高的窗口底部会被截在屏外）
    try:
        _sw = top.winfo_screenwidth()
        _sh = top.winfo_screenheight()
    except Exception:
        _sw, _sh = 1366, 768
    _s = _CUR_SCALE
    _w = max(420, min(int(620 * _s), _sw - 60))
    # v1.3.66：主界面激活入口无试用区，窗口更矮
    _base_h = 690 if show_trial else 600
    _h = max(420, min(int(_base_h * _s), _sh - 120))
    top.geometry("%dx%d" % (_w, _h))

    tk.Label(top, text="激 活 论 文 格 式 医 生", bg=PAPER, fg=INK,
             font=F_TITLE).pack(pady=(14, 4))
    # v1.3.99：卡片种类已扩展到 4 种（永久 / 周卡 / 月卡 / 次卡），文案不再统一说"永久、完全离线"
    tk.Label(top,
             text="请输入您购买的激活码以激活。\n"
                  "永久卡：激活后完全离线、永久可用；周卡 / 月卡：期限内可用；\n"
                  "次卡：每次修正需联网扣一次次数，离线时不可使用。",
             bg=PAPER, fg=MUTED, font=F_SMALL, wraplength=440, justify="center").pack(pady=(0, 10))

    card_var = tk.StringVar()
    tk.Entry(top, textvariable=card_var, width=40, font=F_BODY,
             relief="solid", bd=1, justify="center").pack(pady=(4, 6))

    msg_var = tk.StringVar()
    tk.Label(top, textvariable=msg_var, bg=PAPER, fg=ERRC, font=F_SMALL).pack(pady=(0, 6))

    def do_activate():
        card = card_var.get().strip()
        if not card:
            msg_var.set("请输入激活码")
            return
        btn_activate.config(state="disabled")
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
                    btn_activate.config(state="normal")
                    msg_var.set("验证超时：请检查网络后重试；或联系客服获取离线激活码")
                    license._log("gui: UI 看门狗触发（120 秒未等到结果）")
                    return
                top.after(300, poll)
                return
            btn_activate.config(state="normal")
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

    btn_activate = ttk.Button(top, text="激活", style="Primary.TButton",
               command=do_activate)
    btn_activate.pack(pady=(2, 4))

    # v1.3.58：试用入口放在显眼位置（激活按钮正下方，便于未购买客户先体验）
    # v1.3.66：仅启动时首次弹窗显示；从主界面激活入口打开时客户已在试用模式，无需再显示
    if show_trial:
        ttk.Button(top, text="还没有激活码？先试用（免费 2 次）", style="Ghost.TButton",
                   command=lambda: (result.update(v="trial"), top.destroy())).pack(pady=(4, 2))
        tk.Label(top, text="试用版可完整体验一键修正，输出带水印且为只读预览；正式版可编辑无水印。",
                 bg=PAPER, fg=MUTED, font=F_FOOT, wraplength=540).pack(pady=(0, 6))

    tk.Label(top, text="— 以下为特殊情形使用 —", bg=PAPER, fg=MUTED, font=F_SMALL).pack(pady=(8, 4))
    mc = license.get_machine_code()
    # 机器码 + 复制按钮并排一行（省高度）
    mc_row = tk.Frame(top, bg=PAPER)
    mc_row.pack(pady=(0, 6))
    tk.Label(mc_row, text="本机机器码：" + mc, bg=PAPER, fg=MUTED,
             font=F_MONO).pack(side="left")
    ttk.Button(mc_row, text="复制", style="Ghost.TButton",
               command=lambda: top.clipboard_append(mc)).pack(side="left", padx=(8, 0))

    off_var = tk.StringVar()
    tk.Label(top, text="离线激活码（网络不通时，联系客服获取）：", bg=PAPER, fg=MUTED,
             font=F_SMALL).pack(pady=(2, 2))
    tk.Entry(top, textvariable=off_var, width=46, font=F_MONO,
             relief="solid", bd=1).pack(pady=(2, 4))

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

    ttk.Button(top, text="使用离线激活码激活", style="Ghost.TButton",
               command=do_offline).pack(pady=(2, 4))

    ttk.Button(top, text="退出", command=lambda: top.destroy()).pack(pady=(0, 4))

    hl = tk.Frame(top, bg=PAPER)
    hl.pack(pady=(4, 2))
    tk.Button(hl, text="关于", bg=PAPER, fg=ACCENT, font=F_SMALL,
              relief="flat", cursor="hand2",
              command=lambda: show_about(top)).pack(side="left", padx=14)
    tk.Button(hl, text="使用帮助", bg=PAPER, fg=ACCENT, font=F_SMALL,
              relief="flat", cursor="hand2",
              command=lambda: show_help(top)).pack(side="left", padx=14)
    _site_label(hl).pack(side="left", padx=14)

    tk.Label(top, text="未签名程序提示：Windows 可能弹出 SmartScreen 拦截，点击「详细信息」→「仍要运行」即可打开（官网激活教程有图文演示）。",
             bg=PAPER, fg=MUTED, font=F_FOOT, wraplength=560).pack(pady=(2, 6))
    tk.Label(top, text="© 2026 论文格式医生 · 公众号【芦苇不熬夜】 ID：reedskill · 合作联系：hi@reedskill.com",
             bg=PAPER, fg=MUTED, font=F_FOOT, wraplength=560).pack(pady=(8, 10))

    # v1.3.67：按内容实际所需高度自动伸缩窗口（Tk 自动测量，保证底部版权完整显示不截断），
    # 仅设"不超过屏幕"上限。之前固定 600/690 高在字体缩放/DPI 差异下会截掉底部内容。
    try:
        top.update_idletasks()
        _req_h = top.winfo_reqheight()
        _req_w = top.winfo_reqwidth()
        _w2 = min(max(_req_w, _w), _sw - 60)
        _h2 = min(max(_req_h, _h), _sh - 120)
        top.geometry("%dx%d" % (_w2, _h2))
    except Exception:
        pass

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
    tk.Label(inner, text="论 文 格 式 医 生", bg=PAPER, fg=INK,
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
    wlbl("芦 苇 不 熬 夜  出 品", INK, F_HDR, (2, 8), nowrap=True)
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
    tk.Label(inner, text="使 用 帮 助", bg=PAPER, fg=INK,
             font=F_TITLE).pack(padx=padx, pady=(22, 4))
    tk.Label(inner, text="将学校模板告知软件，导入论文后依向导循序而行，格式自可妥帖。",
             bg=PAPER, fg=BODY, font=F_BODY, justify="center").pack(padx=padx, pady=(0, 12))
    tk.Frame(inner, bg=LINE, height=1).pack(fill="x", padx=padx, pady=(0, 8))

    def section(title):
        tk.Label(inner, text=title, bg=PAPER, fg=ACCENT, font=F_HDR,
                 anchor="w").pack(fill="x", padx=padx, pady=(14, 4))
        tk.Frame(inner, bg=LINE, height=1).pack(fill="x", padx=padx, pady=(0, 4))

    def item(title, body):
        tl = albl(title, INK, ("KaiTi", 12, "bold"))
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
         "试用版：免费试用 2 次，输出带水印的只读预览（不可直接编辑）。\n"
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
         "永久卡 / 周卡 / 月卡激活后可断网使用；次卡仅在扣减次数时联网，且只传输激活码与机器码，不含论文内容。"),
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
