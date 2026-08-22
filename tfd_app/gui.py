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

# 把 tfd_app 加入搜索路径（打包成 exe 后 sys.path 已含，但开发态下保险）
HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import tkinter.font as tkfont

from . import engine, license


ICON = os.path.join(HERE, "assets", "icon.png")

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
    "F_BTN":        ("Microsoft YaHei", 12, "bold"),  # 按钮
    "F_STAT":       ("KaiTi", 12),              # 状态文字（楷体）
    "F_SUBTITLE":   ("Microsoft YaHei", 11),    # 元信息
    "F_CARD_HDR":   ("KaiTi", 13, "bold"),      # 卡片标题（楷体）
    "F_FOOT":       ("Microsoft YaHei", 10),    # 状态栏 / 页脚 / 说明（次要，最小层级）
    "F_MONO":       ("Consolas", 9),            # 机器码 / 离线码（等宽）
    "F_DIALOG_TITLE": ("KaiTi", 14, "bold"),    # 弹窗标题（楷体）
    "F_ICON":       ("KaiTi", 12, "bold"),      # 印章图标（论 / 模，楷体朱砂）
}
APP_VERSION = "1.3.43"   # 与 VERSION 文件保持同步（状态栏显示用）
_FONTS = {}      # name -> (Font, base_size)
_CUR_SCALE = 1.0 # 当前窗口缩放比例（宽度 / 基准宽度，钳制 0.8~1.0：只缩小不放大）
BASE_W = 900     # 设计基准宽度（px），与主窗口默认 900x640 对应

def _init_fonts(root):
    """在 root 创建后调用：把 F_XXX 全局名替换为可缩放的 Font 命名对象。"""
    for name, spec in _FONT_BASE.items():
        kw = {"family": spec[0], "size": spec[1]}
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
    ("正文字号",     [("levels", "body", "size"), ("spec", "body", "size"),
                     ("levels", "body", "sz"), ("body", "size")], "text"),
    ("正文行距(磅)", [("levels", "body", "line_val"), ("spec", "body", "line_val"),
                     ("body", "line_val"), ("body", "line")], "text"),
    ("首行缩进(字符)", [("levels", "body", "indent_chars"), ("spec", "body", "indent_chars"),
                       ("body", "indent_chars"), ("body", "firstLineChars")], "text"),
    ("一级标题字体", [("levels", "1", "zh_font"), ("spec", "h1", "zh_font"),
                     ("spec", "h1", "font"), ("headings", "1", "font")], "text"),
    ("一级标题字号", [("levels", "1", "size"), ("spec", "h1", "size"),
                     ("levels", "1", "sz"), ("headings", "1", "size")], "text"),
    ("一级标题对齐", [("levels", "1", "align"), ("spec", "h1", "align")], "align"),
    ("二级标题字体", [("levels", "2", "zh_font"), ("spec", "h2", "zh_font"),
                     ("spec", "h2", "font"), ("headings", "2", "font")], "text"),
    ("二级标题字号", [("levels", "2", "size"), ("spec", "h2", "size"),
                     ("levels", "2", "sz"), ("headings", "2", "size")], "text"),
    ("三级标题字体", [("levels", "3", "zh_font"), ("spec", "h3", "zh_font"),
                     ("spec", "h3", "font"), ("headings", "3", "font")], "text"),
    ("三级标题字号", [("levels", "3", "size"), ("spec", "h3", "size"),
                     ("levels", "3", "sz"), ("headings", "3", "size")], "text"),
    ("页边距 上(厘米)", [("spec", "page", "top_cm"), ("page", "top")], "text"),
    ("页边距 下(厘米)", [("spec", "page", "bottom_cm"), ("page", "bottom")], "text"),
    ("页边距 左(厘米)", [("spec", "page", "left_cm"), ("page", "left")], "text"),
    ("页边距 右(厘米)", [("spec", "page", "right_cm"), ("page", "right")], "text"),
    ("参考文献格式", [("levels", "reference"), ("spec", "reference"),
                     ("refExample",), ("reference",)], "ref"),
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
        self.root = root
        self.root.title("论文格式医生 · 桌面版")
        self.root.geometry("900x640")
        self.root.minsize(800, 580)
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

        self._build_style()
        self._build_widgets()

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

        # 顶部标题区
        header = tk.Frame(self.root, bg=PAPER)
        header.pack(fill="x", padx=20, pady=(16, 8))
        tk.Label(header, text="论 文 格 式 医 生", bg=PAPER, fg=INK,
                 font=F_TITLE).pack(anchor="center")
        tk.Label(header, text="THESIS FORMAT DOCTOR · 高校论文格式规范引擎",
                 bg=PAPER, fg="#8b8378", font=F_SUBTITLE).pack(anchor="center", pady=(5, 0))
        tk.Frame(self.root, bg=CINNABAR, height=2).pack(fill="x", padx=20)

        # 主体两栏（文件选择 | 处理步骤）
        main = tk.Frame(self.root, bg=PAPER)
        main.pack(fill="both", expand=True, padx=20, pady=14)
        main.columnconfigure(0, weight=1)
        main.columnconfigure(1, weight=1)
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
        tk.Label(footer, text="论文格式医生 · 桌面版 — 完全离线，文件不会上传任何服务器",
                 bg=PAPER, fg=MUTED, font=F_FOOT).pack()
        tk.Label(footer,
                 text="© 2026 论文格式医生 · 高校批量授权 & 期刊格式定制 · 合作联系：reedskill@126.com",
                 bg=PAPER, fg=MUTED, font=F_FOOT).pack(pady=(3, 0))

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
        self.profile_info_var = tk.StringVar(value="已载入格式画像")
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

        self.status_dot = tk.Label(card, text="●", bg=PANEL, fg=MUTED, font=F_BODY)
        self.status_dot.pack(side="left", padx=(14, 6), pady=(8, 4))
        self.status_lbl = tk.Label(card, textvariable=self.status_var, bg=PANEL, fg=INK,
                                   font=F_STAT)
        self.status_lbl.pack(side="left", pady=(8, 4))

        btn_row = tk.Frame(card, bg=PANEL)
        btn_row.pack(fill="x", padx=14, pady=(10, 14))
        self._prev_btn = ttk.Button(btn_row, text="上一步", style="TButton",
                                    command=self._go_prev)
        self._prev_btn.pack(side="left")
        self._next_btn = ttk.Button(btn_row, text="下一步", style="Primary.TButton",
                                    command=self._run_step)
        self._next_btn.pack(side="right")

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
            self._next_btn.config(text="再处理一篇", command=self._reset_wizard, state="normal")
            self._prev_btn.config(text="上一步", state="disabled", command=self._go_prev)
        elif self.step_index == n - 1:
            # 第三步：一键修正 → 保存修正后论文 → 完成（由 _fix_phase 驱动）
            if self._fix_phase == "idle":
                self._next_btn.config(text="一键修正", command=self._run_step,
                                      state="disabled" if self.running else "normal")
                self._prev_btn.config(text="上一步",
                                      state="disabled" if (self.running or self.step_index == 0) else "normal",
                                      command=self._go_prev)
            elif self._fix_phase == "fixed":
                self._next_btn.config(text="保存修正后论文", command=self._export_fix,
                                      state="disabled" if self.running else "normal")
                self._prev_btn.config(text="上一步", state="normal", command=self._go_prev)
            else:  # saved
                self._next_btn.config(text="完成", state="disabled")
                self._prev_btn.config(text="再处理一篇", state="normal",
                                      command=self._reset_wizard)
        else:
            self._next_btn.config(text="下一步", command=self._run_step,
                                  state="disabled" if self.running else "normal")
            self._prev_btn.config(text="上一步",
                                  state="disabled" if (self.running or self.step_index == 0) else "normal",
                                  command=self._go_prev)

    def _set_bar(self, state, hint=None):
        """底部状态栏：idle / running / done / error。"""
        cmap = {"idle": OKC, "running": RUN, "done": OKC, "error": ERRC}
        tmap = {
            "idle": ("就绪 · 论文格式医生 v%s · 完全离线" % APP_VERSION, "请按步骤操作"),
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
                   "请把这段信息与该日志文件发给客服，以便定位原因。" % (err, log_path))
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
        if self.step_index == len(self.step_defs) - 1 and self._fix_phase != "idle":
            self._fix_phase = "idle"
            self._errored = False
            self._set_status("请按步骤操作", MUTED)
            self._set_bar("idle")
            self._refresh_wizard()
            return
        self.step_index -= 1
        self._errored = False
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
                    self._do_fix(src, docx_path, dst)
                    self.root.after(0, lambda: self._on_step_done(idx, mode))
        except Exception as e:
            self._debug("[错误] " + str(e))
            self.root.after(0, lambda: self._on_step_error(idx, mode, str(e)))
        finally:
            self.running = False
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
        self.root.after(0, self._update_profile_box)
        return out

    def _ensure_profile_ready(self):
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
            return None
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
                    # 确认页显示为磅（_field_value ÷2），写回需还原为半磅（×2）
                    try:
                        pt = float(text)
                    except (TypeError, ValueError):
                        continue  # 非法输入：保持提取结果
                    edits.append((candidates[0], str(int(round(pt * 2)))))
                elif k in _NUM_FIELDS:
                    try:
                        float(text)
                    except (TypeError, ValueError):
                        continue  # 非法输入：保持提取结果
                    edits.append((candidates[0], text))
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

    def _do_fix(self, src, docx_path, dst):
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
            report_docx=rep, add_comments=False)
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
            messagebox.showerror("导出失败", str(e))
            return
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

    def _update_profile_box(self):
        p = self.profile_path.get().strip()
        if p and os.path.isfile(p):
            self.profile_info_var.set("已载入格式画像：" + os.path.basename(p))
            if not self.profile_box.winfo_ismapped():
                self.profile_box.pack(fill="x", padx=14, pady=(4, 8), after=self._template_box)
        else:
            if self.profile_box.winfo_ismapped():
                self.profile_box.pack_forget()

    def _reset_wizard(self):
        """“再处理一篇”：回到第 1 步并清空选择。"""
        self.step_index = 0
        self._errored = False
        self._profile_confirmed = False
        self._check_report_saved = None
        self._fix_saved = set()
        self._fix_out = None
        self._fix_phase = "idle"
        self.thesis_path.set("")
        self.template_path.set("")
        self.profile_path.set("")
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


def show_activation(root):
    """激活窗口：卡密通在线激活（主） + 离线备用码（兜底）。返回是否成功。"""
    result = {"ok": False}
    top = tk.Toplevel(root)
    top.title("激活 · 论文格式医生")
    top.configure(bg=PAPER)
    top.resizable(False, False)
    top.geometry(_geo(600, 550))

    tk.Label(top, text="激 活 论 文 格 式 医 生", bg=PAPER, fg=INK,
             font=F_TITLE).pack(pady=(18, 4))
    tk.Label(top, text="请输入您购买的卡密以激活；激活仅需联网一次，之后完全离线使用。",
             bg=PAPER, fg=MUTED, font=F_SMALL, wraplength=440, justify="center").pack(pady=(0, 12))

    card_var = tk.StringVar()
    tk.Entry(top, textvariable=card_var, width=40, font=F_BODY,
             relief="solid", bd=1, justify="center").pack(pady=(4, 6))

    msg_var = tk.StringVar()
    tk.Label(top, textvariable=msg_var, bg=PAPER, fg=ERRC, font=F_SMALL).pack(pady=(0, 6))

    def do_activate():
        card = card_var.get().strip()
        if not card:
            msg_var.set("请输入卡密")
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
                ok, _days, note = license.verify_via_kami(card, mc)
                license._log("gui: 联网验证返回 ok=%s note=%s" % (ok, note))
                q.put(("done", ok, note, mc))
            except Exception:
                import traceback
                tb = traceback.format_exc()
                license._log("gui: 激活线程异常\n%s" % tb)
                q.put(("done", False, "激活过程出错，请重试或联系客服", ""))

        threading.Thread(target=work, daemon=True).start()

        def poll():
            # 主线程轮询队列（线程安全），拿到结果才更新界面
            try:
                _kind, ok, note, mc = q.get_nowait()
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
                license.save_local_license(card, mc)
                license._log("gui: 授权已写入本机")
                result["ok"] = True
                top.destroy()
            else:
                msg_var.set(note)

        top.after(300, poll)

    btn_activate = ttk.Button(top, text="激活", style="Primary.TButton",
               command=do_activate)
    btn_activate.pack(pady=(2, 4))

    tk.Label(top, text="— 以下为特殊情形使用 —", bg=PAPER, fg=MUTED, font=F_SMALL).pack(pady=(10, 4))
    mc = license.get_machine_code()
    tk.Label(top, text="本机机器码：" + mc, bg=PAPER, fg=MUTED, font=F_MONO).pack()
    ttk.Button(top, text="复制机器码", style="Ghost.TButton",
               command=lambda: top.clipboard_append(mc)).pack(pady=(3, 6))

    off_var = tk.StringVar()
    tk.Label(top, text="离线激活码（网络不通时，联系客服获取）：", bg=PAPER, fg=MUTED,
             font=F_SMALL).pack(pady=(4, 2))
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
            result["ok"] = True
            top.destroy()
        else:
            msg_var.set("离线激活码无效，请核对后重试")

    ttk.Button(top, text="使用离线激活码激活", style="Ghost.TButton",
               command=do_offline).pack(pady=(2, 6))
    ttk.Button(top, text="退出", command=lambda: top.destroy()).pack(pady=(2, 4))

    tk.Label(top, text="© 2026 论文格式医生 · 高校批量授权 & 期刊格式定制 · 合作联系：reedskill@126.com",
             bg=PAPER, fg=MUTED, font=F_FOOT, wraplength=560).pack(pady=(16, 12))

    top.wait_window()
    return result["ok"]


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
    if not license.check_local_valid():
        if not show_activation(root):
            root.destroy()
            return
    root.deiconify()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
