# -*- coding: utf-8 -*-
"""
论文格式医生 · 原生 Tkinter 界面（客户安心版）
==============================================
面向客户的简洁流程，不展示任何执行细节：

  壹 · 选择文件        —— 待处理论文（必选）+ 学校模板（可选）
  贰 · 按步骤操作      —— ① 提取学校模板要求 → ② 论文格式检查 → ③ 按学校要求一键修正
  叁 · 处理状态        —— 只显示友好状态与结果确认，无原始日志

交互约定：
  - 「提取学校模板要求」完成后，弹出“学校模板要求”确认页（可修改关键项），
    客户确认后才生效；放弃则本次不使用画像。
  - 「按学校要求一键修正」导出地址由客户选择；完成后弹确认页，
    并附检查报告 + 修改报告。
  - .doc / WPS .wps 一律先自动转 .docx 再处理（进程内转换，无黑框、不弹窗），
    原文件绝不被修改。
"""
import os
import sys
import json
import threading
import subprocess

# 把 tfd_app 加入搜索路径（打包成 exe 后 sys.path 已含，但开发态下保险）
HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from . import engine


ICON = os.path.join(HERE, "assets", "icon.png")

# ---------------------------------------------------------------------------
# 配色：宣纸 / 墨 / 黛蓝 / 朱砂 —— 学术范 + 文艺感
# ---------------------------------------------------------------------------
PAPER   = "#f6f2ea"   # 宣纸底
PANEL   = "#fcfaf5"   # 面板（微亮的纸）
INK     = "#33302b"   # 墨色文字
MUTED   = "#8b8378"   # 灰褐（次要文字）
LINE    = "#d9d2c2"   # 细线
ACCENT  = "#46586f"   # 黛蓝（章节标题 / 强调）
ACCENT_D= "#34455a"
CINNABAR= "#9e4233"   # 朱砂（主按钮）
CINNABAR_D = "#8a382b"
OKC     = "#5f7d5c"   # 完成（墨绿）
ERRC    = "#a0402f"   # 出错（朱红）
RUN     = ACCENT      # 处理中（黛蓝）

# 字体：楷体做标题（学术书卷气），正文用雅黑保证清晰
F_TITLE = ("KaiTi", 21, "bold")
F_SUB   = ("Microsoft YaHei UI", 9)
F_HDR   = ("KaiTi", 13, "bold")
F_BODY  = ("Microsoft YaHei UI", 10)
F_SMALL = ("Microsoft YaHei UI", 9)
F_BTN   = ("Microsoft YaHei UI", 11, "bold")
F_STAT  = ("KaiTi", 13)

# 模板要求确认页里可修改的关键项：(中文标签, profile 内路径)
EDIT_FIELDS = [
    ("正文字体",          ("spec", "body", "font")),
    ("正文字号",          ("spec", "body", "size")),
    ("正文行距(磅)",      ("spec", "body", "line_val")),
    ("首行缩进(字符)",    ("spec", "body", "indent_chars")),
    ("一级标题字体",      ("spec", "h1", "zh_font")),
    ("一级标题字号",      ("spec", "h1", "size")),
    ("二级标题字体",      ("spec", "h2", "zh_font")),
    ("二级标题字号",      ("spec", "h2", "size")),
    ("三级标题字体",      ("spec", "h3", "zh_font")),
    ("三级标题字号",      ("spec", "h3", "size")),
    ("页边距 上(厘米)",   ("spec", "page", "top_cm")),
    ("页边距 下(厘米)",   ("spec", "page", "bottom_cm")),
    ("页边距 左(厘米)",   ("spec", "page", "left_cm")),
    ("页边距 右(厘米)",   ("spec", "page", "right_cm")),
]
# 页边距字段：编辑厘米值时同步写 twips（top/bottom/left/right），兼容两套读取方
_MARGIN_TWIPS = {
    ("spec", "page", "top_cm"): "top",
    ("spec", "page", "bottom_cm"): "bottom",
    ("spec", "page", "left_cm"): "left",
    ("spec", "page", "right_cm"): "right",
}


def _base_no_ext(path):
    """去掉扩展名，返回用于拼输出文件名的基底。"""
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


def _profile_summary(profile):
    """把格式画像渲染成客户看得懂的“学校模板要求”文本行。"""
    lines = []
    spec = profile.get("spec") or {}
    page = spec.get("page") or profile.get("page") or {}
    if page.get("top_cm") is not None:
        lines.append("· 页边距：上 %s 下 %s 左 %s 右 %s（厘米）" % (
            page.get("top_cm"), page.get("bottom_cm"),
            page.get("left_cm"), page.get("right_cm")))
    elif page.get("top") is not None:
        lines.append("· 页边距：上 %s 下 %s 左 %s 右 %s（已提取）" % (
            page.get("top"), page.get("bottom"),
            page.get("left"), page.get("right")))

    body = spec.get("body") or {}
    if body:
        parts = []
        if body.get("font"):
            parts.append("字体 %s" % body["font"])
        if body.get("size"):
            parts.append("字号 %s" % body["size"])
        if body.get("indent_chars"):
            parts.append("首行缩进 %s 字符" % body["indent_chars"])
        if body.get("line_val"):
            parts.append("行距 %s 磅" % body["line_val"])
        lines.append("· 正文：%s" % ("，".join(parts) if parts else "样式已提取"))

    align_map = {"center": "居中", "left": "左对齐", "right": "右对齐", "both": "两端对齐"}
    for lv, name in (("h1", "一级标题"), ("h2", "二级标题"), ("h3", "三级标题")):
        h = spec.get(lv) or {}
        if not h:
            continue
        parts = []
        if h.get("zh_font"):
            parts.append("字体 %s" % h["zh_font"])
        if h.get("size"):
            parts.append("字号 %s" % h["size"])
        if h.get("align"):
            parts.append("对齐 %s" % align_map.get(h["align"], h["align"]))
        if h.get("bold"):
            parts.append("加粗")
        lines.append("· %s：%s" % (name, "，".join(parts) if parts else "样式已提取"))

    for key, label in (("abstract", "摘要"), ("keywords", "关键词"), ("toc", "目录"),
                       ("reference", "参考文献"), ("title", "论文题目"),
                       ("table", "表格"), ("figure", "插图"), ("footnote", "脚注")):
        v = spec.get(key)
        if isinstance(v, dict) and v:
            s = "，".join("%s %s" % (k, val) for k, val in list(v.items())[:4])
            lines.append("· %s：%s" % (label, s))

    if not lines:
        lines.append("（未从模板提取到批注要求，将按模板样式定义进行匹配。）")
    return lines


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("论文格式医生 · 桌面版")
        self.root.geometry("900x680")
        self.root.minsize(800, 620)
        try:
            if os.path.isfile(ICON):
                self.root.iconphoto(True, tk.PhotoImage(file=ICON))
        except Exception:
            pass

        self.thesis_path = tk.StringVar()
        self.template_path = tk.StringVar()
        self.profile_path = tk.StringVar()   # 已确认的格式画像 .json
        self.status_var = tk.StringVar(value="请按左侧步骤操作")
        self.running = False
        self._msgs = []                      # 内部日志（不展示给客户）

        self._build_style()
        self._build_widgets()

    # ------------------------------------------------------------- style
    def _build_style(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("TButton", font=F_BODY, padding=(12, 7),
                        background="#eae3d3", foreground=INK)
        style.map("TButton", background=[("active", "#ded6c2"), ("disabled", "#f0ece1")])
        style.configure("Action.TButton", font=F_BTN, padding=(14, 12),
                        background="#eae3d3", foreground=INK)
        style.map("Action.TButton",
                  background=[("active", "#ded6c2"), ("disabled", "#f0ece1")])
        style.configure("Primary.TButton", font=F_BTN, padding=(14, 12),
                        background=CINNABAR, foreground="white")
        style.map("Primary.TButton",
                  background=[("active", CINNABAR_D), ("disabled", "#c8a79f")],
                  foreground=[("disabled", "#f5e9e6")])
        style.configure("Ghost.TButton", font=F_SMALL, padding=(10, 6),
                        background=PANEL, foreground=ACCENT)
        style.map("Ghost.TButton",
                  background=[("active", "#efe9db"), ("disabled", "#f4f0e6")])

    # ------------------------------------------------------------- layout
    def _build_widgets(self):
        self.root.configure(bg=PAPER)

        header = tk.Frame(self.root, bg=PAPER)
        header.pack(fill="x", padx=24, pady=(18, 8))
        tk.Label(header, text="论 文 格 式 医 生", bg=PAPER, fg=INK,
                 font=F_TITLE).pack(anchor="center")
        tk.Label(header, text="—— 按学校要求检查与修正 · 完全离线 ——",
                 bg=PAPER, fg=MUTED, font=F_SUB).pack(anchor="center", pady=(4, 0))
        tk.Frame(self.root, bg=CINNABAR, height=2).pack(fill="x", padx=24)

        main = tk.Frame(self.root, bg=PAPER)
        main.pack(fill="both", expand=True, padx=20, pady=14)
        main.columnconfigure(0, weight=3)
        main.columnconfigure(1, weight=2)
        main.rowconfigure(0, weight=1)

        left = tk.Frame(main, bg=PAPER)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 16))
        right = tk.Frame(main, bg=PAPER)
        right.grid(row=0, column=1, sticky="nsew")

        self._build_left(left)
        self._build_right(right)

        tk.Frame(self.root, bg=LINE, height=1).pack(fill="x", padx=24)
        tk.Label(self.root,
                 text="论文格式医生 · 桌面版 — 完全离线，文件不会上传任何服务器",
                 bg=PAPER, fg=MUTED, font=F_SUB).pack(fill="x", pady=(8, 10))

    # ---------------------------------------------------------- left panel
    def _build_left(self, parent):
        self._section_title(parent, "壹 · 选择文件")

        self._thesis_box, self._thesis_name = self._upload_zone(
            parent, "待处理论文",
            "Word 文档 .docx / .doc / .wps（必选）",
            "选择论文…", self._pick_input,
            placeholder="未选择 — 点击右侧按钮选择文件")

        self._template_box, self._template_name = self._upload_zone(
            parent, "学校模板",
            "用于按学校要求检查 / 修正，更贴合要求；不选则按通用规范",
            "选择模板…", self._pick_template,
            placeholder="未选择 — 不选也能用通用规范处理")

        self.profile_box = tk.Frame(parent, bg="#f0f3ec",
                                    highlightthickness=1, highlightbackground="#b7c6ae")
        tk.Label(self.profile_box, text="●", bg="#f0f3ec", fg=OKC,
                 font=F_SMALL).pack(side="left", padx=(12, 4), pady=8)
        self.profile_info_var = tk.StringVar(value="已载入格式画像")
        tk.Label(self.profile_box, textvariable=self.profile_info_var, bg="#f0f3ec",
                 fg="#4c5f49", font=F_SMALL).pack(side="left", fill="x", expand=True)
        ttk.Button(self.profile_box, text="清除", width=6, style="Ghost.TButton",
                   command=self._clear_profile).pack(side="right", padx=8, pady=4)

        self._section_title(parent, "贰 · 按步骤操作")
        self._btns = []
        steps = [
            ("①", "提取学校模板要求", "从学校模板生成格式要求，可确认 / 修改", "profile", False),
            ("②", "论文格式检查", "按通用或学校要求检查，出检查报告", "check", False),
            ("③", "按学校要求一键修正", "修正后自选保存位置，出具检查 + 修改报告", "fix", True),
        ]
        for num, title, desc, mode, primary in steps:
            tk.Label(parent, text=num + " " + title, bg=PAPER, fg=ACCENT,
                     font=("KaiTi", 12, "bold")).pack(anchor="w", pady=(8, 0))
            btn = ttk.Button(parent, text=desc,
                             style="Primary.TButton" if primary else "Action.TButton",
                             command=lambda m=mode: self._run_mode(m))
            btn.pack(fill="x", pady=3)
            self._btns.append(btn)

        tk.Label(parent,
                 text="建议顺序：先「提取学校模板要求」（没有学校模板可跳过），再检查，最后修正。",
                 bg=PAPER, fg=MUTED, font=F_SMALL, justify="left",
                 wraplength=400).pack(anchor="w", pady=(8, 0))

    def _section_title(self, parent, text):
        row = tk.Frame(parent, bg=PAPER)
        row.pack(anchor="w", pady=(12, 4))
        tk.Frame(row, bg=ACCENT, width=4, height=16).pack(side="left", padx=(0, 8))
        tk.Label(row, text=text, bg=PAPER, fg=ACCENT, font=F_HDR).pack(side="left")

    def _upload_zone(self, parent, title, desc, btn_text, cmd, placeholder):
        box = tk.Frame(parent, bg=PANEL, highlightthickness=1, highlightbackground=LINE)
        box.pack(fill="x", pady=5)
        tk.Label(box, text=title, bg=PANEL, fg=INK,
                 font=("KaiTi", 12, "bold")).pack(anchor="w", padx=14, pady=(10, 0))
        tk.Label(box, text=desc, bg=PANEL, fg=MUTED,
                 font=F_SMALL).pack(anchor="w", padx=14, pady=(2, 0))
        row = tk.Frame(box, bg=PANEL)
        row.pack(fill="x", padx=14, pady=(8, 10))
        name = tk.Label(row, text=placeholder, bg=PANEL, fg=MUTED,
                        font=F_SMALL, anchor="w")
        name.pack(side="left", fill="x", expand=True, padx=(0, 8))
        ttk.Button(row, text=btn_text, command=cmd, style="Ghost.TButton").pack(side="right")
        return box, name

    # --------------------------------------------------------- right panel
    def _build_right(self, parent):
        self._section_title(parent, "叁 · 处理状态")

        card = tk.Frame(parent, bg=PANEL, highlightthickness=1, highlightbackground=LINE)
        card.pack(fill="x", pady=(0, 8))
        self.status_dot = tk.Label(card, text="●", bg=PANEL, fg=MUTED, font=F_BODY)
        self.status_dot.pack(side="left", padx=(14, 6), pady=16)
        self.status_lbl = tk.Label(card, textvariable=self.status_var, bg=PANEL, fg=INK,
                                   font=F_STAT)
        self.status_lbl.pack(side="left", pady=16)

        self.progress = ttk.Progressbar(parent, mode="indeterminate")
        self.progress.pack(fill="x", pady=(0, 8))
        self.progress.pack_forget()

        tk.Label(parent,
                 text="所有处理均在您本机完成，原文件绝不会被改动，请安心等待。",
                 bg=PAPER, fg=MUTED, font=F_SMALL, justify="left",
                 wraplength=300).pack(anchor="w")

        self.result_lbl = tk.Label(parent, text="", bg=PAPER, fg=OKC,
                                   font=F_BODY, justify="left", wraplength=300,
                                   anchor="w")
        self.result_lbl.pack(anchor="w", pady=(12, 0))

    # --------------------------------------------------------------- picks
    def _pick_input(self):
        p = filedialog.askopenfilename(
            title="选择待处理论文",
            filetypes=[("Word 文档", "*.docx *.doc *.wps"), ("所有文件", "*.*")])
        if p:
            self.thesis_path.set(p)
            self._thesis_name.config(text=os.path.basename(p), fg=INK)

    def _pick_template(self):
        p = filedialog.askopenfilename(
            title="选择学校模板",
            filetypes=[("Word 文档", "*.docx *.doc *.wps"), ("所有文件", "*.*")])
        if p:
            self.template_path.set(p)
            self._template_name.config(text=os.path.basename(p), fg=INK)

    def _clear_profile(self):
        self.profile_path.set("")
        self._update_profile_box()

    # ---------------------------------------------------------------- run
    def _run_mode(self, mode):
        if self.running:
            return
        if mode == "profile":
            src = self.template_path.get().strip() or self.thesis_path.get().strip()
            if not src or not os.path.isfile(src):
                messagebox.showerror("缺少输入", "请先选择“学校模板”（没有模板也可选待处理论文）。")
                return
        else:
            src = self.thesis_path.get().strip()
            if not src or not os.path.isfile(src):
                messagebox.showerror("缺少输入", "请先选择“待处理论文”。")
                return

        dst = None
        if mode == "fix":
            base = _base_no_ext(src)
            dst = filedialog.asksaveasfilename(
                title="选择修正后论文的保存位置",
                initialfile=os.path.basename(base) + "_已修正.docx",
                initialdir=os.path.dirname(base) or None,
                defaultextension=".docx",
                filetypes=[("Word 文档", "*.docx")])
            if not dst:
                return

        self.running = True
        self._set_running(True)
        status = {"profile": "正在提取学校模板要求…",
                  "check": "正在检查论文格式…",
                  "fix": "正在按学校要求修正论文…"}[mode]
        self._set_status(status, RUN)
        threading.Thread(target=self._worker, args=(mode, src, dst), daemon=True).start()

    def _worker(self, mode, src, dst=None):
        try:
            if mode == "profile":
                self._do_profile(src)
            else:
                docx_path, note = engine.normalize_input(src)
                if note:
                    self._debug(note)
                if mode == "check":
                    self._do_check(src, docx_path)
                elif mode == "fix":
                    self._do_fix(src, docx_path, dst)
            self._set_status("已完成", OKC)
        except Exception as e:
            self._debug("[错误] " + str(e))
            self._set_status("未能完成，请查看提示", ERRC)
            self.root.after(0, lambda: messagebox.showerror("处理出错", str(e)))
        finally:
            self.running = False
            self._set_running(False)

    def _set_running(self, running):
        def _apply():
            state = "disabled" if running else "normal"
            for b in self._btns:
                b.config(state=state)
            if running:
                self.progress.pack(fill="x", pady=(0, 8))
                self.progress.start(12)
            else:
                self.progress.stop()
                self.progress.pack_forget()
        self.root.after(0, _apply)

    # --------------------------------------------------- profile 提取与确认
    def _extract_profile(self, src):
        """提取学校模板要求并弹确认页（客户可修改）。返回画像路径或 None（放弃）。"""
        t_docx, note = engine.normalize_input(src)
        if note:
            self._debug(note)
        out = _base_no_ext(src) + "_格式画像.json"
        text = engine.run_build_profile(t_docx, out)
        try:
            profile = json.loads(text)
        except Exception:
            profile = None
        if profile is None:
            self.profile_path.set("")
            return None
        ok, edits = self._ask_profile_confirm(profile)
        if not ok:
            self.profile_path.set("")
            return None
        if edits:
            for path_keys, value in edits:
                _deep_set(profile, path_keys, value)
                twips_key = _MARGIN_TWIPS.get(tuple(path_keys))
                if twips_key:
                    try:
                        _deep_set(profile, ("spec", "page", twips_key),
                                  str(int(float(value) * 567)))
                    except (ValueError, TypeError):
                        pass
            with open(out, "w", encoding="utf-8") as f:
                json.dump(profile, f, ensure_ascii=False, indent=2)
        self.profile_path.set(out)
        self.root.after(0, self._update_profile_box)
        return out

    def _ensure_profile_ready(self):
        """检查 / 修正用的画像：优先已确认的，其次自动从学校模板提取（带确认页）。"""
        p = self.profile_path.get().strip()
        if p and os.path.isfile(p):
            return p
        t = self.template_path.get().strip()
        if t and os.path.isfile(t):
            return self._extract_profile(t)
        return None

    def _ask_profile_confirm(self, profile):
        """在 worker 线程里等客户确认（弹窗在主线程）。返回 (ok, edits)。"""
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
        ev.wait(timeout=600)
        return box.get("ok", False), box.get("edits", [])

    def _profile_confirm_dialog(self, profile):
        """“学校模板要求 · 请确认”弹窗：只读摘要 + 可修改关键项。返回 (ok, edits)。"""
        result = {"ok": False, "edits": []}
        top = tk.Toplevel(self.root)
        top.title("学校模板要求 · 请确认")
        top.configure(bg=PAPER)
        top.transient(self.root)
        top.grab_set()
        top.geometry("600x640+%d+%d" % (self.root.winfo_rootx() + 90,
                                        self.root.winfo_rooty() + 30))

        tk.Label(top, text="已提取出学校模板的格式要求", bg=PAPER, fg=INK,
                 font=("KaiTi", 15, "bold")).pack(pady=(16, 2))
        tk.Label(top, text="请核对是否与学校规定一致；如有不准，可直接修改后确认。",
                 bg=PAPER, fg=MUTED, font=F_SMALL).pack(pady=(0, 8))

        sum_f = tk.Frame(top, bg=PANEL, highlightthickness=1, highlightbackground=LINE)
        sum_f.pack(fill="x", padx=18, pady=4)
        sum_txt = tk.Text(sum_f, height=6, wrap="word", bg="#fbf8f1", fg=INK,
                          font=F_SMALL, relief="flat", padx=10, pady=6)
        sum_txt.insert("1.0", "\n".join(_profile_summary(profile)))
        sum_txt.config(state="disabled")
        sum_txt.pack(fill="x")

        tk.Label(top, text="如需修正，直接修改下列项目（留空表示保持提取结果）",
                 bg=PAPER, fg=ACCENT, font=F_SMALL).pack(anchor="w", padx=20, pady=(10, 2))

        # 可滚动表单
        canvas = tk.Canvas(top, bg=PAPER, highlightthickness=0)
        vbar = ttk.Scrollbar(top, orient="vertical", command=canvas.yview)
        form = tk.Frame(canvas, bg=PAPER)
        form.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=form, anchor="nw")
        canvas.configure(yscrollcommand=vbar.set)
        canvas.pack(fill="both", expand=True, padx=(20, 0), pady=4)
        vbar.pack(side="right", fill="y")

        entries = {}
        for i, (label, path_keys) in enumerate(EDIT_FIELDS):
            tk.Label(form, text=label, bg=PAPER, fg=INK, font=F_SMALL).grid(
                row=i, column=0, sticky="e", padx=(0, 10), pady=3)
            var = tk.StringVar(value=str(_deep_get(profile, path_keys) or ""))
            tk.Entry(form, textvariable=var, width=24, font=F_SMALL,
                     relief="solid", bd=1).grid(row=i, column=1, sticky="w", pady=3)
            entries[path_keys] = var

        def on_confirm():
            edits = []
            for path_keys, var in entries.items():
                v = var.get().strip()
                if v:
                    edits.append((path_keys, v))
            result["ok"] = True
            result["edits"] = edits
            top.destroy()

        def on_cancel():
            result["ok"] = False
            top.destroy()

        btns = tk.Frame(top, bg=PAPER)
        btns.pack(pady=12)
        ttk.Button(btns, text="确认，使用此要求", style="Primary.TButton",
                   command=on_confirm).pack(side="left", padx=6)
        ttk.Button(btns, text="放弃（不使用画像）", command=on_cancel).pack(side="left", padx=6)

        top.wait_window()
        return result["ok"], result["edits"]

    # ------------------------------------------------------------ 各步骤
    def _do_profile(self, src):
        out = self._extract_profile(src)
        if out:
            self._debug("画像已保存：" + out)
            self.root.after(0, lambda: self.result_lbl.config(
                text="学校模板要求已确认并保存：\n" + out))

    def _do_check(self, src, docx_path):
        base = _base_no_ext(src)
        out_md = base + "_格式检查报告.md"
        profile = self._ensure_profile_ready()
        report = engine.run_check(docx_path, profile_path=profile, out_md=out_md)
        self._debug(report)
        self.root.after(0, lambda: self._show_check_done(out_md))

    def _do_fix(self, src, docx_path, dst):
        base_dst = _base_no_ext(dst)
        rep = base_dst + "_修改报告.docx"
        chk = base_dst + "_检查报告.md"
        profile = self._ensure_profile_ready()
        report = engine.run_fix_headings(
            docx_path, dst, profile_path=profile,
            report_docx=rep, add_comments=True)
        self._debug(report)
        check_report = engine.run_check(dst, profile_path=profile, out_md=chk)
        self._debug(check_report)
        self.root.after(0, lambda: self._show_fix_done(dst, chk, rep))

    # ------------------------------------------------------------ 确认页
    def _show_check_done(self, out_md):
        self.result_lbl.config(text="检查报告已生成，保存在原文档旁。")
        k = self._modal("格式检查完成",
                        "检查报告已保存至：\n%s\n\n如需按学校要求修正论文，请继续第③步。" % out_md,
                        [("open", "打开所在文件夹"), ("ok", "完成")])
        if k == "open":
            self._open_folder(out_md)

    def _show_fix_done(self, dst, chk, rep):
        self.result_lbl.config(text="论文已修正并保存完毕。")
        k = self._modal("修正完成",
                        "已为您保存以下文件：\n\n"
                        "① 修正后论文：%s\n② 检查报告：%s\n③ 修改报告：%s\n\n"
                        "全程在本机完成，原文件未改动。" % (dst, chk, rep),
                        [("open", "打开所在文件夹"), ("ok", "完成")])
        if k == "open":
            self._open_folder(dst)

    def _modal(self, title, text, buttons):
        """通用模态确认页，返回所选按钮 key。"""
        result = {"v": None}
        top = tk.Toplevel(self.root)
        top.title(title)
        top.configure(bg=PAPER)
        top.transient(self.root)
        top.grab_set()
        top.geometry("+%d+%d" % (self.root.winfo_rootx() + 140,
                                 self.root.winfo_rooty() + 120))
        tk.Label(top, text=text, bg=PAPER, fg=INK, font=F_BODY, justify="left",
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
                self.profile_box.pack(fill="x", pady=5, after=self._template_box)
        else:
            if self.profile_box.winfo_ismapped():
                self.profile_box.pack_forget()

    # ---------------------------------------------- 内部日志（不展示客户）
    def _debug(self, text):
        self._msgs.append(text)
        if len(self._msgs) > 200:
            self._msgs = self._msgs[-200:]

    def _set_status(self, s, color):
        self.root.after(0, lambda: (self.status_var.set(s),
                                    self.status_lbl.config(fg=color),
                                    self.status_dot.config(fg=color)))


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
