# -*- coding: utf-8 -*-
"""
论文格式医生 · 原生 Tkinter 界面（学术文艺风）
==============================================
布局：顶部标题栏 → 左侧「壹·选择文件 / 贰·执行操作」→ 右侧「叁·运行结果」→ 底部离线提示。

功能按钮：
  1. 论文格式检查        —— 不修改原文件，出检查报告(.md)
  2. 一键修正论文        —— 导出地址由用户选择；修正后自动出具「检查报告(.md) + 修改报告(.docx)」
  3. 参考文献重排        —— 按 GB/T 7714 重排参考文献，生成新文档
  4. 学校模板格式提取    —— 从学校模板(.docx/.doc/.wps)生成格式画像(.json)，自动载入后续使用

兼容性：.doc / WPS .wps 一律先自动转成 .docx 再处理（进程内调用本机 Word/WPS，无黑框），
       原文件绝不被修改。
"""
import os
import sys
import threading

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
F_LOG   = ("Microsoft YaHei UI", 9)


def _base_no_ext(path):
    """去掉扩展名，返回用于拼输出文件名的基底。"""
    return os.path.splitext(path)[0]


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("论文格式医生 · 桌面版")
        self.root.geometry("900x680")
        self.root.minsize(780, 600)
        try:
            if os.path.isfile(ICON):
                self.root.iconphoto(True, tk.PhotoImage(file=ICON))
        except Exception:
            pass

        self.thesis_path = tk.StringVar()
        self.template_path = tk.StringVar()
        self.profile_path = tk.StringVar()   # 已抽取/自动生成的格式画像 .json
        self.status_var = tk.StringVar(value="静候指示…")
        self.running = False

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
        style.configure("Action.TButton", font=F_BTN, padding=(14, 13),
                        background="#eae3d3", foreground=INK)
        style.map("Action.TButton",
                  background=[("active", "#ded6c2"), ("disabled", "#f0ece1")])
        style.configure("Primary.TButton", font=F_BTN, padding=(14, 13),
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

        # 顶部标题（学术封面式）
        header = tk.Frame(self.root, bg=PAPER)
        header.pack(fill="x", padx=24, pady=(18, 8))
        tk.Label(header, text="论 文 格 式 医 生", bg=PAPER, fg=INK,
                 font=F_TITLE).pack(anchor="center")
        tk.Label(header, text="—— 格式体检 · 一键修正 · 完全离线 ——",
                 bg=PAPER, fg=MUTED, font=F_SUB).pack(anchor="center", pady=(4, 0))
        # 细朱线
        tk.Frame(self.root, bg=CINNABAR, height=2).pack(fill="x", padx=24)

        # 主体：左右两栏
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

        # 底部
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

        # 已载入画像提示条
        self.profile_box = tk.Frame(parent, bg="#f0f3ec",
                                    highlightthickness=1, highlightbackground="#b7c6ae")
        tk.Label(self.profile_box, text="●", bg="#f0f3ec", fg=OKC,
                 font=F_SMALL).pack(side="left", padx=(12, 4), pady=8)
        self.profile_info_var = tk.StringVar(value="已载入格式画像")
        tk.Label(self.profile_box, textvariable=self.profile_info_var, bg="#f0f3ec",
                 fg="#4c5f49", font=F_SMALL).pack(side="left", fill="x", expand=True)
        ttk.Button(self.profile_box, text="清除", width=6, style="Ghost.TButton",
                   command=self._clear_profile).pack(side="right", padx=8, pady=4)

        self._section_title(parent, "贰 · 执行操作")
        acts = tk.Frame(parent, bg=PAPER)
        acts.pack(fill="x", pady=(2, 0))
        acts.columnconfigure(0, weight=1)
        acts.columnconfigure(1, weight=1)

        self._btns = []
        specs = [
            (0, 0, "论文格式检查", "primary", "check"),
            (0, 1, "一键修正论文", "normal", "fix"),
            (1, 0, "参考文献重排", "normal", "ref"),
            (1, 1, "学校模板格式提取", "normal", "profile"),
        ]
        for r, c, text, kind, mode in specs:
            btn = ttk.Button(acts, text=text,
                             style="Primary.TButton" if kind == "primary" else "Action.TButton",
                             command=lambda m=mode: self._run_mode(m))
            btn.grid(row=r, column=c, sticky="ew", padx=3, pady=4)
            self._btns.append(btn)

        tk.Label(parent,
                 text="提示：可先「学校模板格式提取」生成格式画像，之后检查 / 修正会自动带上它，结果更贴合学校要求。",
                 bg=PAPER, fg=MUTED, font=F_SMALL, justify="left",
                 wraplength=380).pack(anchor="w", pady=(8, 0))

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
        self._section_title(parent, "叁 · 运行结果")

        st = tk.Frame(parent, bg=PANEL, highlightthickness=1, highlightbackground=LINE)
        st.pack(fill="x", pady=(0, 6))
        self.status_dot = tk.Label(st, text="●", bg=PANEL, fg=MUTED, font=F_SMALL)
        self.status_dot.pack(side="left", padx=(12, 4), pady=8)
        self.status_lbl = tk.Label(st, textvariable=self.status_var, bg=PANEL, fg=INK,
                                   font=F_BODY)
        self.status_lbl.pack(side="left", pady=8)

        logbox = tk.Frame(parent, bg=PANEL, highlightthickness=1, highlightbackground=LINE)
        logbox.pack(fill="both", expand=True)
        self.log = tk.Text(logbox, wrap="word", bg="#fbf8f1", fg=INK,
                           font=F_LOG, relief="flat", padx=10, pady=8)
        self.log.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(logbox, command=self.log.yview)
        sb.pack(side="right", fill="y")
        self.log.config(yscrollcommand=sb.set, state="disabled")

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
            # 一键修正：导出地址由客户选择
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
        for b in self._btns:
            b.config(state="disabled")
        self._set_status("处理中…", RUN)
        threading.Thread(target=self._worker, args=(mode, src, dst), daemon=True).start()

    def _worker(self, mode, src, dst=None):
        try:
            if mode == "profile":
                self._do_profile(src)
            else:
                docx_path, note = engine.normalize_input(src)
                if note:
                    self._append(note + "\n")
                if mode == "check":
                    self._do_check(src, docx_path)
                elif mode == "fix":
                    self._do_fix(src, docx_path, dst)
                elif mode == "ref":
                    self._do_ref(src, docx_path)
            self._append("\n===== 处理完成 =====\n")
            self._set_status("完成", OKC)
        except Exception as e:
            self._append("\n[错误] " + str(e) + "\n")
            self._set_status("出错", ERRC)
            self.root.after(0, lambda: messagebox.showerror("处理出错", str(e)))
        finally:
            self.running = False
            self.root.after(0, self._enable_buttons)

    def _enable_buttons(self):
        for b in self._btns:
            b.config(state="normal")

    def _resolve_profile(self):
        """检查 / 修正用的画像：优先已抽取的，其次用学校模板自动生成。

        注意：模板可能是 .doc/.wps，必须先转成 .docx 再生成画像，否则会报错。
        """
        p = self.profile_path.get().strip()
        if p and os.path.isfile(p):
            return p
        t = self.template_path.get().strip()
        if t and os.path.isfile(t):
            t_docx, note = engine.normalize_input(t)
            if note:
                self._append(note + "\n")
            out = _base_no_ext(t) + "_格式画像.json"
            engine.run_build_profile(t_docx, out)
            self.profile_path.set(out)
            self.root.after(0, self._update_profile_box)
            self._append("已从学校模板自动生成格式画像，本次按“模板驱动”处理。\n")
            return out
        return None

    def _do_check(self, src, docx_path):
        base = _base_no_ext(src)
        out_md = base + "_格式检查报告.md"
        profile = self._resolve_profile()
        report = engine.run_check(docx_path, profile_path=profile, out_md=out_md)
        self._append(report)
        self._append("\n检查报告已保存：" + out_md + "\n")

    def _do_fix(self, src, docx_path, dst):
        """一键修正：导出到用户选择的位置，同时出具检查报告 + 修改报告。"""
        base_dst = _base_no_ext(dst)
        rep = base_dst + "_修改报告.docx"
        chk = base_dst + "_检查报告.md"
        profile = self._resolve_profile()
        report = engine.run_fix_headings(
            docx_path, dst, profile_path=profile,
            report_docx=rep, add_comments=True)
        self._append(report)
        # 对修正后的文档再出一份检查报告
        self._append("\n—— 修正后复检 ——\n")
        check_report = engine.run_check(dst, profile_path=profile, out_md=chk)
        self._append(check_report)
        self._append("\n修正后论文已保存：" + dst + "\n")
        self._append("检查报告已保存：" + chk + "\n")
        self._append("修改报告已保存：" + rep + "\n")

    def _do_ref(self, src, docx_path):
        base = _base_no_ext(src)
        dst = base + "_参考文献重排.docx"
        report = engine.run_reformat_refs(docx_path, dst)
        self._append(report)
        self._append("\n新文档已保存：" + dst + "\n")

    def _do_profile(self, src):
        """学校模板格式提取：模板可能是 .doc/.wps，先转成 .docx 再生成画像。"""
        t_docx, note = engine.normalize_input(src)
        if note:
            self._append(note + "\n")
        base = _base_no_ext(src)
        out = base + "_格式画像.json"
        report = engine.run_build_profile(t_docx, out)
        self.profile_path.set(out)
        self._append(report)
        self._append("\n画像已保存：" + out + "\n")
        self._append("已自动载入该画像，后续检查 / 修正会自动使用它。\n")
        self.root.after(0, self._update_profile_box)

    def _update_profile_box(self):
        p = self.profile_path.get().strip()
        if p and os.path.isfile(p):
            self.profile_info_var.set("已载入格式画像：" + os.path.basename(p))
            if not self.profile_box.winfo_ismapped():
                self.profile_box.pack(fill="x", pady=5, after=self._template_box)
        else:
            if self.profile_box.winfo_ismapped():
                self.profile_box.pack_forget()

    # ----------------------------------------------------- thread-safe UI
    def _append(self, text):
        self.root.after(0, self._append_now, text)

    def _append_now(self, text):
        self.log.config(state="normal")
        self.log.insert("end", text)
        self.log.see("end")
        self.log.config(state="disabled")

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
