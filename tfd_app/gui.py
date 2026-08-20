# -*- coding: utf-8 -*-
"""
论文格式医生 · 原生 Tkinter 界面（仿网页版布局）
================================================
布局参照 v1.3.5 网页版：
  ┌ 蓝色标题栏：论文格式医生 · 桌面版 · 纯本地离线运行 ┐
  ├ 左侧：1·选择文件（待处理论文必选 + 学校模板可选两个大上传区）
  │       2·执行操作（格式体检 / 套标题 / 重排文献 / 抽取模板画像 四个大按钮）
  ├ 右侧：运行结果（状态 + 日志）
  └ 底部：完全离线，文件不会上传任何服务器 ┘

交互逻辑与网页版一致：
  - 体检 / 套用：优先用"已抽取的画像"，其次用"学校模板"自动生成画像，否则按通用规范。
  - 抽取模板画像：用"学校模板"（没选模板就用"待处理论文"）。
老格式 .doc / WPS .wps 若本机装有 LibreOffice 会自动转成 .docx 再处理。
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

# 配色（护眼浅色系，仿网页版）
BLUE = "#2563eb"
BLUE_LIGHT = "#3b82f6"
HEADER_BG = "#1e40af"
BG = "#f3f6fb"
PANEL = "#ffffff"
BORDER = "#dbe3ef"
TEXT = "#111827"
GRAY = "#6b7280"
OK = "#16a34a"
ERR = "#dc2626"
RUN = "#2563eb"


def _base_no_ext(path):
    """去掉扩展名，返回用于拼输出文件名的基底。"""
    return os.path.splitext(path)[0]


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("论文格式医生 · 桌面版")
        self.root.geometry("880x660")
        self.root.minsize(760, 580)
        try:
            if os.path.isfile(ICON):
                self.root.iconphoto(True, tk.PhotoImage(file=ICON))
        except Exception:
            pass

        self.thesis_path = tk.StringVar()
        self.template_path = tk.StringVar()
        self.profile_path = tk.StringVar()   # 已抽取/自动生成的格式画像 .json
        self.status_var = tk.StringVar(value="等待操作…")
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
        style.configure("TButton", font=("Microsoft YaHei UI", 10), padding=(12, 7))
        style.configure("Action.TButton",
                        font=("Microsoft YaHei UI", 11, "bold"), padding=(14, 12))
        style.configure("Primary.TButton",
                        background=BLUE, foreground="white",
                        font=("Microsoft YaHei UI", 11, "bold"), padding=(14, 12))
        style.map("Primary.TButton",
                  background=[("disabled", "#93c5fd"), ("active", BLUE_LIGHT)],
                  foreground=[("disabled", "#eff6ff")])
        style.map("Action.TButton",
                  background=[("disabled", "#e5e7eb"), ("active", "#dbeafe")])

    # ------------------------------------------------------------- layout
    def _build_widgets(self):
        self.root.configure(bg=BG)

        # 顶部蓝色标题栏
        header = tk.Frame(self.root, bg=HEADER_BG, height=62)
        header.pack(fill="x")
        header.pack_propagate(False)
        tk.Label(header, text="格", bg=HEADER_BG, fg="white",
                 font=("Microsoft YaHei UI", 20, "bold")).pack(side="left", padx=(16, 8))
        ht = tk.Frame(header, bg=HEADER_BG)
        ht.pack(side="left")
        tk.Label(ht, text="论文格式医生", bg=HEADER_BG, fg="white",
                 font=("Microsoft YaHei UI", 15, "bold")).pack(anchor="w")
        tk.Label(ht, text="桌面版 · 纯本地离线运行 · 文件不上传",
                 bg=HEADER_BG, fg="#bfdbfe", font=("Microsoft YaHei UI", 9)).pack(anchor="w")

        # 主体：左右两栏
        main = tk.Frame(self.root, bg=BG)
        main.pack(fill="both", expand=True, padx=14, pady=12)
        main.columnconfigure(0, weight=3)
        main.columnconfigure(1, weight=2)
        main.rowconfigure(0, weight=1)

        left = tk.Frame(main, bg=BG)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 14))
        right = tk.Frame(main, bg=BG)
        right.grid(row=0, column=1, sticky="nsew")

        self._build_left(left)
        self._build_right(right)

        # 底部
        tk.Label(self.root,
                 text="论文格式医生 · 桌面版 — 完全离线，文件不会上传任何服务器",
                 bg=BG, fg=GRAY, font=("Microsoft YaHei UI", 9)).pack(fill="x", pady=(0, 8))

    # ---------------------------------------------------------- left panel
    def _build_left(self, parent):
        self._section_title(parent, "1 · 选择文件")

        # 待处理论文（必选）
        self._thesis_box, self._thesis_name = self._upload_zone(
            parent, "📄 待处理论文",
            "Word 文档 .docx / .doc / .wps（必选）",
            "选择论文文件…", self._pick_input,
            placeholder="未选择 — 点击右侧按钮选择文件")

        # 学校模板（可选）—— 清晰明了
        self._template_box, self._template_name = self._upload_zone(
            parent, "🏫 学校模板（可选）",
            "用于“模板驱动诊断 / 套用”，结果更贴合学校要求；不选则按通用规范",
            "选择模板文件…", self._pick_template,
            placeholder="未选择 — 不选也能用通用规范处理")

        # 已载入画像提示条
        self.profile_box = tk.Frame(parent, bg="#ecfdf5",
                                    highlightthickness=1, highlightbackground="#6ee7b7")
        tk.Label(self.profile_box, text="●", bg="#ecfdf5", fg=OK,
                 font=("Microsoft YaHei UI", 10)).pack(side="left", padx=(12, 4), pady=8)
        self.profile_info_var = tk.StringVar(value="已载入格式画像")
        tk.Label(self.profile_box, textvariable=self.profile_info_var, bg="#ecfdf5",
                 fg="#065f46", font=("Microsoft YaHei UI", 9)).pack(side="left", fill="x", expand=True)
        ttk.Button(self.profile_box, text="清除", width=6,
                   command=self._clear_profile).pack(side="right", padx=8, pady=4)

        # 操作区
        self._section_title(parent, "2 · 执行操作")
        acts = tk.Frame(parent, bg=BG)
        acts.pack(fill="x", pady=(2, 0))
        acts.columnconfigure(0, weight=1)
        acts.columnconfigure(1, weight=1)

        self._btns = []
        specs = [
            (0, 0, "🩺 格式体检", "primary", "check"),
            (0, 1, "🎯 套标题 & 出目录", "normal", "fix"),
            (1, 0, "📚 重排参考文献", "normal", "ref"),
            (1, 1, "📐 抽取模板画像", "normal", "profile"),
        ]
        for r, c, text, kind, mode in specs:
            btn = ttk.Button(acts, text=text,
                             style="Primary.TButton" if kind == "primary" else "Action.TButton",
                             command=lambda m=mode: self._run_mode(m))
            btn.grid(row=r, column=c, sticky="ew", padx=3, pady=4)
            self._btns.append(btn)

        tk.Label(parent,
                 text="提示：先“📐 抽取模板画像”，后续体检 / 套用会自动带上它，结果更贴合学校要求。",
                 bg=BG, fg=GRAY, font=("Microsoft YaHei UI", 9), justify="left",
                 wraplength=360).pack(anchor="w", pady=(8, 0))

    def _section_title(self, parent, text):
        tk.Label(parent, text=text, bg=BG, fg=BLUE,
                 font=("Microsoft YaHei UI", 12, "bold")).pack(anchor="w", pady=(10, 4))

    def _upload_zone(self, parent, title, desc, btn_text, cmd, placeholder):
        box = tk.Frame(parent, bg=PANEL, highlightthickness=1, highlightbackground=BORDER)
        box.pack(fill="x", pady=5)
        tk.Label(box, text=title, bg=PANEL, fg=TEXT,
                 font=("Microsoft YaHei UI", 11, "bold")).pack(anchor="w", padx=12, pady=(10, 0))
        tk.Label(box, text=desc, bg=PANEL, fg=GRAY,
                 font=("Microsoft YaHei UI", 9)).pack(anchor="w", padx=12, pady=(2, 0))
        row = tk.Frame(box, bg=PANEL)
        row.pack(fill="x", padx=12, pady=(8, 10))
        name = tk.Label(row, text=placeholder, bg=PANEL, fg=GRAY,
                        font=("Microsoft YaHei UI", 9), anchor="w")
        name.pack(side="left", fill="x", expand=True, padx=(0, 8))
        ttk.Button(row, text=btn_text, command=cmd).pack(side="right")
        return box, name

    # --------------------------------------------------------- right panel
    def _build_right(self, parent):
        self._section_title(parent, "运行结果")

        # 状态条
        st = tk.Frame(parent, bg=PANEL, highlightthickness=1, highlightbackground=BORDER)
        st.pack(fill="x", pady=(0, 6))
        self.status_dot = tk.Label(st, text="●", bg=PANEL, fg=GRAY,
                                   font=("Microsoft YaHei UI", 10))
        self.status_dot.pack(side="left", padx=(12, 4), pady=8)
        self.status_lbl = tk.Label(st, textvariable=self.status_var, bg=PANEL, fg=GRAY,
                                   font=("Microsoft YaHei UI", 10))
        self.status_lbl.pack(side="left", pady=8)

        # 日志
        logbox = tk.Frame(parent, bg=PANEL, highlightthickness=1, highlightbackground=BORDER)
        logbox.pack(fill="both", expand=True)
        self.log = tk.Text(logbox, wrap="word", bg="#fafcff", fg=TEXT,
                           font=("Microsoft YaHei UI", 9), relief="flat", padx=10, pady=8)
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
            self._thesis_name.config(text=os.path.basename(p), fg=TEXT)

    def _pick_template(self):
        p = filedialog.askopenfilename(
            title="选择学校模板",
            filetypes=[("Word 文档", "*.docx *.doc *.wps"), ("所有文件", "*.*")])
        if p:
            self.template_path.set(p)
            self._template_name.config(text=os.path.basename(p), fg=TEXT)

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
        self.running = True
        for b in self._btns:
            b.config(state="disabled")
        self._set_status("处理中…", RUN)
        threading.Thread(target=self._worker, args=(mode, src), daemon=True).start()

    def _worker(self, mode, src):
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
                    self._do_fix(src, docx_path)
                elif mode == "ref":
                    self._do_ref(src, docx_path)
            self._append("\n===== 处理完成 =====\n")
            self._set_status("完成", OK)
        except Exception as e:
            self._append("\n[错误] " + str(e) + "\n")
            self._set_status("出错", ERR)
            self.root.after(0, lambda: messagebox.showerror("处理出错", str(e)))
        finally:
            self.running = False
            self.root.after(0, self._enable_buttons)

    def _enable_buttons(self):
        for b in self._btns:
            b.config(state="normal")

    def _resolve_profile(self):
        """体检/套用用的画像：优先已抽取的，其次用学校模板自动生成。"""
        p = self.profile_path.get().strip()
        if p and os.path.isfile(p):
            return p
        t = self.template_path.get().strip()
        if t and os.path.isfile(t):
            out = _base_no_ext(t) + "_格式画像.json"
            engine.run_build_profile(t, out)      # 自动从模板生成画像（写文件）
            self.profile_path.set(out)
            self.root.after(0, self._update_profile_box)
            self._append("已从学校模板自动生成格式画像，本次按“模板驱动”处理。\n")
            return out
        return None

    def _do_check(self, src, docx_path):
        base = _base_no_ext(src)
        out_md = base + "_格式体检报告.md"
        profile = self._resolve_profile()
        report = engine.run_check(docx_path, profile_path=profile, out_md=out_md)
        self._append(report)
        self._append("\n报告已保存：" + out_md + "\n")

    def _do_fix(self, src, docx_path):
        base = _base_no_ext(src)
        dst = base + "_已套标题样式.docx"
        rep = base + "_修改明细.docx"
        profile = self._resolve_profile()
        report = engine.run_fix_headings(
            docx_path, dst, profile_path=profile,
            report_docx=rep, add_comments=True)
        self._append(report)
        self._append("\n新文档已保存：" + dst + "\n")
        self._append("修改明细已保存：" + rep + "\n")

    def _do_ref(self, src, docx_path):
        base = _base_no_ext(src)
        dst = base + "_参考文献重排.docx"
        report = engine.run_reformat_refs(docx_path, dst)
        self._append(report)
        self._append("\n新文档已保存：" + dst + "\n")

    def _do_profile(self, src):
        base = _base_no_ext(src)
        out = base + "_格式画像.json"
        report = engine.run_build_profile(src, out)
        self.profile_path.set(out)
        self._append(report)
        self._append("\n画像已保存：" + out + "\n")
        self._append("已自动载入该画像，后续体检 / 套用会自动使用它。\n")
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
