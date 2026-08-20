# -*- coding: utf-8 -*-
"""
论文格式医生 · 原生 Tkinter 界面
================================
不依赖浏览器，纯本地运行。提供四种模式：
  1. 格式体检（不修改，输出 Markdown 报告）
  2. 一键套标题样式（生成新文档 + 修改明细）
  3. 参考文献重排（生成新文档）
  4. 从学校模板生成格式画像(.json)

老格式 .doc / WPS .wps 若本机装有 LibreOffice 会自动转成 .docx 再处理；
否则给出明确提示让用户先「另存为 .docx」。
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

import engine


ICON = os.path.join(HERE, "assets", "icon.png")


def _base_no_ext(path):
    """去掉扩展名，返回用于拼输出文件名的基底。"""
    return os.path.splitext(path)[0]


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("论文格式医生 · 桌面版")
        self.root.geometry("760x600")
        try:
            if os.path.isfile(ICON):
                self.root.iconphoto(True, tk.PhotoImage(file=ICON))
        except Exception:
            pass

        self.input_path = tk.StringVar()
        self.profile_path = tk.StringVar()
        self.mode = tk.StringVar(value="check")
        self.add_comments = tk.BooleanVar(value=True)
        self.running = False

        self._build_widgets()

    # ------------------------------------------------------------------ UI
    def _build_widgets(self):
        pad = {"padx": 10, "pady": 6}

        # 目标文档
        f1 = ttk.LabelFrame(self.root, text="目标文档")
        f1.pack(fill="x", **pad)
        ttk.Entry(f1, textvariable=self.input_path, width=70).pack(
            side="left", fill="x", expand=True, padx=(0, 6))
        ttk.Button(f1, text="选择文件...", command=self._pick_input).pack(side="left")

        # 模式
        f2 = ttk.LabelFrame(self.root, text="处理模式")
        f2.pack(fill="x", **pad)
        self.mode_frame = f2
        modes = [
            ("check", "格式体检（不修改，出报告）"),
            ("fix", "一键套标题样式（生成新文档）"),
            ("ref", "参考文献重排（生成新文档）"),
            ("profile", "从模板生成格式画像(.json)"),
        ]
        for val, label in modes:
            ttk.Radiobutton(
                f2, text=label, value=val, variable=self.mode,
                command=self._on_mode_change).pack(anchor="w")

        # 画像（体检/套样式可选）
        self.f3 = ttk.LabelFrame(self.root, text="格式画像(.json，可选)")
        self.f3.pack(fill="x", **pad)
        ttk.Entry(self.f3, textvariable=self.profile_path, width=70).pack(
            side="left", fill="x", expand=True, padx=(0, 6))
        ttk.Button(self.f3, text="选择画像...",
                   command=self._pick_profile).pack(side="left")
        ttk.Label(self.f3, text="提供画像则走模板驱动；留空则用内置通用标准。").pack(
            anchor="w", pady=(4, 0))

        # 选项
        f4 = ttk.Frame(self.root)
        f4.pack(fill="x", **pad)
        ttk.Checkbutton(f4, text="在文档中写批注（套样式时）",
                        variable=self.add_comments).pack(side="left")

        # 运行
        f5 = ttk.Frame(self.root)
        f5.pack(fill="x", **pad)
        self.run_btn = ttk.Button(f5, text="开始处理", command=self._on_run)
        self.run_btn.pack(side="left")
        self.status_var = tk.StringVar(value="就绪")
        ttk.Label(f5, textvariable=self.status_var).pack(side="left", padx=(10, 0))

        # 日志
        f6 = ttk.LabelFrame(self.root, text="处理日志")
        f6.pack(fill="both", expand=True, **pad)
        self.log = tk.Text(f6, wrap="word", height=18)
        self.log.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(f6, command=self.log.yview)
        sb.pack(side="right", fill="y")
        self.log.config(yscrollcommand=sb.set)
        self.log.config(state="disabled")

        self._on_mode_change()

    def _on_mode_change(self):
        # profile 模式下，输入字段即“学校模板”，画像字段隐藏
        if self.mode.get() == "profile":
            self.f3.pack_forget()
        else:
            self.f3.pack(fill="x", padx=10, pady=6, after=self.mode_frame)

    # --------------------------------------------------------------- picks
    def _pick_input(self):
        p = filedialog.askopenfilename(
            title="选择待处理文档",
            filetypes=[("Word 文档", "*.docx *.doc *.wps"), ("所有文件", "*.*")])
        if p:
            self.input_path.set(p)

    def _pick_profile(self):
        p = filedialog.askopenfilename(
            title="选择格式画像",
            filetypes=[("JSON 画像", "*.json"), ("所有文件", "*.*")])
        if p:
            self.profile_path.set(p)

    # ---------------------------------------------------------------- run
    def _on_run(self):
        if self.running:
            return
        src = self.input_path.get().strip()
        if not src or not os.path.isfile(src):
            messagebox.showerror("缺少输入", "请先选择有效的目标文档。")
            return
        mode = self.mode.get()
        if mode == "profile":
            out = _base_no_ext(src) + "_格式画像.json"
            if os.path.exists(out):
                if not messagebox.askyesno("覆盖确认", "输出文件已存在，是否覆盖？\n" + out):
                    return
        else:
            if mode in ("fix", "ref") and os.path.exists(_base_no_ext(src) + (
                    "_已套标题样式.docx" if mode == "fix" else "_参考文献重排.docx")):
                # 仅提示，不强阻；用户可在保存对话框里改
                pass
        self.running = True
        self.run_btn.config(state="disabled")
        self.status_var.set("处理中...")
        threading.Thread(target=self._worker, args=(mode, src), daemon=True).start()

    def _worker(self, mode, src):
        try:
            docx_path, note = engine.normalize_input(src)
            if note:
                self._append(note + "\n")
            if mode == "check":
                self._do_check(src, docx_path)
            elif mode == "fix":
                self._do_fix(src, docx_path)
            elif mode == "ref":
                self._do_ref(src, docx_path)
            elif mode == "profile":
                self._do_profile(src)
            self._append("\n===== 处理完成 =====\n")
            self._set_status("完成")
        except Exception as e:
            self._append("\n[错误] " + str(e) + "\n")
            self._set_status("失败")
            self.root.after(0, lambda: messagebox.showerror("处理出错", str(e)))
        finally:
            self.running = False
            self.root.after(0, lambda: self.run_btn.config(state="normal"))

    def _do_check(self, src, docx_path):
        base = _base_no_ext(src)
        out_md = base + "_格式体检报告.md"
        profile = self.profile_path.get().strip() or None
        if profile and not os.path.isfile(profile):
            profile = None
        report = engine.run_check(docx_path, profile_path=profile, out_md=out_md)
        self._append(report)
        self._append("\n报告已保存：" + out_md + "\n")

    def _do_fix(self, src, docx_path):
        base = _base_no_ext(src)
        dst = base + "_已套标题样式.docx"
        rep = base + "_修改明细.docx"
        report = engine.run_fix_headings(
            docx_path, dst, profile_path=self.profile_path.get().strip() or None,
            report_docx=rep, add_comments=self.add_comments.get())
        self._append(report)
        self._append("\n新文档已保存：" + dst + "\n")
        self._append("修改明细已保存：" + rep + "\n")

    def _do_ref(self, src, docx_path):
        base = _base_no_ext(src)
        dst = base + "_参考文献重排.docx"
        report = engine.run_reformat_refs(docx_path, dst)
        self._append(report)
        self._append("\n新文档已保存：" + dst + "\n")

    def _do_profile(self, src, _=None):
        base = _base_no_ext(src)
        dst = base + "_格式画像.json"
        report = engine.run_build_profile(src, dst)
        self._append(report)
        self._append("\n画像已保存：" + dst + "\n")

    # ----------------------------------------------------- thread-safe UI
    def _append(self, text):
        self.root.after(0, self._append_now, text)

    def _append_now(self, text):
        self.log.config(state="normal")
        self.log.insert("end", text)
        self.log.see("end")
        self.log.config(state="disabled")

    def _set_status(self, s):
        self.root.after(0, lambda: self.status_var.set(s))


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
