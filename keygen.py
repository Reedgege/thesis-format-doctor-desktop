# -*- coding: utf-8 -*-
"""
卖家发码工具（兜底用）· GUI 版 · 论文格式医生
==================================================
仅在「卡密通跑路 / 宕机」时给客户重激活使用。

用法：双击 keygen.exe（打包后），或 python keygen.py
1) 让客户在软件激活页点击「复制机器码」，把机器码发给你
2) 粘贴到下方输入框，点「生成离线备用码」
3) 把生成的码发给客户，他在激活页选「离线激活」粘贴即可
4) 自动记入 keygen_log.txt 台账

安全：离线码用对称签名（密钥与 tfd_app/license.py 共用）。
"""
import os
import sys
import time

import tkinter as tk
from tkinter import messagebox

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

# 直接加载 license.py（独立文件，无包依赖），开发 / 打包后均可运行，
# 且避免 PyInstaller 打包时把整个 tfd_app 包（含主程序引擎）拖进来。
import importlib.util
_LIC_PATH = os.path.join(HERE, "tfd_app", "license.py")
_spec = importlib.util.spec_from_file_location("tfd_license", _LIC_PATH)
lic = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(lic)

# 学术文艺风配色（简版，与主程序呼应）
BG = "#f7f3e9"        # 宣纸底
INNER = "#fffdf7"     # 内层面板
FG = "#2b2b2b"        # 墨色
MUTE = "#6b5d4f"      # 淡墨
ACCENT = "#9a3b34"    # 朱砂（主按钮）
LINE = "#d8cdb8"      # 细线


def build_ui(root):
    root.title("论文格式医生 · 离线发码工具")
    root.geometry("560x470")
    root.resizable(False, False)
    root.configure(bg=BG)
    try:
        root.iconbitmap(os.path.join(HERE, "tfd_app", "assets", "icon.ico"))
    except Exception:
        pass

    tk.Label(root, text="离 线 发 码 工 具", bg=BG, fg=ACCENT,
             font=("楷体", 18, "bold")).pack(pady=(14, 0))
    tk.Label(root, text="（兜底用 · 仅当卡密通跑路 / 宕机时给客户重激活）",
             bg=BG, fg=MUTE, font=("微软雅黑", 9)).pack(pady=(2, 8))

    tk.Label(root,
        text="① 让客户在软件激活页点「复制机器码」，把机器码发给你\n"
             "② 粘贴到下方，点「生成离线备用码」\n"
             "③ 把生成的码发给客户，他在激活页选「离线激活」粘贴即可",
        bg=BG, fg=FG, font=("微软雅黑", 10), justify="left"
    ).pack(padx=24, anchor="w")

    frm = tk.Frame(root, bg=BG)
    frm.pack(padx=24, pady=(10, 4), fill="x")
    tk.Label(frm, text="客户的机器码：", bg=BG, fg=FG,
             font=("微软雅黑", 10)).pack(anchor="w")
    entry = tk.Entry(frm, font=("Consolas", 11), bd=1, relief="solid",
                     bg=INNER, fg=FG, insertbackground=ACCENT)
    entry.pack(fill="x", pady=(4, 0), ipady=4)

    btnf = tk.Frame(root, bg=BG)
    btnf.pack(padx=24, pady=(6, 4), fill="x")
    gen_btn = tk.Button(btnf, text="生成离线备用码", bg=ACCENT, fg="white",
                        font=("微软雅黑", 11, "bold"), relief="flat",
                        activebackground="#7e2f29", padx=14, pady=6,
                        cursor="hand2")
    gen_btn.pack(side="left")

    tk.Label(root, text="生成的离线备用码（已自动复制，可直接发给客户）：",
             bg=BG, fg=MUTE, font=("微软雅黑", 9)
             ).pack(padx=24, pady=(12, 2), anchor="w")
    res_text = tk.Text(root, height=4, font=("Consolas", 11),
                       bg=INNER, fg=FG, relief="solid", bd=1,
                       wrap="word", state="disabled")
    res_text.pack(padx=24, fill="x")

    status = tk.Label(root, text="", bg=BG, fg=MUTE, font=("微软雅黑", 9))
    status.pack(padx=24, pady=(8, 0), anchor="w")

    def generate():
        target = entry.get().strip()
        if not target:
            messagebox.showwarning("提示", "请先粘贴客户的机器码")
            return
        try:
            code = lic.generate_offline_code(target)
        except Exception as e:
            messagebox.showerror("生成失败", str(e))
            return
        res_text.configure(state="normal")
        res_text.delete("1.0", "end")
        res_text.insert("end", code)
        res_text.configure(state="disabled")
        try:
            root.clipboard_clear()
            root.clipboard_append(code)
        except Exception:
            pass
        log_path = os.path.join(HERE, "keygen_log.txt")
        try:
            with open(log_path, "a", encoding="utf-8") as f:
                f.write("%s | 客户机器码 %s | 离线码 %s\n" % (
                    time.strftime("%Y-%m-%d %H:%M"), target, code))
        except Exception:
            pass
        status.configure(
            text="✓ 已生成并复制到剪贴板，已记入台账 keygen_log.txt")

    gen_btn.configure(command=generate)


def main():
    root = tk.Tk()
    build_ui(root)
    root.mainloop()


if __name__ == "__main__":
    main()
