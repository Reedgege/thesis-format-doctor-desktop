#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
论文格式医生 · 桌面版入口
========================
启动原生 Tkinter 窗口。若运行环境缺少 tkinter（极少数情况），
用 Windows API 弹窗告知用户，而非静默退出。
"""
import sys
import os

HERE = os.path.dirname(os.path.abspath(__file__))
if os.path.join(HERE, "tfd_app") not in sys.path:
    sys.path.insert(0, os.path.join(HERE, "tfd_app"))


def _show_error_win(msg):
    """tkinter 缺失时，用 ctypes 弹一个系统对话框。"""
    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(0, msg, "论文格式医生", 0x10)
    except Exception:
        print(msg, file=sys.stderr)


def main():
    try:
        import tkinter  # noqa: F401  尽早发现 tkinter 是否可用
    except ImportError:
        _show_error_win(
            "无法启动：当前环境缺少 Tkinter 图形库。\n\n"
            "请使用官方安装包重新安装 Python（勾选 tcl/tk 组件），\n"
            "或下载本软件官方发布的 Windows / macOS / Linux 版本。"
        )
        sys.exit(1)

    # ---------------------------------------------------------------
    # 界面选择：默认旧界面（稳定、为 tkinter 原生能力量身）。
    # --new-ui 才启用新版（国风）；新版若构造失败，回退旧界面并落盘日志（不静默）。
    # 已起来的窗口后续崩了不再重开一个（否则双窗）。
    # ---------------------------------------------------------------
    if "--new-ui" in sys.argv:
        try:
            from tfd_app.app_ui import App
            app = App()
        except Exception:
            import traceback
            tb = traceback.format_exc()
            sys.stderr.write("[main] 新版界面启动失败，回退旧界面：\n" + tb + "\n")
            try:
                _log_crash(tb)
            except Exception:
                pass
            from tfd_app import gui
            gui.main()
            return
        app.run()
        return

    from tfd_app import gui
    gui.main()


def _log_crash(tb):
    """把启动失败原因落盘（不静默；出问题能查）。"""
    try:
        d = os.path.join(os.path.expanduser("~"), ".thesis_format_doctor")
        os.makedirs(d, exist_ok=True)
        import time
        with open(os.path.join(d, "ui_start_fail.log"), "a", encoding="utf-8") as f:
            f.write("\n===== %s =====\n%s" % (time.strftime("%Y-%m-%d %H:%M:%S"), tb))
    except Exception:
        pass



if __name__ == "__main__":
    main()
