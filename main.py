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

    from tfd_app import gui
    gui.main()


if __name__ == "__main__":
    main()
