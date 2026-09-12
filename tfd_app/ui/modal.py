# -*- coding: utf-8 -*-
"""
论文格式医生 · 中国版弹窗（照 upgrade-modal.svg / confirm-modal.svg）
=============================================================

ModalShell  ：弹窗壳（14 圆角、象牙面、细边、单一主操作 + 安静取消）
ConfirmModal：通用确认（取消 / 确认）
（UpgradeModal / 套餐卡片已按老板 2026-09-12 要求移除：升级入口改为「简单小程序码 + 文字」。）

全部从 ui.theme 取色，禁止第二套样式。
"""
from __future__ import annotations

import tkinter as tk
from . import theme
from .theme import get_theme, RADIUS, TYPE
from .widgets import RoundButton, _rounded_rect


def dpi_scale(widget):
    """当前屏幕的物理缩放系数（96dpi = 1.0；150% 屏 ≈ 1.5）。

    Tk 的字号按「磅」算，屏幕 DPI 一高字就整体变大；弹窗若还按固定的设计像素
    开窗，字会撑破窗口（实测 150% 缩放下升级弹窗的四张套餐卡只剩一张可见）。
    这里把弹窗的**几何尺寸**按同一比例放大，字号与窗口的比例就与 100% 时一致。
    """
    try:
        return max(1.0, widget.winfo_fpixels("1i") / 96.0)
    except Exception:
        return 1.0


class ModalShell(tk.Toplevel):
    """带圆角画布的弹窗基底。subclass 在 self.body 里放内容。"""

    def __init__(self, parent, edition="student", title="", subtitle="",
                 width=620, height=300, **kw):
        super().__init__(parent, **kw)
        self._t = get_theme(edition)
        # 设计像素 -> 物理像素（100% 缩放下 s=1.0，行为不变）
        self._s = dpi_scale(self)
        self._width = int(round(width * self._s))
        self._height = int(round(height * self._s))
        # 父窗口可见时才设 transient：启动时 root 被 withdraw，子窗若 transient 到
        # 已隐藏的父窗会被 Tk 级联隐藏/不出现，导致启动激活窗卡死。模态抓取仍保留。
        if parent is not None and parent.winfo_viewable():
            self.transient(parent)
        self.grab_set()
        self.overrideredirect(True)
        self.configure(bg=self._t.modal_border)
        self.cv = tk.Canvas(self, highlightthickness=0, bd=0,
                            width=self._width, height=self._height)
        self.cv.pack(fill="both", expand=True)
        self.inner = tk.Frame(self.cv, bg=self._t.modal_fill)
        self._win = self.cv.create_window(0, 0, window=self.inner, anchor="nw")
        self.cv.bind("<Configure>", lambda e: self._draw())
        self.after(20, self._draw)
        # 标题区
        self.head = tk.Frame(self.inner, bg=self._t.modal_fill)
        self.head.pack(fill="x", padx=int(40 * self._s), pady=(int(34 * self._s), 6))
        if title:
            tk.Label(self.head, text=title, bg=self._t.modal_fill,
                     fg=self._t.modal_title,
                     font=self._t.serif(TYPE["page_title"], bold=True)).pack(anchor="w")
        if subtitle:
            tk.Label(self.head, text=subtitle, bg=self._t.modal_fill,
                     fg=self._t.modal_sub, font=self._t.sans(TYPE["caption"])).pack(anchor="w", pady=(4, 0))
        self.body = tk.Frame(self.inner, bg=self._t.modal_fill)
        self.body.pack(fill="both", expand=True, padx=int(40 * self._s), pady=(10, 10))
        # Esc 关闭（除升级弹窗外都允许）
        self.bind("<Escape>", lambda e: self._safe_close())

    def _draw(self):
        self.cv.delete("all")
        w = self.cv.winfo_width()
        h = self.cv.winfo_height()
        if w < 2 or h < 2:
            self.after(20, self._draw)
            return
        _rounded_rect(self.cv, 1, 1, w - 1, h - 1, RADIUS["modal"],
                      fill=self._t.modal_fill, outline=self._t.modal_border, width=2)
        self._win = self.cv.create_window(0, 0, window=self.inner, anchor="nw")
        self.cv.coords(self._win, 0, 0)
        self.cv.itemconfig(self._win, width=w)

    def _safe_close(self):
        try:
            self.destroy()
        except Exception:
            pass

    def center_on(self, parent):
        self.update_idletasks()
        pw = parent.winfo_rootx()
        ph = parent.winfo_rooty()
        w = self._width
        h = self._height
        x = pw + (parent.winfo_width() - w) // 2
        y = ph + (parent.winfo_height() - h) // 2
        # 屏幕内约束：四边都不越界。否则弹窗底部（按钮 / 页脚）会被屏幕裁掉，
        # 实测 1280x720 @150% 时「升级正式版」弹窗底部按钮看不见。
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        x = max(0, min(x, sw - w - 10))
        y = max(0, min(y, sh - h - 10))
        self.geometry("%dx%d+%d+%d" % (w, h, x, y))


class ConfirmModal(ModalShell):
    """通用确认：body 放提示文字，底部 取消 / 确认。result 经 callback 回传。"""

    def __init__(self, parent, edition="student", title="确认操作",
                 message="请确认是否继续当前操作。", on_confirm=None,
                 confirm_text="确认", cancel_text="取消", **kw):
        super().__init__(parent, edition=edition, title=title, width=620, height=300)
        tk.Label(self.body, text=message, bg=self._t.modal_fill, fg=self._t.modal_sub,
                 font=self._t.sans(TYPE["body"]), wraplength=520, justify="left",
                 anchor="w").pack(anchor="w", pady=(8, 24))
        btns = tk.Frame(self.body, bg=self._t.modal_fill)
        btns.pack(fill="x", pady=(6, 10))
        cancel = RoundButton(btns, text=cancel_text, edition=edition, style="secondary",
                              height=46, font=self._t.sans(TYPE["btn"], bold=True),
                              command=self._safe_close)
        cancel.pack(side="left")
        ok = RoundButton(btns, text=confirm_text, edition=edition, style="primary",
                         height=46, font=self._t.sans(TYPE["btn"], bold=True),
                         command=lambda: (on_confirm and on_confirm(), self._safe_close()))
        ok.pack(side="right")
        self.center_on(parent)
