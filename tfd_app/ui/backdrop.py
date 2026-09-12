# -*- coding: utf-8 -*-
"""
背景层：把施工包正式背景图（STUDENT_BACKGROUND.png / ADVISOR_BACKGROUND.png）
铺在窗口最底层。禁止使用完整 UI 预览图作背景（AGENT_MASTER_INSTRUCTION.md）。

实现：在 root 最底层放一个 Canvas（place 铺满），用 Tk 8.6 原生 PhotoImage 把背景图
按 cover 缩放后绘上（整数 zoom 放大 + 居中裁切，纯标准库、不引第三方依赖）。
"""
from __future__ import annotations

import math
import os
import tkinter as tk

from . import theme
from .theme import get_theme

_HERE = os.path.dirname(os.path.abspath(__file__))
_ASSETS = os.path.normpath(os.path.join(_HERE, "..", "assets"))


def _backdrop_path(edition):
    name = "student_background.png" if edition == "student" else "advisor_background.png"
    p = os.path.join(_ASSETS, name)
    return p if os.path.isfile(p) else None


class Backdrop:
    """背景画布，创建后须调用 attach(root) 并随窗口 resize 重绘。"""

    def __init__(self, edition="student"):
        self._t = get_theme(edition)
        self._edition = edition
        self._path = _backdrop_path(edition)
        self._img = None
        self._tk = None
        self._scaled = None          # 缓存按 (w,h) 缩放后的图，避免每次 resize 重算
        self._scaled_key = None
        self.canvas = None

    def attach(self, root):
        self.canvas = tk.Canvas(root, highlightthickness=0, bd=0, bg=self._t.bg)
        # place 铺满；用 Tk window lower 命令把背景沉到最底层
        self.canvas.place(x=0, y=0, relwidth=1, relheight=1)
        self.canvas.tk.call("lower", self.canvas._w)
        root.configure(bg=self._t.bg)
        self.canvas.bind("<Configure>", lambda e: self._redraw())
        self._redraw()

    def _redraw(self):
        if not self.canvas:
            return
        w = self.canvas.winfo_width()
        h = self.canvas.winfo_height()
        if w < 2 or h < 2:
            self.canvas.after(20, self._redraw)
            return
        self.canvas.delete("all")
        if not self._path or not os.path.isfile(self._path):
            self.canvas.create_rectangle(0, 0, w, h, fill=self._t.bg, outline=self._t.bg)
            return
        # 纯标准库实现：Tk 8.6 原生 PhotoImage 支持 PNG，无需任何第三方库。
        # cover 缩放：窗口比图大时用整数 zoom 放大；窗口比图小则直接用原图。
        # 居中靠"把图放在负偏移处、由画布裁剪"实现（tkinter 的 PhotoImage 没有 crop 方法），
        # 这样既 cover 又不依赖 PIL。
        try:
            if self._tk is None:
                self._tk = tk.PhotoImage(file=self._path)
            iw, ih = self._tk.width(), self._tk.height()
            key = (w, h)
            if self._scaled is not None and self._scaled_key == key:
                img = self._scaled
                biw, bih = img.width(), img.height()
            elif w > iw or h > ih:
                factor = max(1, math.ceil(max(w / iw, h / ih)))
                img = self._tk.zoom(factor, factor)
                biw, bih = img.width(), img.height()
                self._scaled, self._scaled_key = img, key
            else:
                img, biw, bih = self._tk, iw, ih
            left = (biw - w) // 2
            top = (bih - h) // 2
            # 图标放在负偏移，画布自动裁掉溢出部分 → 居中 cover
            self.canvas.create_image(-left, -top, anchor="nw", image=img)
        except Exception:
            self.canvas.create_rectangle(0, 0, w, h, fill=self._t.bg, outline=self._t.bg)
