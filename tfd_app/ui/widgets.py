# -*- coding: utf-8 -*-
"""
论文格式医生 · 中国版共享 UI 零件（严格照施工包 SVG 规格）
========================================================

所有零件都从 ``ui.theme`` 取色，禁止在零件里硬编码第二套样式。
可用零件（COMPONENT_MAP.md）：
  RoundButton / RoundCard / Badge / Stepper / InfoPanel / DropZone / Footer

按钮：主操作填 primary，次操作填 surface + primary 描边（照 primary-button /
secondary-button.svg）。圆角取 RADIUS.control(8)。
"""
from __future__ import annotations

import tkinter as tk
from . import theme
from .theme import RADIUS, SIZE, get_theme


def _rounded_rect(canvas, x1, y1, x2, y2, r, **kw):
    """在 canvas 上画圆角矩形（polygon 近似，tk 无原生 round-rect）。"""
    r = min(r, (x2 - x1) / 2, (y2 - y1) / 2)
    pts = [
        x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r,
        x2, y2 - r, x2, y2, x2 - r, y2, x1 + r, y2,
        x1, y2, x1, y2 - r, x1, y1 + r, x1, y1,
    ]
    canvas.create_polygon(pts, smooth=True, **kw)


class RoundButton(tk.Frame):
    """圆角按钮（主/次/禁用），API 兼容 ttk.Button 的 config(text/state/command)。"""

    def __init__(self, master, text="", edition="student", style="primary",
                 font=None, height=52, width=None, command=None, **kw):
        # width：显式像素宽（兼容 ttk.Button 的 width= 写法）；为 None 时按文字自适应。
        self._width_hint = width
        super().__init__(master, **kw)
        self._t = get_theme(edition)
        self._style = style
        self._text = text
        self._command = command
        self._enabled = True
        self._hovered = False
        self._font = font or self._t.sans(SIZE["body"], bold=True)
        self.cv = tk.Canvas(self, highlightthickness=0, bd=0, height=height)
        self.cv.pack(fill="both", expand=True)
        self.cv.bind("<Configure>", lambda e: self._draw())
        self._bind_events()
        self._apply_width()
        self._draw()

    # ---- 宽度（避免 tk 画布默认 ~378px 撑爆按钮行）----
    def _apply_width(self):
        if self._width_hint:
            self.cv.configure(width=self._width_hint)
            return
        if self._text:
            try:
                f = tk.font.Font(font=self._font)
                px = int(f.measure(self._text))
            except Exception:
                px = len(self._text) * 14
            self.cv.configure(width=max(80, px + 36))
        else:
            self.cv.configure(width=80)

    # ---- 事件 ----
    def _bind_events(self):
        for w in (self, self.cv):
            w.bind("<Enter>", self._on_enter)
            w.bind("<Leave>", self._on_leave)
            w.bind("<Button-1>", self._on_click)

    def _on_enter(self, _e=None):
        if self._enabled:
            self._hovered = True
            self._draw()

    def _on_leave(self, _e=None):
        self._hovered = False
        self._draw()

    def _on_click(self, _e=None):
        if self._enabled and self._command:
            self._command()

    # ---- 配色 ----
    def _colors(self):
        t = self._t
        if not self._enabled:
            return t.disabled_fill, t.disabled_border, t.disabled_text
        if self._style == "primary":
            fill = t.primary_hover if self._hovered else t.primary
            return fill, fill, "#FFFFFF"
        # secondary
        fill = t.surface
        border = t.primary if self._hovered else t.secondary_border
        return fill, border, t.primary

    def _draw(self):
        self.cv.delete("all")
        w = self.cv.winfo_width()
        h = self.cv.winfo_height()
        if w < 2 or h < 2:
            self.after(20, self._draw)
            return
        fill, border, fg = self._colors()
        _rounded_rect(self.cv, 1, 1, w - 1, h - 1, RADIUS["control"],
                      fill=fill, outline=border, width=2)
        self.cv.create_text(w / 2, h / 2, text=self._text, fill=fg,
                            font=self._font, anchor="center")

    # ---- 兼容 ttk.Button 的 config ----
    def config(self, **kw):
        if "text" in kw:
            self._text = kw.pop("text")
            self._apply_width()
        if "width" in kw:          # ttk 的显示宽度，显式指定则采用
            self._width_hint = kw.pop("width")
            self._apply_width()
        if "state" in kw:
            self._enabled = (kw.pop("state") != "disabled")
        if "command" in kw:
            self._command = kw.pop("command")
        if kw:
            super().config(**kw)
        self._draw()

    def set_text(self, text):
        self._text = text
        self._apply_width()
        self._draw()


class RoundCard(tk.Frame):
    """圆角卡片：canvas 画圆角底，内层 Frame 放内容。"""

    def __init__(self, master, edition="student", fill=None, border=None,
                 radius=None, padx=18, pady=16, **kw):
        super().__init__(master, **kw)
        self._t = get_theme(edition)
        self._fill = fill or self._t.surface
        self._border = border or self._t.border
        self._radius = radius or RADIUS["card"]
        self._padx = padx
        self._pady = pady
        self.cv = tk.Canvas(self, highlightthickness=0, bd=0)
        self.cv.pack(fill="both", expand=True)
        self.inner = tk.Frame(self.cv, bg=self._fill)
        self._win = self.cv.create_window(self._padx, self._pady,
                                          window=self.inner, anchor="nw")
        self.cv.bind("<Configure>", lambda e: self._draw())
        self.inner.bind("<Configure>", lambda e: self._resize_inner())
        self.after(20, self._draw)

    def _resize_inner(self):
        w = self.cv.winfo_width()
        self.cv.coords(self._win, self._padx, self._pady)
        self.cv.itemconfig(self._win, width=max(2, w - self._padx * 2))

    def _draw(self):
        self.cv.delete("all")
        w = self.cv.winfo_width()
        h = self.cv.winfo_height()
        if w < 2 or h < 2:
            self.after(20, self._draw)
            return
        _rounded_rect(self.cv, 1, 1, w - 1, h - 1, self._radius,
                      fill=self._fill, outline=self._border, width=2)
        self._win = self.cv.create_window(self._padx, self._pady,
                                          window=self.inner, anchor="nw")
        self._resize_inner()


class Badge(tk.Frame):
    """胶囊徽标：trial（蓝底）/ formal（绿灰底）。text 可变。"""

    def __init__(self, master, kind="trial", edition="student", text="", **kw):
        super().__init__(master, **kw)
        self._t = get_theme(edition)
        self._kind = kind
        self._text = text
        self.cv = tk.Canvas(self, highlightthickness=0, bd=0, height=34)
        self.cv.pack()                      # 不 fill：宽度由文字决定，避免 Canvas 默认 ~378px 撑成椭圆
        self.cv.bind("<Configure>", lambda e: self._draw())
        self._apply_width()
        self._draw()

    def _apply_width(self):
        """按文字测量宽度（含左右留白），避免 Canvas 默认 ~378px 撑成椭圆。"""
        try:
            f = tk.font.Font(font=self._t.sans(SIZE["caption"], bold=True))
            w = max(64, int(f.measure(self._text)) + 28)
        except Exception:
            w = max(64, len(self._text) * 14 + 28)
        self.cv.configure(width=w)

    def _colors(self):
        t = self._t
        if self._kind == "formal":
            return t.formal_fill, t.formal_text
        return t.primary_soft, t.primary

    def _draw(self):
        self.cv.delete("all")
        w = self.cv.winfo_width()
        if w < 2:
            self.after(20, self._draw)
            return
        fill, fg = self._colors()
        _rounded_rect(self.cv, 1, 1, w - 1, 32, RADIUS["pill"], fill=fill,
                      outline=fill, width=2)
        self.cv.create_text(w / 2, 17, text=self._text, fill=fg,
                            font=self._t.sans(SIZE["caption"], bold=True),
                            anchor="center")

    def set_text(self, text):
        self._text = text
        self._apply_width()
        self._draw()


class Stepper(tk.Frame):
    """水平三步条（准备 / 检查 / 修正）。暴露 circles / titles / lines 列表，
    兼容既有 ``_refresh_wizard`` 的 config 调用。"""

    def __init__(self, master, steps, edition="student", **kw):
        super().__init__(master, **kw)
        self._t = get_theme(edition)
        t = self._t
        self._step_circle = []
        self._step_title = []
        self._step_line = []
        n = len(steps)
        top = tk.Frame(self, bg=t.surface)
        top.pack(fill="x")
        for i, (label, desc) in enumerate(steps):
            col = tk.Frame(top, bg=t.surface)
            col.grid(row=0, column=2 * i)
            circ = tk.Label(col, text=str(i + 1), width=3, height=1,
                            relief="flat", highlightthickness=2,
                            highlightbackground=t.step_pending_border,
                            bg=t.step_pending_fill, fg=t.muted,
                            font=t.sans(13, bold=True))
            circ.pack()
            title = tk.Label(col, text=label, bg=t.surface, fg=t.muted,
                             font=t.sans(SIZE["caption"], bold=True))
            title.pack(pady=(4, 0))
            self._step_circle.append(circ)
            self._step_title.append(title)
            if i < n - 1:
                ln = tk.Frame(top, width=54, height=2, bg=t.stepper_line)
                ln.grid(row=0, column=2 * i + 1, sticky="")
                self._step_line.append(ln)
            else:
                self._step_line.append(None)

    @property
    def circles(self):
        return self._step_circle

    @property
    def titles(self):
        return self._step_title

    @property
    def lines(self):
        return self._step_line


class InfoPanel(tk.Frame):
    """「如何工作」信息面板（info-panel.svg）。"""

    def __init__(self, master, edition="student", title="如何工作",
                 line1="", line2="", **kw):
        super().__init__(master, **kw)
        self._t = get_theme(edition)
        t = self._t
        self.cv = tk.Canvas(self, highlightthickness=0, bd=0)
        self.cv.pack(fill="both", expand=True)
        self.inner = tk.Frame(self.cv, bg=t.info_fill)
        self._win = self.cv.create_window(0, 0, window=self.inner, anchor="nw")
        self.cv.bind("<Configure>", lambda e: self._draw())
        self.title_label = tk.Label(self.inner, text=title, bg=t.info_fill,
                                    fg=t.stepper_label_active,
                                    font=t.sans(SIZE["body"], bold=True))
        self.title_label.pack(anchor="w", padx=16, pady=(12, 4))
        self.line1_lbl = None
        self.line2_lbl = None
        if line1:
            self.line1_lbl = tk.Label(self.inner, text=line1, bg=t.info_fill, fg=t.muted,
                                      font=t.sans(SIZE["caption"]))
            self.line1_lbl.pack(anchor="w", padx=16)
        if line2:
            self.line2_lbl = tk.Label(self.inner, text=line2, bg=t.info_fill, fg=t.muted,
                                      font=t.sans(11))
            self.line2_lbl.pack(anchor="w", padx=16, pady=(4, 12))
        self.after(20, self._draw)

    def _draw(self):
        self.cv.delete("all")
        w = self.cv.winfo_width()
        h = self.cv.winfo_height()
        if w < 2 or h < 2:
            self.after(20, self._draw)
            return
        _rounded_rect(self.cv, 1, 1, w - 1, h - 1, RADIUS["card"],
                      fill=self._t.info_fill, outline=self._t.info_border, width=2)
        self._win = self.cv.create_window(0, 0, window=self.inner, anchor="nw")
        self.cv.coords(self._win, 0, 0)
        self.cv.itemconfig(self._win, width=w)


class DropZone(tk.Frame):
    """上传区（upload-dropzone.svg）：虚线圆角边框 + 箭头 + 标题 + 副标题。"""

    def __init__(self, master, edition="student", title="选择论文文件",
                 subtitle="", command=None, **kw):
        super().__init__(master, **kw)
        self._t = get_theme(edition)
        self._command = command
        self.cv = tk.Canvas(self, highlightthickness=0, bd=0, height=150)
        self.cv.pack(fill="both", expand=True)
        self.cv.bind("<Configure>", lambda e: self._draw())
        self.cv.bind("<Button-1>", lambda e: command and command())
        self._title = title
        self._subtitle = subtitle
        self.after(20, self._draw)

    def _draw(self):
        self.cv.delete("all")
        w = self.cv.winfo_width()
        h = self.cv.winfo_height()
        if w < 2 or h < 2:
            self.after(20, self._draw)
            return
        t = self._t
        _rounded_rect(self.cv, 2, 2, w - 2, h - 2, RADIUS["card"],
                      fill=t.dropzone_fill, outline=t.dropzone_border, width=2,
                      dash=(6, 4))
        # 上箭头
        cx = w / 2
        self.cv.create_line(cx, 40, cx, 74, fill=t.primary, width=4,
                            arrow=tk.LAST, arrowshape=(10, 12, 5))
        self.cv.create_text(cx, 92, text=self._title, fill=t.primary,
                            font=t.sans(16, bold=True), anchor="center")
        if self._subtitle:
            self.cv.create_text(cx, 118, text=self._subtitle, fill=t.muted,
                                font=t.sans(SIZE["caption"]), anchor="center")

    def set_subtitle(self, text):
        self._subtitle = text
        self._draw()


class SelectButton(tk.Frame):
    """选择框（select.svg）：圆角矩形 + 左对齐文字 + 右侧向下箭头，点击触发选择。"""

    def __init__(self, master, edition="student", placeholder="请选择",
                 command=None, height=48, **kw):
        super().__init__(master, **kw)
        self._t = get_theme(edition)
        self._command = command
        self._placeholder = placeholder
        self._text = placeholder
        self._height = height
        self.cv = tk.Canvas(self, highlightthickness=0, bd=0, height=height)
        self.cv.pack(fill="both", expand=True)
        self.cv.bind("<Configure>", lambda e: self._draw())
        self.cv.bind("<Button-1>", lambda e: command and command())
        self.after(20, self._draw)

    def _draw(self):
        self.cv.delete("all")
        w = self.cv.winfo_width()
        h = self.cv.winfo_height()
        if w < 2 or h < 2:
            self.after(20, self._draw)
            return
        t = self._t
        _rounded_rect(self.cv, 1, 1, w - 1, h - 1, RADIUS["control"],
                      fill=t.surface, outline=t.select_border, width=2)
        # 文字左对齐，垂直居中
        self.cv.create_text(18, h / 2, text=self._text, fill=t.select_ph,
                            font=t.sans(SIZE["body"]), anchor="w")
        # 向下箭头
        ax = w - 26
        ay = h / 2 - 4
        self.cv.create_line(ax, ay, ax + 7, ay + 8, ax + 14, ay,
                            fill=t.primary, width=2, smooth=False)

    def set_text(self, text):
        self._text = text if text else self._placeholder
        self._draw()


class Footer(tk.Frame):
    """页脚（footer.svg）：客服微信 / 官方公众号 / 邮箱 / 官网 / 本机处理声明。"""

    def __init__(self, master, edition="student",
                 wechat="芦苇不熬夜", official_id="reedskill",
                 mail="hi@reedskill.com", site="reedskill.com", **kw):
        super().__init__(master, **kw)
        self._t = get_theme(edition)
        t = self._t
        sep = tk.Frame(self, bg=t.border)
        sep.pack(fill="x")
        row = tk.Frame(self, bg=t.bg)
        row.pack(fill="x", pady=(8, 0))
        items = [("客服微信", wechat), ("官方公众号", official_id),
                 ("邮箱", mail), ("官网", site)]
        for i, (k, v) in enumerate(items):
            if i:
                tk.Label(row, text="·", bg=t.bg, fg=t.muted,
                         font=t.sans(SIZE["caption"])).pack(side="left", padx=10)
            tk.Label(row, text="%s：%s" % (k, v), bg=t.bg, fg=t.muted,
                     font=t.sans(SIZE["caption"])).pack(side="left")
        tk.Label(row, text="本机处理 · 文件不会上传任何服务器", bg=t.bg,
                 fg=t.muted, font=t.sans(11)).pack(side="right")
