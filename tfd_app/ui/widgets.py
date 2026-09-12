# -*- coding: utf-8 -*-
"""
论文格式医生 · 中国版共享 UI 零件（严格照施工包 SVG 规格）
========================================================

所有零件都从 ``ui.theme`` 取色，禁止在零件里硬编码第二套样式。
可用零件（COMPONENT_MAP.md）：
  RoundButton / RoundCard / Badge / Stepper / InfoPanel / DropZone / SelectButton
  / Radio / CheckRow / Toast / SlantNote / LinkRow / Footer

按钮：主操作填 primary，次操作填 surface + primary 描边（照 primary-button /
secondary-button.svg），危险操作填 danger（danger-button.svg），禁用态由 state 控制
（disabled-button.svg）。圆角取 RADIUS.control(8)。

文案铁律：零件自带文案一律逐字照 ``_ui_ref/COPY_TABLE.md``，调用方不得改写。
"""
from __future__ import annotations

import tkinter as tk
from . import backdrop
from . import theme
from .theme import RADIUS, TYPE, get_theme


# 按钮形态自带文案（逐字照 COPY_TABLE §5；调用方不得改写）。
#   danger-button.svg  → 危险按钮「删除 / 取消」
#   disabled-button.svg → 禁用按钮「暂不可用」
COPY_DANGER = "删除 / 取消"
COPY_DISABLED = "暂不可用"


def _rounded_rect(canvas, x1, y1, x2, y2, r, **kw):
    """在 canvas 上画圆角矩形（polygon 近似，tk 无原生 round-rect）。"""
    r = min(r, (x2 - x1) / 2, (y2 - y1) / 2)
    pts = [
        x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r,
        x2, y2 - r, x2, y2, x2 - r, y2, x1 + r, y2,
        x1, y2, x1, y2 - r, x1, y1 + r, x1, y1,
    ]
    canvas.create_polygon(pts, smooth=True, **kw)


def fit_font(font, text, max_w, min_size=TYPE["btn_small"]):
    """把字号往下调到 ``text`` 能在 ``max_w`` 内放得下（放得下就原样返回）。

    画布文字没有省略号，宽度不够时会被画布边缘直接裁掉半行；宁可小一号也不裁字。
    """
    try:
        f = tk.font.Font(font=font)
    except Exception:
        return font
    if max_w <= 0 or not text or f.measure(text) <= max_w:
        return font
    if isinstance(font, (tuple, list)):
        fam = font[0]
        size = abs(int(font[1]))
        weight = font[2] if len(font) > 2 else "normal"
        make = lambda s: (fam, s, weight)
    else:
        fam, size, weight = f.cget("family"), abs(int(f.cget("size"))), f.cget("weight")
        make = lambda s: tk.font.Font(family=fam, size=s, weight=weight)
    while size > min_size:
        size -= 1
        try:
            if tk.font.Font(font=make(size)).measure(text) <= max_w:
                return make(size)
        except Exception:
            return font
    return make(size)


def _wrap_canvas_text(font, text, max_w):
    """把画布文字按 ``max_w`` 折行（中文可逐字断行）；放得下就单行返回。

    画布文字项没有省略号，宽度不够会被边缘直接裁掉半句（英文单词也从不折行）。
    宁可折成两行，也不裁字（BRIEF §8.4「不得截断」）。返回 ``[text]``（单行）
    或按需切好的多行列表。
    """
    if not text:
        return []
    try:
        f = tk.font.Font(font=font)
    except Exception:
        return [text]
    if max_w <= 0 or f.measure(text) <= max_w:
        return [text]
    lines, cur = [], ""
    for ch in text:
        if cur and f.measure(cur + ch) > max_w:
            lines.append(cur)
            cur = ch
        else:
            cur += ch
    if cur:
        lines.append(cur)
    return lines


def wrap_height(lbl, wrap_w, extra=0):
    """估算标签在 ``wrap_w`` 宽度下折行后的高度。

    标签被 ``pack_forget`` 后 ``winfo_reqheight()`` 会失真（换行宽度被重置成 1，
    算出来是一条竖线），上层据此判断「放不下」就会把空格一直留着。这里按字体
    真实测量重新折算行数，隐藏状态也能算准。
    """
    try:
        f = tkfont.Font(font=lbl.cget("font"))
        text = lbl.cget("text") or ""
        w = int(max(1, wrap_w))
        lines = max(1, (f.measure(str(text)) + w - 1) // w)
        return lines * f.metrics("linespace") + int(extra)
    except Exception:
        return int(lbl.winfo_reqheight()) + int(extra)


class RoundButton(tk.Frame):
    """圆角按钮（主 / 次 / 危险 / 禁用），API 兼容 ttk.Button 的 config(text/state/command)。

    四种形态文案照装修包零件：主（primary-button.svg）、次（secondary-button.svg）、
    危险 ``删除 / 取消``（danger-button.svg）、禁用 ``暂不可用``（disabled-button.svg）。
    """

    def __init__(self, master, text="", edition="student", style="primary",
                 font=None, height=40, width=None, command=None, bg=None,
                 paper=False, **kw):
        # width：显式像素宽（兼容 ttk.Button 的 width= 写法）；为 None 时按文字自适应。
        # paper=True：按钮直接放在宣纸底图上（如顶栏）——画布改用宣纸“透明”容器，
        # 圆角之外透出底图纹理，不会在标题栏留下一块不透明的浅色方框。
        # 这只改“底色从哪来”，圆角/配色/尺寸与卡片内的按钮仍是同一套规范。
        self._width_hint = width
        super().__init__(master, **kw)
        self._t = get_theme(edition)
        self._style = style
        if not text and style == "danger":
            text = COPY_DANGER          # 危险按钮缺省文案（danger-button.svg）
        self._text = text
        self._command = command
        self._enabled = True
        self._hovered = False
        self._font = font or self._t.sans(TYPE["body"], bold=True)
        if paper:
            self.cv = backdrop.TextureCanvas(self, edition=edition, height=height)
        else:
            # 画布底色取卡片面：圆角之外的四角不会露出 Tk 默认灰底
            self.cv = tk.Canvas(self, highlightthickness=0, bd=0, height=height,
                                bg=bg or self._t.surface)
        self.cv.pack(fill="both", expand=True)
        self.cv.bind("<Configure>", lambda e: self._draw(), add="+")
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
        if self._style == "danger":
            # danger-button.svg：浅红底 + 朱砂边 + 朱砂字（只做「删除 / 取消」这类破坏性操作）
            return t.danger_fill, t.danger_border, t.danger_text
        # secondary
        fill = t.surface
        border = t.primary if self._hovered else t.secondary_border
        return fill, border, t.primary

    def _draw(self):
        self.cv.delete("ctl")               # 只重画控件本身，底图（tfd_tex）保留
        w = self.cv.winfo_width()
        h = self.cv.winfo_height()
        if w < 2 or h < 2:
            self.after(20, self._draw)
            return
        fill, border, fg = self._colors()
        _rounded_rect(self.cv, 1, 1, w - 1, h - 1, RADIUS["control"],
                      fill=fill, outline=border, width=2, tags="ctl")
        self.cv.create_text(w / 2, h / 2, text=self._text, fill=fg,
                            font=fit_font(self._font, self._text, w - 16),
                            anchor="center", tags="ctl")

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
            if not self._enabled and not self._text:
                self._text = COPY_DISABLED   # 禁用按钮缺省文案（disabled-button.svg）
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
        # 画布用宣纸“透明”容器：圆角之外的四角 + 1px 描边外沿透出底图纹理。
        # 旧实现没给画布 bg，Tk 默认灰底（#F0F0F0）会从四个圆角露出来，
        # 在宣纸背景上就是四块显眼的灰角。
        self.cv = backdrop.TextureCanvas(self, edition=edition)
        self.cv.pack(fill="both", expand=True)
        self.inner = tk.Frame(self.cv, bg=self._fill)
        self._win = None
        self.cv.bind("<Configure>", lambda e: self._draw(), add="+")
        self.inner.bind("<Configure>", lambda e: self._resize_inner())
        self.after(20, self._draw)

    def _resize_inner(self):
        if self._win is None:
            return
        w = self.cv.winfo_width()
        h = self.cv.winfo_height()
        self.cv.coords(self._win, self._padx, self._pady)
        # 高度也一并钉死：内层 Frame 被限制在卡片范围内，pack 才有“可用高度”可分，
        # 底部按钮行不会再被挤出卡片外沿（旧问题的根因之一）。
        self.cv.itemconfig(self._win, width=max(2, w - self._padx * 2),
                           height=max(2, h - self._pady * 2))

    def _draw(self):
        self.cv.delete("card")              # 只重画卡片本身，底图（tfd_tex）保留
        w = self.cv.winfo_width()
        h = self.cv.winfo_height()
        if w < 2 or h < 2:
            self.after(20, self._draw)
            return
        if self._win is not None:
            try:
                self.cv.delete(self._win)
            except Exception:
                pass
        _rounded_rect(self.cv, 1, 1, w - 1, h - 1, self._radius,
                      fill=self._fill, outline=self._border, width=2,
                      tags="card")
        self._win = self.cv.create_window(self._padx, self._pady,
                                          window=self.inner, anchor="nw")
        self._resize_inner()


class Badge(tk.Frame):
    """胶囊徽标：trial（蓝底）/ formal（绿灰底）/ soft（浅色小徽章）。text 可变。

    文案照装修包：试用徽章 ``试用版 · 剩余 X 次``（trial-badge.svg）、
    正式徽章 ``正式版 · 月卡``（formal-badge.svg）、浅色小徽章如 ``可选``。
    """

    def __init__(self, master, kind="trial", edition="student", text="", bg=None,
                 paper=False, **kw):
        super().__init__(master, **kw)
        self._t = get_theme(edition)
        self._kind = kind
        self._text = text
        if paper:
            # 与 RoundButton(paper=True) 同理：胶囊直接铺在宣纸底图上，圆角外不留方框
            self.cv = backdrop.TextureCanvas(self, edition=edition, height=30)
        else:
            self.cv = tk.Canvas(self, highlightthickness=0, bd=0, height=30,
                                bg=bg or self._t.surface)
        self.cv.pack()                      # 不 fill：宽度由文字决定，避免 Canvas 默认 ~378px 撑成椭圆
        self.cv.bind("<Configure>", lambda e: self._draw(), add="+")
        self._apply_width()
        self._draw()

    def _apply_width(self):
        """按文字测量宽度（含左右留白），避免 Canvas 默认 ~378px 撑成椭圆。"""
        try:
            f = tk.font.Font(font=self._t.sans(TYPE["caption"], bold=True))
            w = max(64, int(f.measure(self._text)) + 28)
        except Exception:
            w = max(64, len(self._text) * 14 + 28)
        self.cv.configure(width=w)

    def _colors(self):
        t = self._t
        if self._kind == "formal":
            return t.formal_fill, t.formal_text
        if self._kind == "soft":
            # 视觉稿「可选」这类浅色小徽章：安静的中性标记，不是第二种状态样式
            return t.soft_fill, t.soft_text
        return t.primary_soft, t.primary

    def _draw(self):
        self.cv.delete("ctl")
        w = self.cv.winfo_width()
        if w < 2:
            self.after(20, self._draw)
            return
        fill, fg = self._colors()
        _rounded_rect(self.cv, 1, 1, w - 1, 32, RADIUS["pill"], fill=fill,
                      outline=fill, width=2, tags="ctl")
        self.cv.create_text(w / 2, 17, text=self._text, fill=fg,
                            font=self._t.sans(TYPE["caption"], bold=True),
                            anchor="center", tags="ctl")

    def set_text(self, text):
        self._text = text
        self._apply_width()
        self._draw()

    def set_kind(self, kind):
        """切换徽标类型（trial 蓝 / formal 绿灰）；文案由 set_text 另更。"""
        if kind != self._kind:
            self._kind = kind
            self._draw()


class StepCircle(tk.Canvas):
    """步进器圆点（stepper.svg：直径 44 圆形，编号居中）。

    兼容 tk.Label 的 ``config(bg / fg / highlightbackground / text)`` 写法——
    ``_refresh_wizard`` 就是用这套写法驱动三种状态，行为保持不变。
    """

    DIAMETER = 36

    def __init__(self, master, edition="student", text="1", **kw):
        t = get_theme(edition)
        super().__init__(master, width=self.DIAMETER, height=self.DIAMETER,
                         highlightthickness=0, bd=0, bg=t.surface, **kw)
        self._t = t
        self._fill = t.step_pending_fill
        self._outline = t.step_pending_border
        self._fg = t.muted
        self._text = text
        self._draw()

    def _draw(self):
        self.delete("all")
        s = self.DIAMETER
        self.create_oval(2, 2, s - 2, s - 2, fill=self._fill,
                         outline=self._outline, width=2)
        self.create_text(s / 2, s / 2, text=self._text, fill=self._fg,
                         font=self._t.sans(TYPE["step_num"], bold=True), anchor="center")

    def config(self, **kw):
        if "bg" in kw:
            self._fill = kw.pop("bg")
        if "highlightbackground" in kw:
            self._outline = kw.pop("highlightbackground")
        if "fg" in kw:
            self._fg = kw.pop("fg")
        if "text" in kw:
            self._text = kw.pop("text")
        if kw:
            super().config(**kw)
        self._draw()

    configure = config


# stepper.svg 的三个状态标签（数字 1 / 2 / 3 下方那行小字，逐字照 COPY_TABLE）
STEP_STATE_LABELS = ("准备", "检查", "修正")


class Stepper(tk.Frame):
    """步骤条（视觉稿：竖排——圆点＋步骤名＋说明；圆点/连线/配色照 stepper.svg）。

    暴露 circles / titles / lines / descs 列表，兼容既有 ``_refresh_wizard`` 的
    ``config(bg / fg / highlightbackground / text)`` 调用。
    """

    def __init__(self, master, steps, edition="student", state_labels=STEP_STATE_LABELS, **kw):
        super().__init__(master, **kw)
        self._t = get_theme(edition)
        t = self._t
        # 装修包 stepper.svg 的三态标签（准备 / 检查 / 修正），供上层按状态取用
        self._state_labels = tuple(state_labels)
        self._step_circle = []
        self._step_title = []
        self._step_line = []
        self._step_desc = []
        self._compact = False
        self._rows = []
        n = len(steps)
        d = StepCircle.DIAMETER
        for i, (label, desc) in enumerate(steps):
            row = tk.Frame(self, bg=t.surface)
            self._rows.append(row)
            row.pack(fill="x")
            # 左列：圆点 + 向下的连接线（连线对齐圆点圆心）
            rail = tk.Frame(row, bg=t.surface, width=d)
            rail.pack(side="left", fill="y")
            rail.pack_propagate(False)
            circ = StepCircle(rail, edition=edition, text=str(i + 1))
            circ.pack(pady=(4, 0))
            ln = None
            if i < n - 1:
                ln = tk.Frame(rail, width=3, bg=t.stepper_line)
                ln.pack(fill="y", expand=True, pady=(3, 0))
            # 右列：步骤名 + 说明（说明文案照 step_desc，仅渲染，不改逻辑）
            txt = tk.Frame(row, bg=t.surface)
            txt.pack(side="left", fill="x", expand=True, padx=(12, 0), pady=(0, 4))
            title = tk.Label(txt, text=label, bg=t.surface, fg=t.muted,
                             font=t.sans(TYPE["body"], bold=True),
                             anchor="w", justify="left")
            title.pack(fill="x")
            dl = tk.Label(txt, text=desc, bg=t.surface, fg=t.muted,
                          font=t.sans(TYPE["caption"]), anchor="w", justify="left")
            dl.pack(fill="x", pady=(2, 0))
            self._step_circle.append(circ)
            self._step_title.append(title)
            self._step_line.append(ln)
            self._step_desc.append(dl)

    @property
    def circles(self):
        return self._step_circle

    @property
    def titles(self):
        return self._step_title

    @property
    def lines(self):
        return self._step_line

    @property
    def descs(self):
        return self._step_desc

    @property
    def state_labels(self):
        """stepper.svg 三态标签：``准备`` / ``检查`` / ``修正``。"""
        return self._state_labels

    def state_label(self, index):
        """第 ``index`` 步（0 基）的状态标签；越界返回空串。"""
        try:
            return self._state_labels[index]
        except Exception:
            return ""

    def set_compact(self, compact):
        """紧凑模式：收起步骤说明文字（窗口太矮时用，保证步骤名与按钮都完整）。"""
        compact = bool(compact)
        if compact == self._compact:
            return
        self._compact = compact
        for dl in self._step_desc:
            if compact:
                if dl.winfo_manager():
                    dl.pack_forget()
            elif not dl.winfo_manager():
                dl.pack(fill="x", pady=(2, 0))

    def height_with(self, compact):
        """预估步骤条在指定模式下所需高度（供上层在窄高窗口里取舍）。

        用每一行 Frame 的**真实** reqheight 结算，不再用「标题行高 + 固定值」估算——
        旧估算比实际矮 ~13px/行，上层据此判断「放得下」，结果按钮行被挤出卡片底部。
        """
        total = 0
        for i, row in enumerate(self._rows):
            h = row.winfo_reqheight()
            desc = self._step_desc[i]
            showed = bool(desc.winfo_manager())
            if compact and showed:
                h -= desc.winfo_reqheight() + 2
            elif not compact and not showed:
                h += desc.winfo_reqheight() + 2
            total += h
        return total


class InfoPanel(tk.Frame):
    """「如何工作」信息面板（info-panel.svg）。"""

    def __init__(self, master, edition="student", title="如何工作",
                 line1="上传论文 → 检查格式 → 查看问题 → 一键修正",
                 line2="文件默认在本机处理，不会上传到服务器。",
                 height=74, **kw):
        super().__init__(master, **kw)
        self._t = get_theme(edition)
        t = self._t
        self._min_h = height
        # 显式高度：Canvas 没有内容时默认请求 382x268，旧版因此留下一个空荡荡的大框；
        # 底色取面板填充色，圆角之外的四角不会露出 Tk 默认灰底。
        self.cv = tk.Canvas(self, highlightthickness=0, bd=0, height=height,
                            bg=t.info_fill)
        self.cv.pack(fill="both", expand=True)
        self.inner = tk.Frame(self.cv, bg=t.info_fill)
        self._win = self.cv.create_window(0, 0, window=self.inner, anchor="nw")
        self.cv.bind("<Configure>", lambda e: self._draw())
        self.title_label = tk.Label(self.inner, text=title, bg=t.info_fill,
                                    fg=t.stepper_label_active,
                                    font=t.sans(TYPE["body"], bold=True))
        self.title_label.pack(anchor="w", padx=16, pady=(6, 2))
        self.line1_lbl = None
        self.line2_lbl = None
        if line1:
            self.line1_lbl = tk.Label(self.inner, text=line1, bg=t.info_fill, fg=t.muted,
                                      font=t.sans(TYPE["caption"]))
            self.line1_lbl.pack(anchor="w", padx=16)
        if line2:
            self.line2_lbl = tk.Label(self.inner, text=line2, bg=t.info_fill, fg=t.muted,
                                      font=t.sans(TYPE["caption"]))
            self.line2_lbl.pack(anchor="w", padx=16, pady=(2, 10))
        self.inner.bind("<Configure>", lambda e: self._wrap_labels(), add="+")
        self.after(20, self._draw)

    def _wrap_labels(self):
        """说明文字按面板实际宽度折行：窄卡片里不会被右沿裁掉半句。"""
        try:
            w = self.inner.winfo_width()
        except Exception:
            return
        if w < 40:
            return
        for lbl in (self.line1_lbl, self.line2_lbl):
            if lbl is None:
                continue
            try:
                target = max(80, w - 32)
                if int(str(lbl.cget("wraplength"))) != target:
                    lbl.configure(wraplength=target)
            except Exception:
                pass

    def _fit_height(self):
        """按内容实际高度伸缩面板：字体/DPI 变大时，最后一行不会被下沿裁掉。"""
        try:
            need = max(self._min_h, self.inner.winfo_reqheight())
            cur = int(self.cv.cget("height"))
        except Exception:
            return
        if need != cur:
            try:
                self.cv.configure(height=need)
            except Exception:
                pass

    def _draw(self):
        self._fit_height()
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
    """上传区（视觉稿左卡）：虚线圆角边框 + 上箭头 + 主文案 + 副文案 + 次按钮。

    文案逐字照 COPY_TABLE：主 ``点击上传或拖拽文件到此处``、副
    ``支持 Word 文档（.docx、.docx）、WPS 格式（.wps）``、按钮 ``选择文件``。
    整块可点，按钮也走同一个 command（沿用既有 ``_pick_input``）。
    """

    def __init__(self, master, edition="student",
                 title="点击上传或拖拽文件到此处",
                 subtitle="支持 Word 文档（.docx、.docx）、WPS 格式（.wps）",
                 button_text="选择文件", command=None, height=150, **kw):
        super().__init__(master, **kw)
        self._t = get_theme(edition)
        self._command = command
        self.cv = tk.Canvas(self, highlightthickness=0, bd=0, height=height,
                            bg=self._t.dropzone_fill)
        self.cv.pack(fill="both", expand=True)
        self.cv.bind("<Configure>", lambda e: self._draw())
        self.cv.bind("<Button-1>", lambda e: command and command())
        self.min_height = height          # 上传区最低高度（上层据此取舍说明面板）
        self._title = title
        self._subtitle = subtitle
        self._button_text = button_text
        self._picked = ""          # 已选文件名（由 set_subtitle 写入；副文案永不被顶掉）
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
        # 内容（上传箭头 + 主文案 + 副文案/已选文件 + 选择文件按钮）按**真实行高**整体居中。
        # 副文案是长句（「支持 Word 文档（.docx、.docx）、WPS 格式（.wps）」）：窄卡片里
        # 单行放不下，按画布宽度折行——绝不让画布边缘把字裁掉半句（BRIEF §8.4）。
        # 已选文件与副文案**共用同一行槽位**：块高恒定，最小窗口下也不会把按钮挤出画布。
        cx = w / 2
        title_font = fit_font(t.sans(TYPE["card_title"], bold=True), self._title, w - 40)
        sub_font = t.sans(TYPE["caption"])
        sub_text = self._picked or self._subtitle
        sub_lines = _wrap_canvas_text(sub_font, sub_text, w - 28)
        try:
            lh_title = tk.font.Font(font=title_font).metrics("linespace")
            lh_sub = tk.font.Font(font=sub_font).metrics("linespace")
        except Exception:
            lh_title, lh_sub = 24, 18
        arrow_h, gap, btn_h = 24, 4, 26
        block = (arrow_h + gap + lh_title + 4 + len(sub_lines) * lh_sub
                 + 8 + btn_h)
        # 画布请求高度不得小于内容实高：否则最小窗口里底部按钮会被画布下沿裁掉半截。
        need_h = int(block) + 10
        try:
            if int(self.cv.cget("height")) != need_h:
                self.cv.configure(height=need_h)
        except Exception:
            pass
        y = max(6.0, (h - block) / 2.0)
        self.cv.create_line(cx, y + 2, cx, y + arrow_h, fill=t.primary, width=4,
                            arrow=tk.LAST, arrowshape=(10, 12, 5))
        y += arrow_h + gap
        self.cv.create_text(cx, y + lh_title / 2.0, text=self._title, fill=t.primary,
                            font=title_font, anchor="center")
        y += lh_title + 4
        for line in sub_lines:
            self.cv.create_text(cx, y + lh_sub / 2.0, text=line,
                                fill=t.primary if self._picked else t.muted,
                                font=t.sans(TYPE["caption"], bold=True) if self._picked
                                else sub_font, anchor="center")
            y += lh_sub
        y += 8
        # 次按钮「选择文件」（secondary-button.svg）：与整块点击走同一个 command
        if self._button_text:
            bf = t.sans(TYPE["btn_small"], bold=True)
            bw = int(tk.font.Font(font=bf).measure(self._button_text)) + 44
            x1, y1 = cx - bw / 2.0, y
            x2, y2 = cx + bw / 2.0, y1 + btn_h
            _rounded_rect(self.cv, x1, y1, x2, y2, RADIUS["control"],
                          fill=t.surface, outline=t.secondary_border, width=2)
            self.cv.create_text(cx, (y1 + y2) / 2.0, text=self._button_text,
                                fill=t.primary, font=bf, anchor="center")

    def set_subtitle(self, text):
        """显示「已选文件名」（占副文案那一行）；传空串即回到装修包副文案。"""
        self._picked = text or ""
        self._draw()


class SelectButton(tk.Frame):
    """选择框（select.svg）：圆角矩形 + 左对齐文字 + 右侧向下箭头，点击触发选择。"""

    def __init__(self, master, edition="student", placeholder="请选择模板",
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
                            font=fit_font(t.sans(TYPE["body"]), self._text, w - 46),
                            anchor="w")
        # 向下箭头
        ax = w - 26
        ay = h / 2 - 4
        self.cv.create_line(ax, ay, ax + 7, ay + 8, ax + 14, ay,
                            fill=t.primary, width=2, smooth=False)

    def set_text(self, text):
        self._text = text if text else self._placeholder
        self._draw()


class TickRow(tk.Frame):
    """标记行（radio-on/off.svg、checkbox-on/off.svg）：圆形或方形标记 + 文字。

    * ``mode="radio"``：文字由调用方给（主界面左卡用它显示「未选择论文 / 模板未选（可选）」）
    * ``mode="checkbox"``：默认文案照装修包 —— 未选 ``选择此项``、已选 ``已选择``

    两种形态同源（同一套描边/主色），只是标记形状不同，不算第二套样式。
    """

    MARK = 20

    def __init__(self, master, edition="student", mode="radio", text=None,
                 off_text="选择此项", on_text="已选择", selected=False,
                 empty_text="", bg=None, command=None, **kw):
        super().__init__(master, **kw)
        self._t = get_theme(edition)
        self._mode = "checkbox" if mode == "checkbox" else "radio"
        self._off_text = off_text
        self._on_text = on_text
        self._empty_text = empty_text
        self._selected = bool(selected)
        self._bg = bg or self._t.surface
        self._command = command
        if text is None:
            text = self._on_text if self._selected else self._off_text
        self._text = (text or "").strip() or self._empty_text
        self.cv = tk.Canvas(self, width=self.MARK, height=self.MARK,
                            highlightthickness=0, bd=0, bg=self._bg)
        self.cv.pack(side="left")
        self.lbl = tk.Label(self, text=self._text, bg=self._bg,
                            fg=self._t.stepper_label_active,
                            font=self._t.sans(TYPE["caption"]))
        self.lbl.pack(side="left", padx=(8, 0))
        if command:
            for wdg in (self, self.cv, self.lbl):
                wdg.bind("<Button-1>", lambda e: self._command())
        self._draw()

    def _draw(self):
        t = self._t
        c = self.cv
        c.delete("all")
        s = self.MARK
        border = t.radio_on if self._selected else t.radio_off_border
        if self._mode == "checkbox":
            _rounded_rect(c, 2, 2, s - 2, s - 2, 4,
                          fill=t.radio_on if self._selected else t.surface,
                          outline=border, width=2)
            if self._selected:
                c.create_line(6, s / 2.0, s * 0.40, s - 7, s - 5, 6,
                              fill="#FFFFFF", width=2)
        else:
            c.create_oval(2, 2, s - 2, s - 2, fill=t.surface,
                          outline=border, width=2)
            if self._selected:
                c.create_oval(6, 6, s - 6, s - 6, fill=t.radio_on,
                              outline=t.radio_on)

    def set_text(self, text):
        """设置文字；空串回落到 ``empty_text``（未选态文案），选中态另由 set_selected 管。"""
        text = (text or "").strip() or self._empty_text
        self._text = text
        self.lbl.config(text=text)

    def set_selected(self, on):
        self._selected = bool(on)
        self._draw()

    def config(self, **kw):
        """兼容 tk.Label 风格调用：text= / selected= / fg=。"""
        if "text" in kw:
            self.set_text(kw.pop("text"))
        if "selected" in kw:
            self.set_selected(kw.pop("selected"))
        if "fg" in kw:
            self.lbl.config(fg=kw.pop("fg"))
        if kw:
            super().config(**kw)

    configure = config


class CheckRow(TickRow):
    """复选框（checkbox-off.svg 未选「选择此项」 / checkbox-on.svg 已选「已选择」）。"""

    def __init__(self, master, edition="student", text=None, selected=False, **kw):
        super().__init__(master, edition=edition, mode="checkbox", text=text,
                         selected=selected, **kw)


class Radio(TickRow):
    """单选行（radio-on.svg / radio-off.svg）：主界面左卡显示当前选择状态。"""

    def __init__(self, master, edition="student", text="", selected=False, **kw):
        super().__init__(master, edition=edition, mode="radio", text=text,
                         selected=selected, **kw)


class Toast(tk.Frame):
    """提示条（toast.svg）：圆角浅底 + 圆点 + 文字。默认文案逐字照装修包。"""

    def __init__(self, master, edition="student",
                 text="已保存 · 你的设置已更新", height=44, **kw):
        super().__init__(master, **kw)
        self._t = get_theme(edition)
        self._text = text
        self.cv = tk.Canvas(self, highlightthickness=0, bd=0, height=height,
                            bg=self._t.surface)
        self.cv.pack(fill="both", expand=True)
        self.cv.bind("<Configure>", lambda e: self._draw())
        self._job = None
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
                      fill=t.modal_fill, outline=t.modal_border, width=2)
        self.cv.create_text(16, h / 2.0, text="●", fill=t.success_dot,
                            font=t.sans(TYPE["caption"]), anchor="w")
        self.cv.create_text(32, h / 2.0, text=self._text, fill=t.stepper_label_active,
                            font=t.sans(TYPE["caption"]), anchor="w")

    def show(self, text=None, ms=2600):
        """显示提示条，``ms`` 毫秒后自动收起。"""
        if text:
            self._text = text
            self._draw()
        try:
            self.pack(fill="x")
        except Exception:
            pass
        if self._job is not None:
            try:
                self.after_cancel(self._job)
            except Exception:
                pass
        self._job = self.after(int(ms), self.hide)

    def hide(self):
        self._job = None
        try:
            self.pack_forget()
        except Exception:
            pass


class SlantNote(tk.Canvas):
    """右上斜排小字（视觉稿：``三步完成`` / ``论文格式检查与修正``）。

    画在**卡片面**上（``bg`` 默认取 ``theme.surface``，与卡片内层同色）：
    卡片内部本来就不透明，这里若用宣纸底图画布，会在卡片上留出一块底色
    不同的方块（实测就是标题右侧那块「白补丁」）。字整体轻微倾斜；
    Tk 8.6 不支持 ``angle`` 时自动退回水平排布（内容一字不差，只是不斜）。

    用色取 ``stepper_label_active``（同「当前步骤」的墨蓝）：视觉稿里这两行是
    深色批注，压在卡片面上仍要看得清；高度留足旋转后的包围盒，
    否则倾斜会把首行的上缘裁掉。
    """

    def __init__(self, master, lines, edition="student", angle=-7,
                 height=None, bg=None, **kw):
        t = get_theme(edition)
        self._t = t
        self._lines = list(lines)
        self._angle = angle
        self._font = t.sans(TYPE["caption"])
        try:
            f = tk.font.Font(font=self._font)
            width = max(f.measure(s) for s in self._lines) + 34
            self._lh = int(f.metrics("linespace"))
        except Exception:
            width = max(len(s) for s in self._lines) * 15 + 34
            self._lh = 20
        kw.setdefault("width", int(width))
        # 旋转包围盒：Tk 8.6 的画布文字一旦超出画布下沿，渲染就会糊成一片
        # （实测第二行「论文格式检查与修正」被裁 → 字形互相叠压、还多出一条横线）。
        # 故高度按「两行行高 + 斜排下垂量 + 上下留白」算足，绝不让斜排文字越界。
        self._pad_top = 4
        self._slack = 14
        need_h = self._pad_top + self._lh * len(self._lines) + self._slack
        kw.setdefault("height", max(int(height or 0), need_h))
        kw.setdefault("highlightthickness", 0)
        kw.setdefault("bd", 0)
        kw.setdefault("bg", bg or t.surface)
        super().__init__(master, **kw)
        self.bind("<Configure>", lambda _e: self.draw_foreground(), add="+")
        self.draw_foreground()

    def draw_foreground(self):
        self.delete("slant")
        w = self.winfo_width()
        h = self.winfo_height()
        if w < 2 or h < 2:
            return
        t = self._t
        n = max(1, len(self._lines))
        for i, text in enumerate(self._lines):
            x = 10
            # 行位从顶部排：斜排会让右端下垂，行距用真实行高、不用 h/n 均分，
            # 否则画布一变高，两行之间的距离就被拉散（视觉稿是紧凑的批注小字）。
            y = self._pad_top + (i + 0.5) * self._lh
            try:
                self.create_text(x, y, text=text, fill=t.stepper_label_active,
                                 font=self._font, anchor="w", angle=self._angle,
                                 tags="slant")
            except Exception:
                self.create_text(x, y, text=text, fill=t.stepper_label_active,
                                 font=self._font, anchor="w", tags="slant")


class CanvasText:
    """画布文本项：只提供 tk.Label 的 ``config(text= / fg= / font=)`` 最小接口。

    给状态栏这类需要“无底色文字”的地方用：画布文字没有底色块，宣纸底图从字周围透出来。
    """

    def __init__(self, canvas, text="", x=0, y=0, anchor="w", fill="#000000",
                 font=None, right=False):
        self._cv = canvas
        self._x = x
        self._y = y
        self._anchor = anchor
        self._right = right
        self._text = text
        self._fill = fill
        self._font = font
        self._item = None
        canvas.bind("<Configure>", lambda e: self.redraw(), add="+")
        self.redraw()

    def redraw(self):
        cv = self._cv
        try:
            if not cv.winfo_exists() or cv.winfo_width() < 2:
                return
        except Exception:
            return
        x = (cv.winfo_width() - self._x) if self._right else self._x
        if self._item is None:
            self._item = cv.create_text(x, self._y, text=self._text,
                                        fill=self._fill, font=self._font,
                                        anchor=self._anchor, tags="fg")
        else:
            cv.coords(self._item, x, self._y)
            cv.itemconfigure(self._item, text=self._text, fill=self._fill,
                             font=self._font)
        cv.tag_raise(self._item)

    def config(self, **kw):
        if "text" in kw:
            self._text = kw.pop("text")
        if "fg" in kw:
            self._fill = kw.pop("fg")
        if "font" in kw:
            self._font = kw.pop("font")
        self.redraw()

    configure = config


class LinkRow(backdrop.TextureCanvas):
    """一行文字入口（视觉稿右上：客服微信 ｜ 官方公众号 ｜ 关于 ｜ 帮助）。

    文字是画布文本项：没有底色块，宣纸底图从字间透出来；宽度按实际字体测量，
    DPI 放大时也不会与旁边的按钮/胶囊挤在一起。
    """

    def __init__(self, master, items, edition="student", gap=16, sep="｜",
                 height=26, **kw):
        t = get_theme(edition)
        self._t = t
        self._items = list(items)              # [(文字, 点击回调), ...]
        self._gap = gap
        self._sep = sep
        self._font = t.sans(TYPE["caption"])
        self._fg = t.muted
        self._boxes = []
        try:
            f = tk.font.Font(font=self._font)
            widths = [f.measure(txt) for txt, _ in self._items]
            sep_w = f.measure(sep)
        except Exception:
            widths = [len(txt) * 14 for txt, _ in self._items]
            sep_w = 14
        n = max(0, len(self._items) - 1)
        total = sum(widths) + n * (sep_w + gap * 2)
        kw.setdefault("width", int(total) + 6)
        kw.setdefault("height", height)
        super().__init__(master, edition=edition, **kw)
        self._widths = widths
        self._sep_w = sep_w
        self.bind("<Button-1>", self._on_click)
        self.bind("<Motion>", self._on_motion)
        self.bind("<Leave>", lambda e: self.configure(cursor=""))
        self.draw_foreground()

    def draw_foreground(self):
        self.delete("link")
        w = self.winfo_width()
        h = self.winfo_height()
        if w < 2 or h < 2:
            return
        x = 0.0
        y = h / 2.0
        self._boxes = []
        for i, (txt, cmd) in enumerate(self._items):
            if i:
                self.create_text(x + self._gap + self._sep_w / 2.0, y, text=self._sep,
                                 fill=self._t.border, font=self._font,
                                 anchor="center", tags="link")
                x += self._gap * 2 + self._sep_w
            self.create_text(x, y, text=txt, fill=self._fg, font=self._font,
                             anchor="w", tags="link")
            self._boxes.append((x - 3, x + self._widths[i] + 3, cmd))
            x += self._widths[i]

    def _hit(self, x):
        for x0, x1, cmd in self._boxes:
            if x0 <= x <= x1:
                return cmd
        return None

    def _on_motion(self, event):
        self.configure(cursor="hand2" if self._hit(event.x) else "")

    def _on_click(self, event):
        cmd = self._hit(event.x)
        if cmd:
            cmd()


class Footer(backdrop.TextureCanvas):
    """页脚（视觉稿 + footer.svg）：细线 + 「邮箱 ｜ 标语 ｜ 官网」三件套 + 必留条目。

    第 1 行照视觉稿：``✉ hi@reedskill.com`` ／ ``专注学术 · 让格式更专业`` ／ ``🌐 reedskill.com``；
    第 2 行保留规格文档要求一条不落的条目：``客服微信``、``官方公众号``、
    ``本机处理 · 文件不会上传任何服务器``、``© 2026 论文格式医生``。
    文字都是画布项（没有底色块），宣纸底图从字间透出来。
    """

    MAIL_TEXT = "✉ hi@reedskill.com"
    SLOGAN_TEXT = "专注学术 · 让格式更专业"
    SITE_TEXT = "🌐 reedskill.com"
    PRIVACY_TEXT = "本机处理 · 文件不会上传任何服务器"
    COPYRIGHT_TEXT = "© 2026 论文格式医生"

    def __init__(self, master, edition="student", on_wechat=None, on_official=None, **kw):
        kw.setdefault("height", 70)
        super().__init__(master, edition=edition, **kw)
        t = get_theme(edition)
        self._t = t
        self._on_wechat = on_wechat
        self._on_official = on_official
        self._font = t.sans(TYPE["footer"])
        self._font_small = t.sans(TYPE["footer"])
        self._boxes = []
        self.bind("<Button-1>", self._on_click)
        self.bind("<Motion>", self._on_motion)

    def draw_foreground(self):
        self.delete("footer")
        w = self.winfo_width()
        h = self.winfo_height()
        if w < 2 or h < 2:
            return
        t = self._t
        f = self._font
        line_y = h * 0.34
        soft_y = h * 0.74
        self.create_line(0, 1, w, 1, fill=t.footer_line, tags="footer")
        # ---- 第 1 行：邮箱 / 标语（两侧细装饰线）/ 官网 ----
        self.create_text(0, line_y, text=self.MAIL_TEXT, fill=t.footer_text,
                         font=f, anchor="w", tags="footer")
        self.create_text(w, line_y, text=self.SITE_TEXT, fill=t.footer_text,
                         font=f, anchor="e", tags="footer")
        self.create_text(w / 2.0, line_y, text=self.SLOGAN_TEXT, fill=t.primary,
                         font=f, anchor="center", tags="footer")
        try:
            gap = tk.font.Font(font=f).measure(self.SLOGAN_TEXT) / 2.0 + 16
        except Exception:
            gap = 110
        for x0, x1 in ((w / 2.0 - gap - 74, w / 2.0 - gap),
                       (w / 2.0 + gap, w / 2.0 + gap + 74)):
            self.create_line(x0, line_y, x1, line_y, fill=t.footer_line,
                             tags="footer")
        # ---- 第 2 行：客服入口（可点）+ 必留声明 ----
        self._boxes = []
        x = 0.0
        for text, cmd in (("客服微信", self._on_wechat),
                          ("官方公众号", self._on_official)):
            try:
                item_w = tk.font.Font(font=self._font_small).measure(text)
            except Exception:
                item_w = len(text) * 12
            if x:
                self.create_text(x, soft_y, text="·", fill=t.footer_line,
                                 font=self._font_small, anchor="w", tags="footer")
                x += 14
            self.create_text(x, soft_y, text=text,
                             fill=t.primary if cmd else t.footer_text_soft,
                             font=self._font_small, anchor="w", tags="footer")
            self._boxes.append((x - 2, x + item_w + 2, cmd))
            x += item_w
        self.create_text(w, soft_y,
                         text=self.PRIVACY_TEXT + " · " + self.COPYRIGHT_TEXT,
                         fill=t.footer_text_soft, font=self._font_small,
                         anchor="e", tags="footer")

    def _hit(self, x):
        for x0, x1, cmd in self._boxes:
            if cmd and x0 <= x <= x1:
                return cmd
        return None

    def _on_motion(self, event):
        try:
            self.configure(cursor="hand2" if self._hit(event.x) else "")
        except Exception:
            pass

    def _on_click(self, event):
        cmd = self._hit(event.x)
        if cmd:
            cmd()
