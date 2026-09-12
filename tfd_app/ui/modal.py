# -*- coding: utf-8 -*-
"""
论文格式医生 · 中国版弹窗（照 upgrade-modal.svg / confirm-modal.svg）
=============================================================

ModalShell  ：弹窗壳（14 圆角、象牙面、细边、单一主操作 + 安静取消）
UpgradeModal：套餐升级（次/周/月/永久，月卡「推荐」高亮）+ 已有激活码入口
ConfirmModal：通用确认（取消 / 确认）

全部从 ui.theme 取色，禁止第二套样式。
"""
from __future__ import annotations

import tkinter as tk
from . import theme
from .theme import get_theme, RADIUS, SIZE
from .widgets import RoundButton, _rounded_rect


class ModalShell(tk.Toplevel):
    """带圆角画布的弹窗基底。subclass 在 self.body 里放内容。"""

    def __init__(self, parent, edition="student", title="", subtitle="",
                 width=620, height=300, **kw):
        super().__init__(parent, **kw)
        self._t = get_theme(edition)
        self._width = width
        self._height = height
        # 父窗口可见时才设 transient：启动时 root 被 withdraw，子窗若 transient 到
        # 已隐藏的父窗会被 Tk 级联隐藏/不出现，导致启动激活窗卡死。模态抓取仍保留。
        if parent is not None and parent.winfo_viewable():
            self.transient(parent)
        self.grab_set()
        self.overrideredirect(True)
        self.configure(bg=self._t.modal_border)
        self.cv = tk.Canvas(self, highlightthickness=0, bd=0,
                            width=width, height=height)
        self.cv.pack(fill="both", expand=True)
        self.inner = tk.Frame(self.cv, bg=self._t.modal_fill)
        self._win = self.cv.create_window(0, 0, window=self.inner, anchor="nw")
        self.cv.bind("<Configure>", lambda e: self._draw())
        self.after(20, self._draw)
        # 标题区
        self.head = tk.Frame(self.inner, bg=self._t.modal_fill)
        self.head.pack(fill="x", padx=40, pady=(34, 6))
        if title:
            tk.Label(self.head, text=title, bg=self._t.modal_fill,
                     fg=self._t.modal_title,
                     font=self._t.serif(SIZE["page_title"], bold=True)).pack(anchor="w")
        if subtitle:
            tk.Label(self.head, text=subtitle, bg=self._t.modal_fill,
                     fg=self._t.modal_sub, font=self._t.sans(SIZE["caption"])).pack(anchor="w", pady=(4, 0))
        self.body = tk.Frame(self.inner, bg=self._t.modal_fill)
        self.body.pack(fill="both", expand=True, padx=40, pady=(10, 10))
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
        self.geometry("%dx%d+%d+%d" % (w, h, max(0, x), max(0, y)))


class ConfirmModal(ModalShell):
    """通用确认：body 放提示文字，底部 取消 / 确认。result 经 callback 回传。"""

    def __init__(self, parent, edition="student", title="确认操作",
                 message="请确认是否继续当前操作。", on_confirm=None,
                 confirm_text="确认", cancel_text="取消", **kw):
        super().__init__(parent, edition=edition, title=title, width=620, height=300)
        tk.Label(self.body, text=message, bg=self._t.modal_fill, fg=self._t.modal_sub,
                 font=self._t.sans(SIZE["body"]), wraplength=520, justify="left",
                 anchor="w").pack(anchor="w", pady=(8, 24))
        btns = tk.Frame(self.body, bg=self._t.modal_fill)
        btns.pack(fill="x", pady=(6, 10))
        cancel = RoundButton(btns, text=cancel_text, edition=edition, style="secondary",
                              height=46, font=self._t.sans(SIZE["body"], bold=True),
                              command=self._safe_close)
        cancel.pack(side="left")
        ok = RoundButton(btns, text=confirm_text, edition=edition, style="primary",
                         height=46, font=self._t.sans(SIZE["body"], bold=True),
                         command=lambda: (on_confirm and on_confirm(), self._safe_close()))
        ok.pack(side="right")
        self.center_on(parent)


class PricingCard(tk.Frame):
    """单张套餐卡（upgrade-modal / pricing/*.svg）。"""

    def __init__(self, master, edition, tier, on_choose, **kw):
        super().__init__(master, **kw)
        t = get_theme(edition)
        self._t = t
        self._tier = tier
        rec = tier.get("recommended")
        fill = t.info_fill if rec else t.surface
        border = t.info_border if rec else t.border
        self.cv = tk.Canvas(self, highlightthickness=0, bd=0)
        self.cv.pack(fill="both", expand=True)
        self.inner = tk.Frame(self.cv, bg=fill)
        self._win = self.cv.create_window(0, 0, window=self.inner, anchor="nw")
        self.cv.bind("<Configure>", lambda e: self._draw())
        # 内容
        tk.Label(self.inner, text=tier["name"], bg=fill, fg=t.modal_title,
                 font=t.sans(SIZE["section_title"], bold=True)).pack(anchor="w", padx=20, pady=(24, 0))
        tk.Label(self.inner, text=tier.get("desc", ""), bg=fill, fg=t.muted,
                 font=t.sans(SIZE["caption"])).pack(anchor="w", padx=20, pady=(4, 0))
        tk.Label(self.inner, text=tier["price"], bg=fill, fg=t.primary,
                 font=t.latin(24, bold=True)).pack(anchor="w", padx=20, pady=(18, 0))
        btn = RoundButton(self.inner, text="推荐" if rec else "选择此方案",
                          edition=edition, style="primary" if rec else "secondary",
                          height=28, font=t.sans(11, bold=True),
                          command=lambda: on_choose(tier["key"]))
        btn.pack(anchor="w", padx=20, pady=(14, 22))
        self._fill = fill
        self._border = border
        self.after(20, self._draw)

    def _draw(self):
        self.cv.delete("all")
        w = self.cv.winfo_width()
        h = self.cv.winfo_height()
        if w < 2 or h < 2:
            self.after(20, self._draw)
            return
        _rounded_rect(self.cv, 1, 1, w - 1, h - 1, RADIUS["card"],
                      fill=self._fill, outline=self._border, width=2)
        self._win = self.cv.create_window(0, 0, window=self.inner, anchor="nw")
        self.cv.coords(self._win, 0, 0)
        self.cv.itemconfig(self._win, width=w)


class UpgradeModal(ModalShell):
    """升级正式版：套餐卡行 + 已有激活码入口。"""

    def __init__(self, parent, edition="student", tiers=None,
                 on_tier=None, on_activate=None, **kw):
        super().__init__(parent, edition=edition,
                         title="%s · 升级正式版" % ("学生版" if edition == "student" else "导师版"),
                         subtitle="选择适合你的使用方式。", width=760, height=520)
        self._tiers = tiers or self._default_tiers(edition)
        self._on_tier = on_tier
        self._on_activate = on_activate
        # 套餐卡行
        cards = tk.Frame(self.body, bg=self._t.modal_fill)
        cards.pack(fill="x", pady=(6, 10))
        for tier in self._tiers:
            c = PricingCard(cards, edition, tier, self._choose,
                            width=160 if tier["key"] != "lifetime" else 140, height=150)
            c.pack(side="left", padx=(0, 12) if tier is not self._tiers[-1] else 0)
        # 已有激活码
        row = tk.Frame(self.body, bg=self._t.modal_fill)
        row.pack(fill="x", pady=(10, 4))
        tk.Label(row, text="已有激活码？", bg=self._t.modal_fill, fg=self._t.modal_title,
                 font=self._t.sans(SIZE["section_title"], bold=True)).pack(anchor="w", pady=(4, 6))
        ent_row = tk.Frame(self.body, bg=self._t.modal_fill)
        ent_row.pack(fill="x")
        self._code_var = tk.StringVar()
        ent = tk.Entry(ent_row, textvariable=self._code_var, font=self._t.sans(SIZE["body"]),
                       relief="solid", bd=1, highlightthickness=1,
                       highlightbackground=self._t.select_border)
        ent.pack(side="left", fill="x", expand=True, ipady=6)
        act = RoundButton(ent_row, text="激活", edition=edition, style="primary",
                          height=40, font=self._t.sans(SIZE["body"], bold=True),
                          command=self._activate)
        act.pack(side="right", padx=(10, 0))
        self.center_on(parent)

    @staticmethod
    def _default_tiers(edition):
        if edition == "student":
            return [
                {"key": "once", "name": "次卡", "desc": "偶尔使用", "price": "¥15"},
                {"key": "weekly", "name": "周卡", "desc": "短期项目", "price": "¥39"},
                {"key": "monthly", "name": "月卡", "desc": "持续使用", "price": "¥99", "recommended": True},
                {"key": "lifetime", "name": "永久", "desc": "长期使用", "price": "¥299"},
            ]
        return [
            {"key": "once", "name": "次卡", "desc": "偶尔使用", "price": "¥20"},
            {"key": "monthly", "name": "月卡", "desc": "持续使用", "price": "¥299", "recommended": True},
            {"key": "lifetime", "name": "永久", "desc": "长期使用", "price": "¥999"},
        ]

    def _choose(self, key):
        if self._on_tier:
            self._on_tier(key)

    def _activate(self):
        code = self._code_var.get().strip()
        if not code:
            return
        if self._on_activate:
            self._on_activate(code)
        self._safe_close()
