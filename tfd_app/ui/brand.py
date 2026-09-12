# -*- coding: utf-8 -*-
"""品牌标识（照 student-mark / advisor-mark / *-seal.svg 精确绘制）。"""
from __future__ import annotations

import tkinter as tk
from . import theme
from .theme import get_theme, SIZE, RADIUS


def build_brand(master, edition="student"):
    """品牌区：左侧 logo 字形 + 右侧「论文格式医生」标题 + 副标题。返回 Frame。"""
    t = get_theme(edition)
    f = tk.Frame(master, bg=t.bg)
    # logo 字形（照 mark.svg 的 path，stroke=primary 4px）
    logo = tk.Canvas(f, width=60, height=78, bg=t.bg, highlightthickness=0)
    logo.pack(side="left", padx=(0, 14))
    _draw_logo(logo, t.primary)
    txt = tk.Frame(f, bg=t.bg)
    txt.pack(side="left", anchor="w")
    sub = "学生版 · 学术排版助手" if edition == "student" else "导师版 · 批量论文处理"
    tk.Label(txt, text="论文格式医生", bg=t.bg, fg=t.navy,
             font=t.serif(SIZE["page_title"], bold=True)).pack(anchor="w")
    tk.Label(txt, text=sub, bg=t.bg, fg=t.muted,
             font=t.sans(SIZE["caption"])).pack(anchor="w", pady=(4, 0))
    return f


def _draw_logo(canvas, color):
    # 照 mark.svg：M25 24 l30 -10 30 10 -30 10 z   (顶部菱形书签)
    #           M34 31 v20 c0 8 42 8 42 0 V31      (左侧竖边+底弧)
    #           M55 34 v28                          (中竖)
    p = [25, 24, 55, 14, 85, 24, 55, 34, 25, 24,
         34, 31, 34, 51, 76, 51, 76, 31,
         55, 34, 55, 62]
    pts = []
    for i in range(0, len(p), 2):
        # 把 svg 坐标平移到 canvas（加少量偏移居中）
        pts += [p[i] - 18, p[i + 1] - 6]
    canvas.create_line(pts[0], pts[1], pts[2], pts[3], pts[4], pts[5],
                       pts[6], pts[7], pts[0], pts[1],
                       fill=color, width=4, joinstyle="round")
    canvas.create_line(pts[8], pts[9], pts[10], pts[11], pts[12], pts[13],
                       pts[14], pts[15], pts[8], pts[9],
                       fill=color, width=4, joinstyle="round")
    canvas.create_line(pts[16], pts[17], pts[18], pts[19],
                       fill=color, width=4, joinstyle="round")


def build_seal(master, edition="student"):
    """朱砂印章：圆角红底 + 竖排「学/生」或「导/师」。返回 Frame。"""
    t = get_theme(edition)
    chars = ("学", "生") if edition == "student" else ("导", "师")
    f = tk.Canvas(master, width=78, height=112, bg=t.bg, highlightthickness=0)
    _rounded_rect_seal(f, chars, t.seal)
    return f


def _rounded_rect_seal(canvas, chars, color):
    w, h = 78, 112
    r = 18
    x1, y1, x2, y2 = 7, 5, 7 + w, 5 + h
    rr = min(r, w / 2, h / 2)
    pts = [x1 + rr, y1, x2 - rr, y1, x2, y1, x2, y1 + rr,
           x2, y2 - rr, x2, y2, x2 - rr, y2, x1 + rr, y2,
           x1, y2, x1, y2 - rr, x1, y1 + rr, x1, y1]
    canvas.create_polygon(pts, smooth=True, fill=color, outline=color)
    fs = 30 if len(chars) == 2 and chars[0] == "学" else 27
    canvas.create_text(46, 48, text=chars[0], fill="#FFF8ED",
                       font=("SimSun", fs, "bold"), anchor="center")
    canvas.create_text(46, 82, text=chars[1], fill="#FFF8ED",
                       font=("SimSun", fs, "bold"), anchor="center")
