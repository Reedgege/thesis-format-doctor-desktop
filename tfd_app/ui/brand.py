# -*- coding: utf-8 -*-
"""品牌标识（照 01_STUDENT/brand/student-mark.svg + student-seal.svg 精确绘制）。

品牌区与朱砂竖章都画在 ``backdrop.TextureCanvas`` 上：宣纸底图从四周透出来，
不会像 tk.Label 那样在标题后面留一块纯色方块（那会让背景看起来“没渲染”）。
画布尺寸按实际字体测量，DPI/字体放大时标题不会被裁、也不会挤到竖章上。
"""
from __future__ import annotations

import os
import tkinter as tk
import tkinter.font as tkfont

from . import backdrop
from .. import assetpath
from .theme import get_theme, TYPE

# 做旧传统印章 PNG（抄作业：形状/色/竖排两字照设计文件；优化：边缘侵蚀+印泥斑驳+楷体）。
# 运行时走 assetpath.find_asset，打包后也能找到。
_SEAL_SS = 3

# student-mark.svg：画布 410x78，笔画 #286D9F 宽 4，标题/副标题左边界 x=112
_MARK_BOX_W, _MARK_BOX_H = 410, 78
_MARK_INK_TOP, _MARK_INK_BOTTOM = 14, 62      # SVG 里 mark 墨迹的纵向范围
# TYPE_SCALE.md §2 硬约束：界面上的 mark（学士帽图形）高度必须 34～38px，
# 与大标题 20pt 视觉重量相称；这里按墨迹高度等比缩放，不许再放大。
_MARK_H = 36
_MARK_SCALE = _MARK_H / float(_MARK_INK_BOTTOM - _MARK_INK_TOP)
_TITLE_X = int(round(112 * _MARK_SCALE))
# 印章形状严格照 student-seal.svg / advisor-seal.svg：
# 竖圆角矩形，实体宽:高 = 78:102（rx 占比 18/78≈0.23），朱砂 #B54D43，竖排两字。
# 尺寸不再照抄 SVG 的固定像素（那会撑爆窗口），改由权威 TYPE["seal"] 反推（见 build_seal）。
_SEAL_INK = "#FFF8ED"                   # seal.svg 字色
_SEAL_RX_RATIO = 18.0 / 78.0            # seal.svg 圆角 rx / 实体宽

BRAND_TITLE = "论文格式医生"


def _subtitle(edition):
    """副标题（视觉稿文案：学生版 = 让论文格式更简单）。"""
    return "让论文格式更简单" if edition == "student" else "导师版 · 批量论文处理"


def _fonts(edition):
    t = get_theme(edition)
    f1 = tkfont.Font(font=t.serif(TYPE["page_title"], bold=True))
    f2 = tkfont.Font(font=t.sans(TYPE["subtitle"]))
    return f1, f2


def brand_size(edition="student"):
    """品牌区画布的 (宽, 高)：按真实字体测量，保证大标题不被裁、不压到朱砂竖章。"""
    f1, f2 = _fonts(edition)
    w = _TITLE_X + max(f1.measure(BRAND_TITLE), f2.measure(_subtitle(edition))) + 12
    h = max(int(round(_MARK_BOX_H * _MARK_SCALE)),
            f1.metrics("linespace") + f2.metrics("linespace") + 16)
    return int(w), int(h)


def build_brand(master, edition="student"):
    """品牌区：线描书本 mark + 「论文格式医生」大标题 + 副标题。返回画布（用法同 Frame）。"""
    t = get_theme(edition)
    f1, f2 = _fonts(edition)
    lh1, lh2 = f1.metrics("linespace"), f2.metrics("linespace")
    w, h = brand_size(edition)
    cv = backdrop.TextureCanvas(master, edition=edition, width=w, height=h)
    # mark 按 TYPE_SCALE 的 34～38px 等比缩到界面坐标，再在画布内垂直居中，保持与文字成组
    _draw_logo(cv, t.primary, (h - _MARK_BOX_H * _MARK_SCALE) / 2.0)
    top = (h - (lh1 + lh2)) / 2.0
    # SVG 里的 y 是基线；画布文字按行中心定位，故用行高的一半换算
    cv.create_text(_TITLE_X, top + lh1 / 2.0, text=BRAND_TITLE, fill=t.navy,
                   anchor="w", font=t.serif(TYPE["page_title"], bold=True), tags="brand")
    cv.create_text(_TITLE_X + 1, top + lh1 + lh2 / 2.0, text=_subtitle(edition),
                   fill=t.muted, anchor="w", font=t.sans(TYPE["subtitle"]), tags="brand")
    return cv


def _bezier(p0, p1, p2, p3, steps=8):
    """三次贝塞尔采样（纯 python，用来还原 SVG 底弧，不引第三方依赖）。"""
    pts = []
    for i in range(steps + 1):
        t = i / float(steps)
        u = 1.0 - t
        x = (u * u * u * p0[0] + 3 * u * u * t * p1[0]
             + 3 * u * t * t * p2[0] + t * t * t * p3[0])
        y = (u * u * u * p0[1] + 3 * u * u * t * p1[1]
             + 3 * u * t * t * p2[1] + t * t * t * p3[1])
        pts += [x, y]
    return pts


def _draw_logo(canvas, color, dy=0.0):
    # 照 mark.svg 路径：M25 24l30-10 30 10-30 10z / M34 31v20c0 8 42 8 42 0V31
    #                   / M55 34v28 / M85 24v28（旧实现漏了最后一段右竖）
    # SVG 设计坐标（画布 410×78）→ 界面坐标：整体乘 _MARK_SCALE，
    # 让 mark 墨迹高度正好落在 TYPE_SCALE.md §2 的 34～38px 内。
    s = _MARK_SCALE
    lw = max(1.0, 4 * s)
    def _px(v): return v * s
    def _py(v): return v * s + dy
    canvas.create_line(_px(25), _py(24), _px(55), _py(14), _px(85), _py(24),
                       _px(55), _py(34), _px(25), _py(24),
                       fill=color, width=lw, joinstyle="round", tags="brand")
    arc = _bezier((_px(34), _py(51)), (_px(34), _py(59)),
                  (_px(76), _py(59)), (_px(76), _py(51)))
    canvas.create_line(_px(34), _py(31), _px(34), _py(51), *arc,
                       _px(76), _py(51), _px(76), _py(31),
                       fill=color, width=lw, joinstyle="round", tags="brand")
    canvas.create_line(_px(55), _py(34), _px(55), _py(62),
                       fill=color, width=lw, tags="brand")
    canvas.create_line(_px(85), _py(24), _px(85), _py(52),
                       fill=color, width=lw, tags="brand")


def build_seal(master, edition="student"):
    """朱砂印章：做旧传统风（边缘侵蚀 + 印泥斑驳 + 楷体），落盘 PNG 贴图。

    抄作业（硬规矩，不动）：竖圆角矩形、朱砂 #B54D43、竖排「学/生」「导/师」、
    字色 #FFF8ED、尺寸由 TYPE["seal"]=12pt 反推（与界面布局对齐）。
    优化（老板授权）：矢量硬边 → 距离场侵蚀的自然破损边；满红底 → 印泥厚薄斑驳；
    宋体 → 楷体（simkai.ttf），更有手写篆刻韵味。
    """
    chars = ("学", "生") if edition == "student" else ("导", "师")
    fnt = tkfont.Font(font=get_theme(edition).serif(TYPE["seal"], bold=True))
    cw = max(8, fnt.measure(chars[0]))
    lh = max(8, fnt.metrics("linespace"))
    pad_x = max(3, int(round(cw * 0.30)))
    gap = max(1, int(round(lh * 0.12)))
    pad_y = max(3, int(round(lh * 0.20)))
    x1, y1 = 7, 5
    w = float(cw + 2 * pad_x)
    h = float(2 * lh + gap + 2 * pad_y)
    cv = backdrop.TextureCanvas(master, edition=edition,
                                width=int(round(x1 + w + 7)),
                                height=int(round(y1 + h)))
    png = assetpath.find_asset("seal_%s.png" % edition)
    if png:
        try:
            img = tk.PhotoImage(file=png).subsample(_SEAL_SS)
            cv.create_image(x1 + w / 2.0, y1 + h / 2.0, image=img,
                            anchor="center", tags="brand")
            cv._seal_img = img             # 防止 PhotoImage 被 GC
            return cv
        except Exception:
            pass
    # 兜底矢量绘制（资源缺失时仍可显示，不应触发）
    t = get_theme(edition)
    rr = min(w * _SEAL_RX_RATIO, h / 2)
    pts = [x1 + rr, y1, x1 + w - rr, y1, x1 + w, y1, x1 + w, y1 + rr,
           x1 + w, y1 + h - rr, x1 + w, y1 + h, x1 + w - rr, y1 + h,
           x1 + rr, y1 + h, x1, y1 + h, x1, y1 + h - rr, x1, y1 + rr,
           x1, y1]
    cv.create_polygon(pts, smooth=True, fill=t.seal, outline=t.seal, tags="brand")
    cx = x1 + w / 2
    cy1 = y1 + (pad_y + lh * 0.5)
    cy2 = y1 + (pad_y + lh + gap + lh * 0.5)
    cv.create_text(cx, cy1, text=chars[0], fill=_SEAL_INK,
                   font=t.serif(TYPE["seal"], bold=True),
                   anchor="center", tags="brand")
    cv.create_text(cx, cy2, text=chars[1], fill=_SEAL_INK,
                   font=t.serif(TYPE["seal"], bold=True),
                   anchor="center", tags="brand")
    return cv


def seal_size(master, edition="student"):
    """返回印章画布建议尺寸 (w, h)，供布局预留顶栏高度（与 build_seal 同算法）。"""
    fnt = tkfont.Font(font=get_theme(edition).serif(TYPE["seal"], bold=True))
    cw = max(8, fnt.measure("学"))
    lh = max(8, fnt.metrics("linespace"))
    pad_x = max(3, int(round(cw * 0.30)))
    gap = max(1, int(round(lh * 0.12)))
    pad_y = max(3, int(round(lh * 0.20)))
    w = float(cw + 2 * pad_x)
    h = float(2 * lh + gap + 2 * pad_y)
    return int(round(7 + w + 7)), int(round(5 + h))
