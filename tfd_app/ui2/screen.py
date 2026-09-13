# -*- coding: utf-8 -*-
"""论文格式医生 · 新版界面（视图层）—— 定稿于原型 v19，外观不再改动

本模块只做两件事：① 按模块级状态绘制界面；② 把可点区域登记进 HOT。
业务、线程、弹窗、引擎调用一律在 app_ui.py（外壳层）。

芦苇 2026-09-13 交办：不要背景图，布局不要太呆板，参考海外版被认可的观感。

设计前提（为什么要这么做）
--------------------------
没有背景图，页面就是一张纯色底 —— 最容易「平、呆、死」。所以观感必须靠四招撑：

  ① 层次（色阶）  卡片面 #FFFDF9 > 信息卡 #F0F6FA > 页面底 #F7F3EA，三级明度；
                  卡片再叠一层「贴地投影」（纯 canvas 多层圆角，不用 PIL）
  ② 对比（大小）  品牌大标题（宋体 20pt）对副标题（10pt）；主按钮整宽 48 高对次按钮 36 高
  ③ 节奏（留白）  顶部信笺抬头带整条留白，卡片内紧凑、字段间只给一根极细线
  ④ 点缀（矢量）  区块标题前的圆徽标、品牌标、朱砂竖章、卡片角上极浅的大号序号

技术底座与海外版**完全一致**：纯标准库 canvas 自绘。
  - 不引 PIL（国内版源码零 PIL 依赖，引了就得改打包）
  - 不引任何第三方
  - **零图片资源** → 打包不可能再出现「资源静默丢失」，三平台渲染零差异

布局骨架
--------
    ┌─ 信笺抬头带（surface 底 + 下沿细线）────────────────┐
    │ [品牌标] 论文格式医生 [朱砂章]  〔授权胶囊〕[升级正式版→] │
│          让论文格式更简单      客服微信｜公众号｜关于｜帮助 │
    │          让论文格式更简单                             │
    ├─ 内容区（两栏等高卡片，垂直居中）─────────────────────┤
    │ ┌ 文件选择   01(水印) ┐  ┌ 处理步骤   02(水印) ┐    │
    │ │ ▤ 标题 + 细线       │  │ ⚙ 标题 + 细线        │    │
    │ │ ① 上传论文 [必选]   │  │ ①——②——③ 横排步骤    │    │
    │ │   [拖放区]          │  │ [引导卡·浅蓝底]      │    │
    │ │ ② 学校模板 [可选]   │  │ [ 开始检查 ] 主按钮  │    │
    │ │   [拖放区]          │  │ [一键修正][导出报告] │    │
    │ │ ③ 模板说明 ○按批注  │  │ [信任卡·浅蓝底]      │    │
    │ │ [底注卡·浅蓝底]     │  │                      │    │
    │ └─────────────────────┘  └──────────────────────┘    │
    ├─ 页脚（细线 + 三段式）───────────────────────────────┤
    └──────────────────────────────────────────────────────┘

用法
----
    python guofeng_nobg.py                 # 五档尺寸跑一遍，截图到 _smoke/guofeng2/
    GF_EDITION=advisor python guofeng_nobg.py
"""
from __future__ import annotations

import math
import os
import sys
import time

import tkinter as tk
import tkinter.font as tkfont

from ..ui import theme as TH          # noqa: E402
from ..assetpath import find_asset    # noqa: E402

OUT = os.environ.get("GF_OUT", r"D:\AgentSpace\_smoke\guofeng2")
SIZES = [(1100, 690), (1280, 720), (1366, 768), (1440, 900), (1600, 900)]

T = TH.get_theme("student")

# ---------------------------------------------------------------------------
# v3 覆盖：配色（学生蓝 / 导师朱砂红）+ 字号（参考稿）
# ---------------------------------------------------------------------------
PALETTE_OVERRIDE = {
    # ---------------- 学生版：整体走「那种蓝」 #1E5AA8（稿子 --theme-color） ----------------
    "student": {
        # 主色系
        "primary": "#1E5AA8",
        "primary_hover": "#1A4F92",
        "primary_soft": "#E8F1FB",
        "navy": "#14406F",
        "card_title": "#14406F",          # ★补：原 #193E5D（= 旧 navy，必须跟着走）
        "modal_title": "#14406F",         # ★补：原 #193E5D
        "radio_on": "#1E5AA8",
        # 状态 / 步进（原值残留旧主色与旧墨青）
        "processing_dot": "#1E5AA8",      # ★补：原 #286D9F —— 旧主色残留
        "processing_text": "#1A4F92",     # ★补：原 #355F73 —— 旧墨青残留
        "processing_fill": "#EDF3FB",
        "processing_border": "#CCDDEF",
        "stepper_label_active": "#14406F",   # ★补：原 #284E68
        # 边框系（跟主色同族）
        "border": "#D3DDE8",
        "secondary_border": "#B7CCE4",
        "dropzone_border": "#B7CCE4",
        "select_border": "#C9D9EA",
        "info_border": "#D2E2F2",
        "soft_border": "#D2E2F2",
        "step_pending_border": "#B7CCE4",
        "radio_off_border": "#A9BCCE",
        # 成功态（绿系，通用语义，不跟主色走）
        "success_fill": "#EEF6F1",
        "success_border": "#CBE0D6",
        # 页脚
        "footer_line": "#D6E1EC",
        "footer_text": "#6C819A",
        "footer_text_soft": "#8798AB",
    },
    # ---------------- 导师版：整体走朱砂红 #8C2D19（稿子 --theme-color） ----------------
    "advisor": {
        "primary": "#8C2D19",
        "primary_hover": "#7A2513",
        "primary_soft": "#FAEDE9",
        "navy": "#5E2416",
        "card_title": "#5E2416",          # ★补：原 #3F3632
        "modal_title": "#5E2416",         # ★补：原 #193E5D —— 学生版深蓝漏到导师版
        "radio_on": "#8C2D19",
        "processing_dot": "#8C2D19",      # ★补：原 #286D9F
        "processing_text": "#7A2513",     # ★补：原 #355F73
        "processing_fill": "#FBF0ED",
        "processing_border": "#E8D0C9",
        "stepper_label_active": "#5E2416",    # ★补：原 #3F3632
        "stepper_label_pending": "#776E68",   # ★补：原 #6F8493（学生版灰蓝漏到导师版）
        "modal_sub": "#776E68",               # ★补：原 #6F8493
        "select_ph": "#776E68",               # ★补：原 #506B7C
        "stepper_line": "#E2D2CC",            # ★补：原 #D8E2EA（冷蓝灰）
        "modal_border": "#E0CEC8",            # ★补：原 #D8E2EA
        # 边框系
        "border": "#E0CEC8",
        "secondary_border": "#DCC0B8",
        "dropzone_border": "#DCC0B8",
        "select_border": "#DEC9C3",
        "info_border": "#E6D1CB",
        "soft_border": "#E6D1CB",
        "step_pending_border": "#DCC0B8",
        "radio_off_border": "#C6B2AC",
        # 已激活态：绿灰在朱红主题里发脏，改暖褐（仍是「安静的成功态」）
        "formal_fill": "#F5E9E4",
        "formal_text": "#8A5449",
        # 成功态暖化，避免冷绿与朱红打架
        "success_fill": "#F5F1EC",
        "success_border": "#E0D6CE",
        # 页脚
        "footer_line": "#E2D2CC",
        "footer_text": "#8A736C",
        "footer_text_soft": "#A08C85",
    },
}

# 字号覆盖：只动这两项（其余与稿子本已吻合）
TYPE_OVERRIDE = {
    "subtitle": 9,          # 10 → 9   稿子 logo-sub 11px≈8pt，保留可读性取 9
    "section_title": 12,    # 14 → 12  稿子 card-title 16px≈12pt（层级更紧凑的关键）
}

# 当前版本（draw_left/draw_right 按它分支）

# 品牌标形态：a = 主题色方块 + 白羽毛笔；b = 纯羽毛笔（默认，更轻更贴"小羽毛笔"）
MARK_STYLE = os.environ.get("GF_MARK", "b")

# —— PNG 品牌标（老板 2026-09-13 交办：图标他自己改，PNG 直接放进来）——
#   GF_MARKSRC=<png 路径>  → 用 PNG；留空 → 用上面的矢量羽毛
#   GF_MARKMODE=tint       → 自动上主题色（学生蓝 / 导师红），一张图走两版
#   GF_MARKMODE=asis       → 保留 PNG 自己的颜色（含描边、渐变）
#   PNG 里的大片底色/留白会被自动裁掉，底色方块不会带进界面。
from . import mark_png

MARK_IMG_REFS = []                       # Tk 不持图片引用 → 必须自己留，否则被 GC 掉变空白

# ---------------------------------------------------------------------------
# 热区登记（视图层与外壳层的唯一接口）
#   视图层只负责「画 + 登记这块能点」；点了要做什么由外壳层（app_ui.py）决定。
#   为什么登记而不是在外壳里重算坐标：布局坐标只在这里算一次，
#   外壳复算等于把布局知识存两份，改一处忘一处 —— 这类 bug 最难查。
# ---------------------------------------------------------------------------
HOT = []          # [(x0, y0, x1, y1, action, payload)]


def hot(x0, y0, x1, y1, action, payload=None):
    """登记一个可点区域（画布坐标）。"""
    HOT.append((float(x0), float(y0), float(x1), float(y1), action, payload))


def clear_hot():
    del HOT[:]


def hit(x, y):
    """命中测试：后登记者优先（按钮登记在整块投放区之后 → 按钮压住整块）。"""
    for x0, y0, x1, y1, action, payload in reversed(HOT):
        if x0 <= x <= x1 and y0 <= y <= y1:
            return action, payload
    return None

MARK_SRC = os.environ.get("GF_MARKSRC") or (find_asset("brand_mark.png") or "")
MARK_MODE = os.environ.get("GF_MARKMODE", "asis")     # 定稿：两版都用原图色
MARK_PNG_H = 32                          # PNG 羽毛显示高度（与矢量版等高，便于对比）
# 墨量加浓：细笔画的图缩到 24~40px 会发虚，把中间调抬成实心（"" = 不加浓）
MARK_BOOST = os.environ.get("GF_MARKBOOST", "20,170")  # 定稿：加浓，保小尺寸可读
# 标与标题的间距（px）
MARK_GAP = float(os.environ.get("GF_MARKGAP", "13"))
# 垂直对齐：band = 对齐整块文字中心（现状）；title = 对齐大标题中线
MARK_VMODE = os.environ.get("GF_MARKV", "title")     # 定稿：与大标题同线


def mark_boost():
    """把 "20,170" 解析成 (20, 170)；无效或空则返回 None。"""
    if not MARK_BOOST:
        return None
    try:
        a, b = MARK_BOOST.split(",")
        return (int(a), int(b))
    except Exception:
        return None


def mark_src_path():
    """能用的 PNG 才用；路径无效就回落到矢量羽毛（不会静默变空白）。"""
    if MARK_SRC and os.path.isfile(MARK_SRC):
        return MARK_SRC
    return None
# 入口行是否用竖线分隔：1 = 保留；0 = 去掉、改用更大字距（默认）
NAV_DIV = os.environ.get("GF_NAVDIV", "0") == "1"

# 顶部入口行是否带通用符号图标（A/B 对照用；GF_NAVICON=1 开启）
NAV_ICON = os.environ.get("GF_NAVICON", "0") == "1"
NAV_ICONS = {"客服微信": "chat", "官方公众号": "megaphone",
             "关于": "info", "帮助": "help"}

# 信任卡文案变体（GF_TRUST 切换，仅本原型用于对比）
TRUST_VARIANTS = {
    # 现状（3 行版）：注意「无需联网」偏松（激活需联网）、「无需注册」与"本地处理"无逻辑关联
    "now":  {"title": "论文不出本机",
             "lines": ["全程离线运行，文件不上传", "无需注册、无需联网"]},
    # 推荐：标题短、两条都落在「隐私」且都可核验
    "rec":  {"title": "本地处理",
             "lines": ["论文不上传，全程只在本机完成", "不注册账号，不收集文档内容"]},
    # 备选一：语气更明确地强调"不上传"
    "priv": {"title": "本地处理",
             "lines": ["论文只在本机打开，不上传任何文件", "不注册账号，不采集文档内容"]},
    # 备选二：最省字（单行），适合嫌两行啰嗦时
    "short": {"title": "本地处理",
              "lines": ["论文不上传，不收集文档内容"]},
    # 备选三：把"隐私"提到标题（信息层级最高，但略抢）
    "titled": {"title": "本地处理 · 隐私优先",
               "lines": ["论文不上传，全程只在本机完成", "不注册账号，不收集文档内容"]},
}
TRUST = TRUST_VARIANTS["now"]

EDITION = "student"
# 顶部方案：a = 授权区跟在「学生版」徽章后面（老板建议）；c = 授权区留右上、退到窗口按钮下方
TOP_LAYOUT = os.environ.get("GF_TOP", "a")
# 窗口控制按钮占位（右上角，无边框窗口必须自绘）
WIN_BTN_W = 44.0
WIN_BTN_H = 28.0
WIN_BTN_N = 3

# 导师版内容更多（批量导入 / 记住模板 / 批注署名 / 交付三选一），高度另算
LEFT_H_MENTOR = 505
RIGHT_H_MENTOR = 505


def apply_style(edition):
    """把配色与字号覆盖应用到当前主题实例（仅本原型用，不改 theme.py）。"""
    for k, v in PALETTE_OVERRIDE.get(edition, {}).items():
        setattr(T, k, v)
    TYPE.update(TYPE_OVERRIDE)

TYPE = TH.TYPE
RADIUS = TH.RADIUS

# ---------------------------------------------------------------------------
# 授权状态（严格对应 license.py / trial.py 的真实判定，不发明新状态）
# ---------------------------------------------------------------------------
# gui.py::_update_trial_badge() 是权威逻辑，本原型逐条照搬：
#   is_licensed() == True  → 徽章 kind=formal，文案取 license_summary()["kind"]
#                            （正式版 · 月卡 / 周卡 / 次卡 / 永久版）
#                            按钮 → 「已激活 ✓」且置灰（不可再点）
#   is_licensed() == False → 徽章 kind=trial
#                            trials_left() > 0 → 「试用版 · 剩余 N 次」
#                            否则               → 「试用版 · 已用完」
#                            按钮 → 「升级正式版 →」（可点）
# 注：revoked（退款撤销）/ expired（到期）在真实代码里同样走 trial 分支
#     （check_local_valid() 直接返回 False），所以界面表现与「试用已用完」同构，
#     差异在弹窗文案里体现，这里不额外发明一个界面上不存在的新状态。
LIC_STATES = {
    "trial_new":  {"kind": "trial",  "text": "试用版 · 剩余 1 次",
                   "btn": "升级正式版 →", "enabled": True},
    "trial_used": {"kind": "trial",  "text": "试用版 · 已用完",
                   "btn": "升级正式版 →", "enabled": True},
    "monthly":    {"kind": "formal", "text": "正式版 · 月卡",
                   "btn": "已激活 ✓", "enabled": False},
    "permanent":  {"kind": "formal", "text": "正式版 · 永久版",
                   "btn": "已激活 ✓", "enabled": False},
}
LIC = LIC_STATES["trial_new"]
LIC_TAG = "trial_new"
LAYOUT = {}          # render 时记录关键边界，供「品牌区/授权区不重叠」断言用

# ---------------------------------------------------------------------------
# 运行时状态（由外壳层 app_ui.py 驱动）
#   默认值 == 初始态，所以视图层单独跑（原型模式）时看到的还是定稿那副样子。
# ---------------------------------------------------------------------------
STATUS = "就绪"                 # 页脚运行状态：就绪 / 正在检查… / 已完成 / 出错
STEP_INDEX = 0                  # 当前步骤 0/1/2（≤ 它的都算已完成 → 高亮）
MAIN_BTN_TEXT = "开始检查"       # 主按钮文案（随步骤变）
MAIN_BTN_ENABLED = True         # 处理中置灰，防重复点击
PICKED = ""                     # 已选论文（学生版：单文件）
BATCH_COUNT = 0                 # 已导入篇数（导师版：批量）
TEMPLATE_NAME = ""              # 已选模板文件名
TPL_REMEMBERED = False          # 是否已记住此模板


def _short(name, n=16):
    """文件名过长时截断 —— 投放区宽度有限，不许溢出到框外。"""
    name = os.path.basename(str(name))
    return name if len(name) <= n else name[:n] + "…"



def lic_colors(kind):
    """徽章配色 —— 与 widgets.Badge._colors() 完全一致。"""
    if kind == "formal":
        return T.formal_fill, T.formal_text      # #F0F2EF / #4D6C61
    return T.primary_soft, T.primary             # 蓝底 / 蓝字


# 版式常量（页面级）
PAD_X = 34          # 页面左右留白
HEAD_H = 92         # 信笺抬头带高度
FOOT_H = 36         # 页脚高度
COL_GAP = 22        # 两栏间距
CARD_PAD = 26       # 卡片内左右留白
IN_CARD_H = 84      # 底部说明卡统一高度（两卡一致 → 底对齐后上下沿都齐）

# 内容块高度（用于算卡片高度，保证两栏等高、内容垂直居中）
LEFT_H = 428          # 学生版左卡：两投放区（含按钮行 104）+ 底卡 84
RIGHT_H = 34 + 18 + 44 + 20 + 115 + 22 + 48 + 14 + 36 + 22 + 84  # = 446


# ---------------------------------------------------------------------------
# 颜色 / 几何小工具
# ---------------------------------------------------------------------------
def _rgb(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def mix(c1, c2, k):
    """c1 向 c2 混合 k 比例（k=0 → c1，k=1 → c2）。"""
    a, b = _rgb(c1), _rgb(c2)
    k = max(0.0, min(1.0, float(k)))
    return "#%02X%02X%02X" % tuple(int(round(a[i] + (b[i] - a[i]) * k)) for i in range(3))


def rpts(x0, y0, x1, y1, r, seg=18):
    """圆角矩形顶点（顺时针；屏幕坐标 y 向下）。"""
    r = max(0.0, min(float(r), (x1 - x0) / 2.0, (y1 - y0) / 2.0))
    pts = []
    for cx, cy, a0, a1 in ((x1 - r, y0 + r, -90, 0),
                           (x1 - r, y1 - r, 0, 90),
                           (x0 + r, y1 - r, 90, 180),
                           (x0 + r, y0 + r, 180, 270)):
        for i in range(seg + 1):
            a = math.radians(a0 + (a1 - a0) * i / float(seg))
            pts.append(cx + r * math.cos(a))
            pts.append(cy + r * math.sin(a))
    return pts


def rr(cv, x0, y0, x1, y1, r, fill="", outline="", width=1):
    return cv.create_polygon(rpts(x0, y0, x1, y1, r),
                             fill=fill or "", outline=outline or "", width=width)


def tw(text, font):
    """量文本宽度（tk 需要 Font 对象）。"""
    return tkfont.Font(font=font).measure(text)


# ---------------------------------------------------------------------------
# 矢量线描图标（24 设计网格，纯 canvas，零图片）
# ---------------------------------------------------------------------------
def icon(cv, name, cx, cy, size, color, w=2, bg=None):
    s = float(size) / 24.0

    def P(x, y):
        return (cx + (x - 12.0) * s, cy + (y - 12.0) * s)

    def ln(*pts):
        """扁平坐标：ln(x0, y0, x1, y1, ...)。"""
        flat = []
        for i in range(0, len(pts) - 1, 2):
            flat.extend(P(pts[i], pts[i + 1]))
        cv.create_line(*flat, fill=color, width=w, capstyle="round", joinstyle="round")

    def ov(x0, y0, x1, y1, fill="", outline=None):
        a, b = P(x0, y0), P(x1, y1)
        aa_oval(cv, a[0], a[1], b[0], b[1], fill,
                outline=color if outline is None else outline,
                bg=bg if bg is not None else T.surface, width=w)

    def pg(pts, fill=""):
        flat = []
        for x, y in pts:
            flat.extend(P(x, y))
        col = fill or color
        cv.create_polygon(*flat, fill=col, outline=col)

    if name == "doc":
        ln(7, 3, 14, 3, 19, 8, 19, 21, 7, 21, 7, 3)
        ln(14, 3, 14, 8, 19, 8)
        ln(10, 12, 16, 12)
        ln(10, 15.5, 16, 15.5)
    elif name == "sliders":
        ln(4, 8, 20, 8)
        ln(4, 16, 20, 16)
        ov(8.4, 5.5, 12.4, 10.5, fill=color, outline=color)
        ov(13.4, 13.5, 17.4, 18.5, fill=color, outline=color)
    elif name == "quill":
        # 羽毛笔（源自 icon.png 的轮廓提取：主轴 -60°、长宽比 3.92）
        #   ① 柳叶形羽片：沿主轴采样的上下边界合成闭合多边形
        pg([(6.8, 16.8), (7.2, 15.4), (7.7, 14.1), (8.4, 12.9), (8.9, 11.5),
            (10.5, 9.1), (12.0, 6.8), (13.7, 4.5), (14.8, 3.5), (16.0, 2.6),
            (17.6, 2.4), (18.2, 3.8), (18.0, 5.3), (17.7, 6.7), (16.6, 9.4),
            (15.4, 11.9), (13.9, 14.3), (13.1, 15.6), (12.3, 16.7),
            (11.4, 17.8), (10.4, 18.8)])
        #   ② 左下伸出的细笔杆
        a0, b0 = P(7.0, 22.2), P(8.6, 18.3)
        cv.create_line(a0[0], a0[1], b0[0], b0[1], fill=color,
                       width=max(1.0, w * 0.85), capstyle="round")
    elif name == "chat":
        # 客服：对话气泡（通用符号，不用微信官方 Logo）
        ln(4.5, 5.5, 19.5, 5.5, 19.5, 15, 4.5, 15, 4.5, 5.5)
        ln(9, 15, 7, 19.5, 12.5, 15)
    elif name == "megaphone":
        # 公众号：广播喇叭（通用符号）
        pg([(4, 9), (9.5, 9), (17, 4), (17, 20), (9.5, 15), (4, 15)])
        ln(6.5, 15, 6.5, 19.5, 9.5, 19.5)
    elif name == "info":
        ov(4, 4, 20, 20)
        ln(12, 11, 12, 16.5)
        ov(11, 6.6, 13, 8.6, fill=color, outline=color)
    elif name == "help":
        ov(4, 4, 20, 20)
        ln(9.6, 9.8, 11.4, 7.9, 14.4, 8.7, 14.4, 11.2, 12, 12.8, 12, 14.6)
        ov(11, 16.6, 13, 18.6, fill=color, outline=color)
    elif name == "globe":
        # 网站图标：外圆 + 赤道 + 中线经圈（纯矢量，不用 emoji 🌐）
        ov(3.6, 3.6, 20.4, 20.4)
        ln(3.6, 12, 20.4, 12)
        ov(8.6, 3.6, 15.4, 20.4)
    elif name == "shield":
        ln(12, 3.5, 19.5, 6.6, 19.5, 12, 12, 20.5, 4.5, 12, 4.5, 6.6, 12, 3.5)
        ln(9, 12, 11.3, 14.4, 15.3, 9.7)
    elif name == "bulb":
        ov(8, 3.6, 16, 11.6)
        ln(12, 11.6, 12, 14.6)
        ln(10, 14.6, 14, 14.6)
        ln(10.6, 17.6, 13.4, 17.6)
    elif name == "search":
        ov(4.6, 4.6, 15.4, 15.4)
        ln(14.4, 14.4, 20, 20)
    elif name == "upload":
        ln(12, 16, 12, 5.5)
        ln(7.6, 9.9, 12, 5.5, 16.4, 9.9)
        ln(6, 19, 18, 19)
    elif name == "ruler":
        ln(3, 9, 21, 9, 21, 15, 3, 15, 3, 9)
        ln(7, 9, 7, 12)
        ln(11, 9, 11, 13.5)
        ln(15, 9, 15, 12)
        ln(19, 9, 19, 13.5)
    elif name == "book":
        pg([(12, 6.4), (21, 9.4), (21, 18.6), (12, 15.6)])
        pg([(12, 6.4), (3, 9.4), (3, 18.6), (12, 15.6)])
        ln(12, 6.4, 12, 15.6)


# ---------------------------------------------------------------------------
# 复用件
# ---------------------------------------------------------------------------
def badge(cv, x, y, size, name, fill=None, fg=None):
    """柔和圆底 + 线描图标（区块标题的行首印记）。"""
    fill = fill or T.primary_soft
    fg = fg or T.primary
    aa_oval(cv, x, y, x + size, y + size, fill, outline=fill)
    icon(cv, name, x + size / 2.0, y + size / 2.0, size * 0.60, fg, 2)


def pill(cv, x, y, text, fill=None, fg=None):
    """小胶囊徽标（必选 / 可选）。"""
    f = T.sans(TYPE["caption"])
    w = tw(text, f) + 15
    h = 17
    fill = fill or T.soft_fill
    rr(cv, x, y, x + w, y + h, h / 2.0, fill=fill, outline=fill)
    cv.create_text(x + w / 2.0, y + h / 2.0 + 0.5, text=text,
                   fill=fg or T.soft_text, font=f)
    return w


def aa_oval(cv, x0, y0, x1, y1, fill, outline=None, bg=None, width=1.5):
    """圆 + 手工抗锯齿。

    Tk canvas 的图形**不做抗锯齿**（文字走 ClearType、图形不走），半径十几像素的小圆
    锯齿最明显。补偿法：在圆外叠两层「线色 ↔ 底色」的中间色描边，形成 1~2px 的过渡环，
    视觉上等效于 AA。零依赖、零图片资源。
    """
    bgc = bg if bg is not None else T.surface
    line = outline or fill
    cv.create_oval(x0 - 1, y0 - 1, x1 + 1, y1 + 1, fill="",
                   outline=mix(line, bgc, 0.66), width=2)
    cv.create_oval(x0 - 0.5, y0 - 0.5, x1 + 0.5, y1 + 0.5, fill="",
                   outline=mix(line, bgc, 0.30), width=2)
    cv.create_oval(x0, y0, x1, y1, fill=fill, outline=line, width=width)


def num_dot(cv, x, y, n, size=18, fill=None):
    """小实心圆 + 白数字（字段序号）。"""
    fill = fill or T.primary
    aa_oval(cv, x, y, x + size, y + size, fill, outline=fill)
    cv.create_text(x + size / 2.0, y + size / 2.0 + 0.5, text=str(n),
                   fill="#FFFFFF", font=T.sans(TYPE["step_num"]))


def hairline(cv, x0, x1, y):
    cv.create_line(x0, y, x1, y, fill=mix(T.bg, T.border, 0.85))


def card_base(cv, x, y, w, h, fill=None, outline=None, radius=None):
    """卡片面 + 贴地投影（三层，纯 canvas 圆角；不用 PIL）。"""
    radius = RADIUS["card"] if radius is None else radius
    for off, k in ((4.5, 0.30), (3, 0.17), (1.5, 0.08)):
        col = mix(T.bg, "#A2947A", k)
        rr(cv, x + 1.5, y + off, x + w - 1.5, y + h - 1.2 + off, radius,
           fill=col, outline=col)
    rr(cv, x, y, x + w, y + h, radius,
       fill=fill or T.surface, outline=outline or T.border)


def mini_btn(cv, x, y, w, h, text, primary=False):
    """投放区内的按钮（描边=次按钮 / 实底=主按钮）。"""
    BTN_LOG.append(text)
    hot(x, y, x + w, y + h, "btn:" + text)
    rr(cv, x, y, x + w, y + h, 8,
       fill=T.primary if primary else T.surface,
       outline=T.primary if primary else T.secondary_border)
    cv.create_text(x + w / 2.0, y + h / 2.0 + 0.5, text=text,
                   fill="#FFFFFF" if primary else T.primary,
                   font=T.sans(TYPE["btn_small"], bold=True))


def dropzone(cv, x, y, w, h, main, sub, buttons=(), primary_idx=0):
    """投放区：主/副文案 + 按钮行。

    照真实 DropZone 的做法 —— **整块可点**与**框内明确按钮**并存
    （gui.py L707：主/副文案 +「选择文件」次按钮；整块可点，接既有 _pick_input）。
    buttons 为空则退回纯文案样式。h 建议 104（含按钮行）或 72（无按钮）。
    """
    rr(cv, x, y, x + w, y + h, 10, fill=T.dropzone_fill, outline=T.dropzone_border)
    hot(x, y, x + w, y + h, "zone:" + main)      # 整块可点（按钮在其后登记 → 优先命中）
    cx = x + w / 2.0
    if buttons:
        icon(cv, "upload", cx, y + 24, 22, T.primary, 2)
        cv.create_text(cx, y + 47, text=main, fill=T.primary,
                       font=T.sans(TYPE["body"], bold=True))
        cv.create_text(cx, y + 64, text=sub, fill=T.muted,
                       font=T.sans(TYPE["caption"]))
        f = T.sans(TYPE["btn_small"], bold=True)
        ws = [int(tw(b, f)) + 26 for b in buttons]
        gap = 10
        bx = cx - (sum(ws) + gap * (len(buttons) - 1)) / 2.0
        for i, (b, bw) in enumerate(zip(buttons, ws)):
            mini_btn(cv, bx, y + 72, bw, 30, b, primary=(i == primary_idx))
            bx += bw + gap
    else:
        icon(cv, "upload", cx, y + h * 0.32, 24, T.primary, 2)
        cv.create_text(cx, y + h * 0.63, text=main, fill=T.primary,
                       font=T.sans(TYPE["body"], bold=True))
        cv.create_text(cx, y + h * 0.83, text=sub, fill=T.muted,
                       font=T.sans(TYPE["caption"]))


def info_card(cv, x, y, w, h, name, title, lines):
    """浅蓝底信息卡（色阶的第二级 —— 撑层次的关键件）。

    ``h <= 0`` → 高度自适应：标题行 44 + 每行 19 + 底部留白 14（无正文行时 54）。
    文案变短时卡片跟着变矮，不再留一片空白。
    """
    if not h:
        h = (44 + 19 * len(lines) + 14) if lines else 54
    # 自检：末行文字底不得越出卡底（超出只会在画布外静默绘制，不报错）
    _last = (y + 46 + 19 * (len(lines) - 1) + 7) if lines else (y + 29)
    if _last > y + h:
        print("     [警告] info_card 文字越出卡片：文字底 %.0f > 卡底 %.0f （%s）"
              % (_last, y + h, title), flush=True)
    rr(cv, x, y, x + w, y + h, 10, fill=T.info_fill, outline=T.info_border)
    icon(cv, name, x + 24, y + 22, 20, T.primary, 2)
    cv.create_text(x + 42, y + 22, text=title, anchor="w",
                   fill=T.primary_hover, font=T.sans(TYPE["card_title"], bold=True))
    yy = y + 46
    for ln_ in lines:
        cv.create_text(x + 42, yy, text="·  " + ln_, anchor="w",
                       fill=T.muted, font=T.sans(TYPE["caption"]))
        yy += 19


BTN_LOG = []          # 运行时记录本项目实际绘制的按钮文案（验证用）


def btn(cv, x, y, w, h, text, kind="primary", icon_name=None):
    BTN_LOG.append(text)
    hot(x, y, x + w, y + h, "btn:" + text)
    if kind == "primary":
        bg, fg, bd = T.primary, "#FFFFFF", T.primary
    elif kind == "secondary":
        bg, fg, bd = T.surface, T.primary, T.secondary_border
    else:
        bg, fg, bd = T.surface, T.muted, T.surface
    rr(cv, x, y, x + w, y + h, 9, fill=bg, outline=bd)
    f = T.sans(TYPE["btn"] if kind == "primary" else TYPE["btn_small"],
               bold=(kind == "primary"))
    label_w = tw(text, f)
    if icon_name:
        total = 20 + 9 + label_w
        ix = x + w / 2.0 - total / 2.0 + 10
        icon(cv, icon_name, ix, y + h / 2.0, 19, fg, 2)
        cv.create_text(ix + 10 + 9 + label_w / 2.0, y + h / 2.0 + 0.5,
                       text=text, fill=fg, font=f)
    else:
        cv.create_text(x + w / 2.0, y + h / 2.0 + 0.5, text=text, fill=fg, font=f)


def stepper(cv, x, y, w, names):
    """横排步骤器：①——②——③（比竖排时间线活）。"""
    n = len(names)
    seg = w / float(n)
    r = 13
    for i, name in enumerate(names):
        cx = x + seg * (i + 0.5)
        if i < n - 1:
            cv.create_line(cx + r + 7, y + r, cx + seg - r - 7, y + r,
                           fill=T.stepper_line, width=1.5)
        done = (i <= STEP_INDEX)
        aa_oval(cv, cx - r, y, cx + r, y + r * 2,
                T.primary if done else T.step_pending_fill,
                outline=T.primary if done else T.step_pending_border)
        cv.create_text(cx, y + r + 0.5, text=str(i + 1),
                       fill="#FFFFFF" if done else T.stepper_label_pending,
                       font=T.sans(TYPE["step_num"]))
        cv.create_text(cx, y + r * 2 + 14, text=name,
                       fill=T.stepper_label_active if done else T.stepper_label_pending,
                       font=T.sans(TYPE["caption"]))


def field_head(cv, x, y, n, label, tag):
    num_dot(cv, x, y, n)
    cv.create_text(x + 26, y + 9, text=label, anchor="w",
                   fill=T.navy, font=T.sans(TYPE["body"], bold=True))
    if tag:                                  # 空 tag = 不画胶囊（老板：标注多余就去掉）
        pill(cv, x + 26 + tw(label, T.sans(TYPE["body"], bold=True)) + 9, y + 1, tag)


def radio_row(cv, x, y, label, on):
    r = 9
    aa_oval(cv, x, y, x + r * 2, y + r * 2,
            T.radio_on if on else T.surface,
            outline=T.radio_on if on else T.radio_off_border)
    if on:
        aa_oval(cv, x + 4.5, y + 4.5, x + r * 2 - 4.5, y + r * 2 - 4.5,
                "#FFFFFF", outline="#FFFFFF", bg=T.radio_on)
    cv.create_text(x + r * 2 + 10, y + r, text=label, anchor="w",
                   fill=T.navy, font=T.sans(TYPE["body"]))


# ---------------------------------------------------------------------------
# 页面三段
# ---------------------------------------------------------------------------
def draw_window_buttons(cv, W):
    """右上角自绘窗口控制按钮（最小化 / 最大化 / 关闭）。

    无边框窗口系统不给按钮，必须自绘，否则用户没法关窗口。
    纯线描 muted 色，克制不抢戏；关闭按钮不做红（导师版主色已是朱沙红，再红就撞）。
    返回按钮组的最左边界。
    """
    col = T.muted
    x = float(W)
    for nm in ("close", "max", "min"):        # 右起：关闭在最外
        x0 = x - WIN_BTN_W
        cx, cy = (x0 + x) / 2.0, WIN_BTN_H / 2.0
        hot(x0, 0, x0 + WIN_BTN_W, WIN_BTN_H, "win", nm)
        if nm == "min":
            cv.create_line(cx - 5, cy + 0.5, cx + 5, cy + 0.5, fill=col, width=1)
        elif nm == "max":
            cv.create_rectangle(cx - 4.5, cy - 4.5, cx + 4.5, cy + 4.5,
                                outline=col, width=1)
        else:
            cv.create_line(cx - 4.5, cy - 4.5, cx + 4.5, cy + 4.5, fill=col, width=1)
            cv.create_line(cx + 4.5, cy - 4.5, cx - 4.5, cy + 4.5, fill=col, width=1)
        x = x0
    return x


def draw_license_block(cv, x0, cy):
    """〔授权胶囊〕+ [升级正式版 → / 已激活 ✓]，从 x0 起向右排。返回右边界。"""
    f_lic = T.sans(TYPE["caption"], bold=True)
    bh = 30.0
    y0, y1 = cy - bh / 2.0, cy + bh / 2.0

    pw = max(64, int(tw(LIC["text"], f_lic)) + 30)
    lf, lc = lic_colors(LIC["kind"])
    rr(cv, x0, y0, x0 + pw, y1, bh / 2.0, fill=lf, outline=lf)
    cv.create_text(x0 + pw / 2.0, cy + 0.5, text=LIC["text"], fill=lc, font=f_lic)

    bx0 = x0 + pw + 10
    bw = max(102, int(tw(LIC["btn"], f_lic)) + 30)
    rad_c = RADIUS.get("control", 8)
    if LIC["enabled"]:
        rr(cv, bx0, y0, bx0 + bw, y1, rad_c, fill=T.primary, outline=T.primary)
        hot(bx0, y0, bx0 + bw, y1, "license")
        cv.create_text(bx0 + bw / 2.0, cy + 0.5, text=LIC["btn"],
                       fill="#FFFFFF", font=f_lic)
    else:
        rr(cv, bx0, y0, bx0 + bw, y1, rad_c, fill=T.formal_fill, outline=T.formal_text)
        cv.create_text(bx0 + bw / 2.0, cy + 0.5, text=LIC["btn"],
                       fill=T.formal_text, font=f_lic)
    return bx0 + bw


def draw_header(cv, W, edition_label):
    """信笺抬头带：品牌区（左）+ 窗口按钮（右上角）+ 授权区 + 入口行。"""
    cv.create_rectangle(0, 0, W, HEAD_H, fill=T.surface, outline=T.surface)
    cv.create_line(0, HEAD_H, W, HEAD_H, fill=mix(T.bg, T.border, 0.9))

    # —— 窗口控制按钮（右上角最外侧，y 0~28）——
    win_left = draw_window_buttons(cv, W)

    # —— 品牌标 ——
    # 关键：按「墨迹实际比例」算紧贴宽度（mark_size），不要塞进正方形盒子 ——
    #   塞盒子会在图形两侧留下看不见的空白，把间距撑大、还让标偏离页边距线。
    ms = 40                                   # 只用于矢量标（无 PNG 时）
    mx, my = PAD_X, (HEAD_H - ms) / 2.0
    cy_hdr = HEAD_H / 2.0                     # 抬头带中心
    _png = mark_src_path()
    if _png:
        _boost = mark_boost()
        ow, oh = mark_png.mark_size(_png, MARK_PNG_H, boost=_boost)
        _col = None if MARK_MODE == "asis" else T.primary
        _im = mark_png.build_mark(cv, _png, MARK_PNG_H, _col, T.surface,
                                  box_w=ow, box_h=oh, boost=_boost)
        MARK_IMG_REFS.append(_im)
        # 垂直对齐：band = 整块文字中心；title = 大标题中线
        _cy = (cy_hdr - 11) if MARK_VMODE == "title" else cy_hdr
        cv.create_image(mx, _cy, anchor="w", image=_im)
        mark_w = ow
    elif MARK_STYLE == "a":
        # 主题色圆角方块 + 白色羽毛笔
        rr(cv, mx, my, mx + ms, my + ms, 11, fill=T.primary, outline=T.primary)
        icon(cv, "quill", mx + ms / 2.0, my + ms / 2.0, ms * 0.66, "#FFFFFF", 1.5)
        mark_w = ms
    else:
        # 纯羽毛笔：颜色接主题色 → 学生版蓝、导师版朱砂红（老板 2026-09-13 要求）
        icon(cv, "quill", mx + ms / 2.0, my + ms / 2.0, ms * 0.94, T.primary, 1.5)
        mark_w = ms

    # —— 大标题 + 副标题 ——
    f_big = T.serif(TYPE["page_title"], bold=True)
    tx = mx + mark_w + MARK_GAP
    cy = HEAD_H / 2.0
    cv.create_text(tx, cy - 11, text="论文格式医生", anchor="w",
                   fill=T.navy, font=f_big)
    cv.create_text(tx, cy + 13, text="让论文格式更简单", anchor="w",
                   fill=T.muted, font=T.sans(TYPE["subtitle"]))

    # —— 版本徽章（横版小徽章，紧贴大标题右侧）——
    f_seal = T.sans(TYPE["caption"], bold=True)
    bw2 = tw(edition_label, f_seal) + 17
    bh2 = 20.0
    sx = tx + tw("论文格式医生", f_big) + 9
    sy = (cy - 11) - bh2 / 2.0
    rr(cv, sx, sy, sx + bw2, sy + bh2, 4, fill=T.primary, outline=T.primary)
    cv.create_text(sx + bw2 / 2.0, sy + bh2 / 2.0 + 0.5, text=edition_label,
                   fill="#FFFFFF", font=f_seal)
    brand_right = sx + bw2

    # —— 授权区：两套位置方案 ——
    if TOP_LAYOUT == "c":
        # 方案 c：留在右上，退到窗口按钮下方一行（右对齐）
        f_lic = T.sans(TYPE["caption"], bold=True)
        pw = max(64, int(tw(LIC["text"], f_lic)) + 30)
        bw = max(102, int(tw(LIC["btn"], f_lic)) + 30)
        total = pw + 10 + bw
        # 右对齐到「窗口按钮左侧 - 20px」：与窗口按钮同排但互不侵入
        lic_left = (W - WIN_BTN_W * WIN_BTN_N) - 24 - total
        lic_right = draw_license_block(cv, lic_left, cy - 11)
        nav_right = W - PAD_X
    else:
        # 方案 a（老板建议）：授权区直接跟在版本徽章后面，窗口按钮独占右上角
        lic_left = sx + bw2 + 14
        lic_right = draw_license_block(cv, lic_left, cy - 11)
        nav_right = W - PAD_X

    # —— 入口行：居右下（权威口径 COPY_ENTRY_LINE）——
    # 海外版注意：前两项「客服微信 / 官方公众号」须由 edition 驱动隐藏（零微信铁律）
    f_nav = T.sans(TYPE["caption"])
    cx = nav_right
    nav = ("客服微信", "官方公众号", "关于", "帮助")
    for i, t in enumerate(reversed(nav)):
        w = int(tw(t, f_nav))
        cv.create_text(cx, 76, text=t, anchor="e", fill=T.muted, font=f_nav)
        hot(cx - w - 4, 64, cx + 4, 88, "nav", t)
        cx -= w
        if NAV_ICON:                      # 通用符号图标（自绘，非商标）
            icon(cv, NAV_ICONS[t], cx - 8, 76, 12, T.muted, 1.2, bg=T.surface)
            cx -= 17
        cx -= 9
        if i < len(nav) - 1:
            if NAV_DIV:
                cv.create_text(cx, 76, text="｜", anchor="e", fill=T.border, font=f_nav)
                cx -= 9
            else:
                cx -= 13          # 无竖线：更大字距分隔，更安静（老板 2026-09-13 倾向）

    LAYOUT.update(brand_right=brand_right, lic_left=lic_left, nav_left=cx,
                  gap=lic_left - brand_right, nav_gap=cx - brand_right,
                  win_left=win_left, lic_right=lic_right)
    return lic_left, lic_right


def sign_box(cv, x, y, w, h, text):
    """带边框的输入框示意（导师版「批注署名」）。"""
    rr(cv, x, y, x + w, y + h, RADIUS.get("control", 8),
       fill=T.surface, outline=T.secondary_border)
    hot(x, y, x + w, y + h, "signer")
    cv.create_text(x + 12, y + h / 2.0, text=text, anchor="w",
                   fill=T.navy, font=T.sans(TYPE["body"]))


def mode_pills(cv, x, y, items, active=0, gap=8):
    """交付方式三选一 · 胶囊样式（比小卡轻）。

    它是**选择器**不是按钮 —— 真实流程里由它决定第③步主按钮做什么
    （tfd-mentor 的 _fix_mode_var：annotate / fix / both，默认 annotate 不改原稿）。
    """
    f = T.sans(TYPE["caption"], bold=True)
    cx, h = x, 30.0
    for i, t in enumerate(items):
        w = tw(t, f) + 26
        on = (i == active)
        rr(cv, cx, y, cx + w, y + h, h / 2.0,
           fill=T.primary if on else T.surface,
           outline=T.primary if on else T.secondary_border)
        cv.create_text(cx + w / 2.0, y + h / 2.0 + 0.5, text=t,
                       fill="#FFFFFF" if on else T.muted, font=f)
        hot(cx, y, cx + w, y + h, "mode", i)
        cx += w + gap
    return cx


def draw_footer(cv, W, H, edition_label):
    """页脚三段：左（身份 + 运行状态）｜中（官网）｜右（信任点）。"""
    y = H - FOOT_H
    cv.create_line(PAD_X, y, W - PAD_X, y, fill=mix(T.bg, T.border, 0.9))
    f = T.sans(TYPE["footer"])
    cy = y + FOOT_H / 2.0 + 1

    # 左：软件名 · 版本，紧跟运行状态（状态反馈不能丢，故与身份并排）
    left = "论文格式医生 · %s  v1.3.111" % edition_label
    cv.create_text(PAD_X, cy, text=left, anchor="w",
                   fill=T.footer_text_soft, font=f)
    cv.create_text(PAD_X + tw(left, f) + 14, cy, text="●  " + STATUS, anchor="w",
                   fill=T.footer_text, font=f)

    # 中：官网 —— 纯文字（老板反馈地球图标像"小爬虫"，去掉；也不用 emoji）
    cv.create_text(W / 2.0, cy, text="官网：reedskill.com", anchor="center",
                   fill=T.footer_text, font=f)

    # 右：客服微信
    # ⚠️ 海外版必须由 edition 驱动隐藏（零微信铁律）
    cv.create_text(W - PAD_X, cy, text="客服微信：reedgege",
                   anchor="e", fill=T.footer_text_soft, font=f)


def draw_left(cv, x, y, w, h, top):
    """左卡（文件选择）。学生版 3 字段；导师版加「批注署名」与批量导入提示。"""
    mentor = (EDITION == "advisor")
    card_base(cv, x, y, w, h)
    p = CARD_PAD
    cy = y + top

    badge(cv, x + p, cy, 26, "doc")
    cv.create_text(x + p + 36, cy + 13, text="文件选择", anchor="w",
                   fill=T.navy, font=T.sans(TYPE["section_title"], bold=True))
    cv.create_text(x + w - p, cy - 2, text="01", anchor="ne",
                   fill=mix(T.bg, T.navy, 0.12), font=T.serif(24))
    cy += 34
    hairline(cv, x + p, x + w - p, cy)
    cy += 18

    # ① 上传论文（导师版两个入口：多选文件 / 整文件夹，照 gui.py L944-950）
    if mentor:
        field_head(cv, x + p, cy, 1, "上传论文", "可多选")
        cy += 22
        dropzone(cv, x + p, cy, w - p * 2, 104,
                 ("已导入 %d 篇论文" % BATCH_COUNT) if BATCH_COUNT else "批量导入论文",
                 "可选择文件夹，也可选择多个文件",
                 buttons=("选择文件", "选择文件夹"))
    else:
        field_head(cv, x + p, cy, 1, "上传论文", "必选")
        cy += 22
        dropzone(cv, x + p, cy, w - p * 2, 104,
                 _short(PICKED) if PICKED else "点击选择论文文件",
                 "Word 文档（.docx）", buttons=("选择文件",))
    cy += 126

    # ② 学校模板：框内两个按钮 ——「选择模板」+「记住此模板」
    #    （真实文案见 gui.py L982/L1403：「🔒 记住此模板」→「✓ 已记住（点此取消）」；
    #      原型按跨平台原则改用矢量图标，不用 🔒/✓ 字符）
    field_head(cv, x + p, cy, 2, "学校模板", "可选")
    cy += 22
    dropzone(cv, x + p, cy, w - p * 2, 104,
             _short(TEMPLATE_NAME) if TEMPLATE_NAME else "点击选择学校模板",
             ("已记住，下次自动载入" if TPL_REMEMBERED
              else "记住后可自动载入，不用每次重选") if mentor
             else ("已记住此模板" if TPL_REMEMBERED else "上传后按模板要求检查"),
             buttons=("选择模板", "记住此模板"))
    cy += 126

    if mentor:
        # ③ 批注署名（导师版专有；出现在 Word 批注气泡作者栏）
        field_head(cv, x + p, cy, 3, "批注署名", "只批注时生效")
        cy += 22
        sign_box(cv, x + p, cy, w - p * 2, 34, "论文格式医生·导师版")
        cy += 34 + 20
    # 学生版到此为止：真实界面左卡顺序 = 上传区 → 学校模板 → 模板说明（浅底框）。
    # 原「③ 模板说明 + 单选」是与下方浅底框语义重复的自造交互，已删除（老板 2026-09-13）。

    # 底部说明卡：钉在卡片底部（与右卡底对齐 —— 老板 2026-09-13 要求）
    # 底部说明框：标题用权威文案「模板说明」（gui.py COPY_NOTE_TITLE），两版统一；
    # 正文照 COPY_NOTE_BODY 精简成两行
    info_card(cv, x + p, y + h - CARD_PAD - IN_CARD_H, w - p * 2, IN_CARD_H,
              "ruler", "模板说明",
              ["请优先使用学校官方模板（批注中通常写明要求）",
               "无批注时按模板样式／通用规范处理"])


def draw_right(cv, x, y, w, h, top):
    """右卡（处理步骤）。

    导师版：步进器 → 交付方式（三选一）→ 三步看懂 → **唯一主按钮** → 信任卡。
      不给「一键修正 / 导出报告」次按钮：修正已含在交付方式里（语义重复），
      检查报告是完成后自动产出的次要交付物（gui.py L2055），不占按钮位。
    学生版：无交付方式概念，保持主按钮 + 两个次按钮。
    """
    mentor = (EDITION == "advisor")
    card_base(cv, x, y, w, h)
    p = CARD_PAD
    cy = y + top

    badge(cv, x + p, cy, 26, "sliders")
    cv.create_text(x + p + 36, cy + 13, text="处理步骤", anchor="w",
                   fill=T.navy, font=T.sans(TYPE["section_title"], bold=True))
    cv.create_text(x + w - p, cy - 2, text="02", anchor="ne",
                   fill=mix(T.bg, T.navy, 0.12), font=T.serif(24))
    cy += 34
    hairline(cv, x + p, x + w - p, cy)
    cy += 18

    stepper(cv, x + p, cy, w - p * 2,
            ("上传论文", "检查格式", "交付") if mentor
            else ("上传论文", "检查格式", "一键修正"))
    cy += 44 + 20

    if mentor:
        # 交付方式：选择器（决定第③步生成什么），所以下面不再放独立的「一键修正」按钮
        field_head(cv, x + p, cy, 1, "交付方式", "")
        cy += 22
        mode_pills(cv, x + p, cy, ("① 只批注", "② 一键修正", "③ 两者都要"), active=0)
        cy += 30 + 22
        info_card(cv, x + p, cy, w - p * 2, 0, "bulb", "三步看懂",
                  ["导入单篇，或整个文件夹",
                   "先检查格式，原文件不动",
                   "再按所选方式生成"])
        cy += 115 + 22
    else:
        info_card(cv, x + p, cy, w - p * 2, 0, "bulb", "三步看懂",
                  ["上传论文，可选本校模板",
                   "检查后逐条列出格式问题",
                   "一键修正并导出新文档"])
        cy += 115 + 22

    btn(cv, x + p, cy, w - p * 2, 48, MAIN_BTN_TEXT,
        "primary" if MAIN_BTN_ENABLED else "ghost",
        icon_name="search" if MAIN_BTN_ENABLED else None)

    if not mentor:
        cy += 48 + 14
        half = (w - p * 2 - 12) / 2.0
        btn(cv, x + p, cy, half, 36, "一键修正", "secondary")
        btn(cv, x + p + half + 12, cy, half, 36, "导出报告", "ghost")

    # 底部说明卡：钉在卡片底部 → 与左卡「模板优先级」上下沿对齐
    info_card(cv, x + p, y + h - CARD_PAD - IN_CARD_H, w - p * 2, IN_CARD_H,
              "shield", TRUST["title"], TRUST["lines"])


# 主渲染
# ---------------------------------------------------------------------------
def render(cv, W, H, edition_label):
    clear_hot()
    cv.delete("all")
    cv.create_rectangle(0, 0, W + 1, H + 1, fill=T.bg, outline=T.bg)

    draw_header(cv, W, edition_label)
    draw_footer(cv, W, H, edition_label)

    # 内容区：宽度随窗口（上限 1180），两栏等高、垂直居中
    avail_top = HEAD_H
    avail_bot = H - FOOT_H
    avail_h = avail_bot - avail_top
    cont_w = min(W - PAD_X * 2, 1180)
    x0 = (W - cont_w) / 2.0
    # 左右栏比例照稿子 workspace 的 1.15 : 0.85（左宽右窄）
    _cols = cont_w - COL_GAP
    lw = _cols * 1.15 / 2.0
    rw = _cols * 0.85 / 2.0
    col_w = lw

    content_h = max(LEFT_H_MENTOR, RIGHT_H_MENTOR) if EDITION == "advisor" \
        else max(LEFT_H, RIGHT_H)
    card_h = min(avail_h - 24, max(content_h + 44, int(avail_h * 0.86)))
    cy = avail_top + (avail_h - card_h) / 2.0
    top = (card_h - content_h) / 2.0

    draw_left(cv, x0, cy, lw, card_h, top)
    draw_right(cv, x0 + lw + COL_GAP, cy, rw, card_h, top)

    return {
        "card": (x0, cy, lw, card_h),
        "content_h": content_h,
        "top_off": top,
        "avail_h": avail_h,
        # 采样点（都落在「确定没有内容」的地方，保证数据可信）
        "p_header": (W * 0.5, 6),
        "p_page": (PAD_X / 2.0, cy + card_h / 2.0),
        "p_cardL": (x0 + 10, cy + 10),
        "p_cardR": (x0 + lw + COL_GAP + 10, cy + 10),
        # 坐标级布局断言（不依赖截图，故不受屏幕尺寸约束）
        "brand_right": LAYOUT.get("brand_right"),
        "lic_left": LAYOUT.get("lic_left"),
        "lic_gap": LAYOUT.get("gap"),
        "nav_gap": LAYOUT.get("nav_gap"),
        "lic_right": LAYOUT.get("lic_right"),
        "win_left": LAYOUT.get("win_left"),
        "right_edge": LAYOUT.get("lic_left") is not None and W - PAD_X,
    }
