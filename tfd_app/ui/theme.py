# -*- coding: utf-8 -*-
"""
论文格式医生 · 中国版共享设计令牌（唯一来源）
==============================================

严格照施工包 ``03_SHARED/tokens/DESIGN_TOKENS.json`` + 各版 SVG 零件规格。
学生版 / 导师版共用同一套零件，只换调色板（edition 参数）。

配色铁律（AGENT_MASTER_INSTRUCTION.md）：
  - 宣纸米白底 + 学术蓝/墨青主色 + 少量印章红
  - 圆角：控件 8 / 卡片 12 / 弹窗 14 / 胶囊 999
  - 字号：标题 26 / 区块 20 / 正文 14 / 说明 12
  - 无渐变、无玻璃拟态、无霓虹、无厚重阴影
  - 朱砂红只作身份印章 + 极克制点缀，不作主操作色
"""
from __future__ import annotations

import platform as _platform
import tkinter.font as tkfont

# ---------------------------------------------------------------------------
# 圆角 / 字号（来自 DESIGN_TOKENS.json，两版共用）
# ---------------------------------------------------------------------------
RADIUS = {"control": 8, "card": 12, "modal": 14, "pill": 999}
SIZE = {"page_title": 26, "section_title": 20, "body": 14, "caption": 12}
SPACING = {"xs": 4, "sm": 8, "md": 12, "lg": 16, "xl": 24, "xxl": 32}

# ---------------------------------------------------------------------------
# 两版调色板（照 DESIGN_TOKENS.json 的 student / advisor 段）
# ---------------------------------------------------------------------------
_PALETTES = {
    "student": {
        "bg": "#F7F3EA",          # 宣纸米白
        "surface": "#FFFDF9",     # 卡片面
        "primary": "#286D9F",     # 学术蓝（主操作色）
        "navy": "#193E5D",        # 标题深蓝
        "seal": "#B54D43",        # 印章红
        "muted": "#6F8493",       # 次要文字
        "border": "#D7E0E5",      # 卡片 / 控件边框
    },
    "advisor": {
        "bg": "#F5EFE5",          # 更暖的米白
        "surface": "#FFFDF9",
        "primary": "#355F73",     # 墨青（主操作色）
        "navy": "#3F3632",        # 暖褐黑
        "seal": "#B54D43",        # 印章红（同）
        "muted": "#776E68",
        "border": "#D9D0C5",
    },
}

# 由主色派生（主色加深作 hover；主色极浅作徽标底）
def _derive(edition, base):
    is_student = (edition == "student")
    return {
        "primary_hover": "#1F5A8A" if is_student else "#2C4F60",
        "primary_soft": "#EAF3FB" if is_student else "#EAF0F2",
        "secondary_border": "#BFD0DA",     # 次按钮 / 选择行边框（SVG 实测）
        "dropzone_fill": "#F7FBFE",
        "dropzone_border": "#BBD2E2",
        "info_fill": "#F0F6FA" if is_student else "#F1F4F2",
        "info_border": "#D5E3EC" if is_student else "#DCD3C8",
        "select_border": "#D0DCE4",
        "select_ph": "#506B7C",
        "stepper_line": "#D8E2EA",
        "step_pending_fill": "#FFFDF9",
        "step_pending_border": "#BFD0DA",
        "stepper_label_active": "#284E68" if is_student else "#3F3632",
        "stepper_label_pending": "#6F8493",
        # 状态三色（03_SHARED/components 公共，两版一致）
        "success_fill": "#F1F7F3", "success_border": "#D0E2D8",
        "success_dot": "#3E7B67", "success_text": "#3E6C5D",
        "processing_fill": "#F1F6FA", "processing_border": "#D2E0EA",
        "processing_dot": "#286D9F", "processing_text": "#355F73",
        "error_fill": "#FFF5F3", "error_border": "#E5C8C2",
        "error_dot": "#A94A43", "error_text": "#98453F",
        # 正式版徽标（green-grey，区别于试用蓝）
        "formal_fill": "#F0F2EF", "formal_text": "#4D6C61",
        # 弹窗壳
        "modal_fill": "#FFFDF9", "modal_border": "#D8E2EA",
        "modal_title": "#193E5D", "modal_sub": "#6F8493",
        # 卡片标题（区块标题用 navy）
        "card_title": base["navy"],
        # 禁用手写色
        "disabled_fill": "#ECE9E2" if is_student else "#E7E1D6",
        "disabled_text": "#9AA1A5",
        "disabled_border": "#ECE9E2" if is_student else "#E7E1D6",
    }


class Theme:
    """edition 调色板 + 派生色 + 字体解析，集中一处避免散落硬编码。"""

    def __init__(self, edition):
        self.edition = edition
        base = _PALETTES[edition]
        self.__dict__.update(base)
        self.__dict__.update(_derive(edition, base))
        self.edition_label = "学生版" if edition == "student" else "导师版"

    # ---- 字体（中文标题走宋体/思源宋体；正文走雅黑/思源黑体；拉丁数字走 Georgia）----
    @staticmethod
    def _resolve(fam):
        fams = tkfont.families()
        if fam in fams:
            return fam
        mapping = {
            "SimSun": {"Darwin": "Songti SC", "Linux": "Noto Serif CJK SC"},
            "Microsoft YaHei": {"Darwin": "PingFang SC", "Linux": "Noto Sans CJK SC"},
            "Georgia": {},
        }
        cand = mapping.get(fam, {}).get(_platform.system())
        if cand and cand in fams:
            return cand
        return fam

    def serif(self, size, bold=False):
        return (self._resolve("SimSun"), size, "bold" if bold else "normal")

    def sans(self, size, bold=False):
        return (self._resolve("Microsoft YaHei"), size, "bold" if bold else "normal")

    def latin(self, size, bold=False):
        return ("Georgia", size, "bold" if bold else "normal")


_THEME_CACHE = {}


def get_theme(edition="student"):
    if edition not in _THEME_CACHE:
        _THEME_CACHE[edition] = Theme(edition)
    return _THEME_CACHE[edition]
