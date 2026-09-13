# -*- coding: utf-8 -*-
"""运行时资源路径解析 · 学生版
==================================================================

gui.py 要读窗口图标/二维码，watermark.py 要读水印图，backdrop.py 要读主界面
背景图，它们都放在 ``tfd_app/assets/``。但**打包后布局会变**（Nuitka 把包编进
exe、数据目录被拷到别处），写死一条 ``<模块目录>/assets/xxx`` 就会在打包版里
静默读不到。

本模块做两件事：
  1. ``find_asset(name)``：按几种可能的布局依次探测，返回第一个真实存在的路径；
     都没有返回 ``None``（**背景图不允许静默降级**，调用方必须显式报错，见 backdrop.py）。
  2. ``REQUIRED_ASSETS``：**必须随产物携带**的资源清单，作为 build.py 的唯一事实
     来源（``--include-data-files`` 由它生成，缺件直接红灯），并供 tests 断言。

学生版清单里**只有学生版自己的背景图** ``student_background.png``；导师版专属的
``advisor_background.png`` 不得混进本仓库（BRIEF 第七节第 6 条）。
"""
from __future__ import annotations

import os
import sys

# 必须随产物携带的资源（build.py 据此生成 --include-data-files=…）。
REQUIRED_ASSETS = (
    "brand_mark.png",  # 界面品牌标（羽毛；MarkSrc 优先用它，缺则退回矢量）
    "icon.png",                 # 窗口图标（Tk 用 PNG）
    "icon.ico",                 # Windows 窗口/文件图标
    "icon.icns",                # macOS
    "qrcode.png",  # 公众号关注码（旧 UI 沿用；新版 UI 用 gz_qrcode.png）
    "gz_qrcode.png",  # 官方公众号二维码（关注入口浮层）
    "kf_qrcode.png",  # 客服微信二维码（真人客服号）
    "miniapp_qrcode.png",       # 小程序二维码
    "watermark.png",            # 试用水印图（正文穿插）
    "watermark_bg.png",         # 试用水印背景图（页眉 VML）—— 业务资源，非界面背景
    "student_background.png",   # 学生版主界面背景层（Backdrop 铺底，唯一背景层）
    "seal_student.png",         # 学生版朱砂印章（做旧传统风，PNG 贴图）
)

# 学生版仓库里**不该存在**的素材（导师版专属）。tests 会断言它不在。
FORBIDDEN_ASSETS = (
    "advisor_background.png",
)

# 不得作为主界面背景图的视觉稿（BRIEF 第七节第 4 条：严禁把整张视觉稿当背景）
MOCKUP_UPLOAD_NAMES = (
    "STUDENT_UI_REFERENCE.png",
    "STUDENT_FINAL_LAYOUT_PREVIEW.png",
    "01_UI_MOCKUP_REFERENCE.png",
)


def _asset_dir() -> str:
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets")


def _runtime_bases() -> list:
    """打包后「可执行文件所在目录」的候选（源码运行时基本用不上）。"""
    bases = []
    raws = []
    argv = getattr(sys, "argv", None) or []
    if argv:
        raws.append(argv[0])
    raws.append(getattr(sys, "executable", "") or "")
    for raw in raws:
        try:
            if raw:
                bases.append(os.path.dirname(os.path.abspath(raw)))
        except Exception:
            continue
    try:
        bases.append(os.getcwd())
    except Exception:
        pass
    out = []
    for b in bases:
        if b and b not in out:
            out.append(b)
    return out


def candidates(name: str) -> list:
    """按「最可能命中」的顺序返回候选路径（含不存在的）。"""
    here = os.path.dirname(os.path.abspath(__file__))
    cands = [
        os.path.join(here, "assets", name),          # 源码运行 / Nuitka 包内布局
        os.path.join(here, "..", "assets", name),    # 资源被放在包外一层
    ]
    for base in _runtime_bases():
        cands += [
            os.path.join(base, "assets", name),          # 资源与 exe 同级
            os.path.join(base, "tfd_app", "assets", name),
            os.path.join(base, name),
        ]
    out = []
    for c in cands:
        try:
            c = os.path.normpath(os.path.abspath(c))
        except Exception:
            continue
        if c not in out:
            out.append(c)
    return out


def find_asset(name: str):
    """返回第一个真实存在的资源路径；都没有则 None（调用方自行决定是否报错）。"""
    for path in candidates(name):
        try:
            if os.path.isfile(path):
                return path
        except Exception:
            continue
    return None


def missing_assets() -> list:
    """返回在源码树里缺失的必需资源名（供测试 / 构建前自检红灯）。"""
    out = []
    for name in REQUIRED_ASSETS:
        if find_asset(name) is None:
            out.append(name)
    return out


def asset_dir() -> str:
    """源码树里的 assets 目录（构建脚本据此拼 --include-data-files 的源路径）。"""
    return _asset_dir()
