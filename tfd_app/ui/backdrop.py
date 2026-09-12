# -*- coding: utf-8 -*-
"""
背景层 + 宣纸“透明”容器
=========================

学生版窗口的**唯一背景层** = 装修包正式背景图 ``background/STUDENT_BACKGROUND.png``
（仓库内文件名 ``tfd_app/assets/student_background.png``）。旧实现「用宣纸纯色铺满
窗口当底」的做法已整块删除：纯色只作为**背景图之下的兜底色**，正常渲染永远看不见。

两个角色，解决“底图被上层不透明 Frame 整块盖住、宣纸纹理看不见”的问题：

* ``Backdrop``      —— 把这张背景图按 cover 铺满窗口：用整数 zoom/subsample 组合近似
                       任意比例（Tk 8.6 原生 PhotoImage 只支持整数倍缩放，不引第三方
                       依赖），居中裁切，并按窗口尺寸缓存。
* ``TextureCanvas`` —— 内容容器画布：把背景图**在自己位置的那一块**画到画布上。
                       Tk 没有透明通道，所以“透明容器”只能各自画自己那块底图；
                       因为都取自同一张缩放图、同一套偏移，拼起来仍是连续的一张底图，
                       卡片/按钮之间的空隙就会透出宣纸纹理（无缝、无残影）。

禁止使用完整 UI 预览图作背景（AGENT_MASTER_INSTRUCTION.md）。
"""
from __future__ import annotations

import math
import os
import sys
import time
import tkinter as tk

from . import theme
from .theme import get_theme, TYPE

from .. import assetpath

# 学生版窗口唯一背景层（装修包 background/STUDENT_BACKGROUND.png）
BACKDROP_ASSET = "student_background.png"

_TEX_TAG = "tfd_tex"        # 底图片段统一用这个 tag，重绘时只删它，不碰前景内容
_MAX_FACTOR = 4             # zoom 上限：zoom(4) 中间图约 16 倍源图，本机实测 0.13s
_UNDER = 0.97               # 允许最多 3% 的“差一点没铺满”，换取更贴近的比例

_ERR_TAG = "tfd_tex_err"    # 背景图缺失时的醒目报错层
_ERR_TEXT = "背景图缺失：%s（请重新完整安装 / 解压）" % BACKDROP_ASSET


def _backdrop_path():
    """背景图真实路径；读不到返回 None（调用方必须**明显报错**，不许静默铺纯色）。"""
    return assetpath.find_asset(BACKDROP_ASSET)


def _log_missing(reason):
    """背景图缺失：写日志 + 控制台（打包版无控制台，故同时落盘）。"""
    line = "[%s] backdrop: %s\n" % (time.strftime("%Y-%m-%d %H:%M:%S"), reason)
    try:
        sys.stderr.write(line)
        sys.stderr.flush()
    except Exception:
        pass
    try:
        import tempfile
        path = os.path.join(tempfile.gettempdir(), "tfd_backdrop_error.log")
        with open(path, "a", encoding="utf-8") as fh:
            fh.write(line)
    except Exception:
        pass


def best_ratio(cover_ratio):
    """cover 需要的缩放比例 → 最接近的 (zoom, subsample) 整数组合。

    Tk 的 PhotoImage 只能整数放大 / 整数抽样，任意比例要用 zoom（方块放大）
    + subsample（等距抽样）组合近似。这里在 1.._MAX_FACTOR 内挑误差最小的组合，
    并要求结果 >= cover_ratio * _UNDER：宁可差一点没铺满，也不要过度放大——
    旧实现一律 zoom(2) 把源图放大 2 倍再居中裁切，结果两边的水墨装饰全被裁掉，
    屏幕上只剩中间那块纯色，看起来就跟“底图没渲染”一样。
    """
    best = None
    for z in range(1, _MAX_FACTOR + 1):
        for s in range(1, _MAX_FACTOR + 1):
            r = float(z) / float(s)
            if r < cover_ratio * _UNDER:
                continue
            cand = (abs(r - cover_ratio), z, s)
            if best is None or cand < best:
                best = cand
    if best is None:
        return max(1, int(math.ceil(cover_ratio))), 1
    return best[1], best[2]


class Backdrop:
    """窗口底图：``attach(root)`` 后铺满整个窗口，并作为内容容器的底图来源。"""

    _registry = {}          # str(toplevel) -> Backdrop（供 TextureCanvas 反查）

    def __init__(self, edition="student"):
        self._t = get_theme(edition)
        self._edition = edition
        # 学生版只有一张背景图（装修包 STUDENT_BACKGROUND.png）；不再按 edition 二选一。
        self._path = _backdrop_path()
        self._src = None                  # 原图（首次用到时才加载）
        self._src_failed = False
        self._warned = False              # 缺件只报错一次，不刷屏
        self._fit_cache = None            # (缩放图, left, top)，按窗口尺寸缓存
        self._fit_key = None
        self._sweep_job = None            # 拖拽收尾后的整体重绘
        self._rebuild_job = None          # 缩放图精确重算
        self._hosts = []                  # 已注册的 TextureCanvas
        self._root = None
        self.canvas = None

    # ------------------------------------------------------------ 挂载 / 注销
    def attach(self, root):
        """在 root 最底层铺底图。须在内容控件创建前调用，底图才在最底层。"""
        self._root = root
        Backdrop._registry[str(root)] = self
        self.canvas = tk.Canvas(root, highlightthickness=0, bd=0, bg=self._t.bg)
        self.canvas.place(x=0, y=0, relwidth=1, relheight=1)
        self.canvas.tk.call("lower", self.canvas._w)
        root.configure(bg=self._t.bg)
        self.canvas.bind("<Configure>", lambda e: self.redraw())
        try:
            root.bind("<Configure>", lambda e: self._sweep_later(), add="+")
        except Exception:
            pass
        # 缺件不许静默：立刻写日志（画布上的醒目提示在 paint_into 里画）
        if self._path is None:
            self._warn_missing("找不到背景图资源 %s（REQUIRED_ASSETS 缺件）" % BACKDROP_ASSET)
        self.redraw()

    def _warn_missing(self, reason):
        if self._warned:
            return
        self._warned = True
        _log_missing(reason)

    @classmethod
    def for_widget(cls, widget):
        """按控件反查它所在窗口的底图（找不到返回 None，调用方需容错）。"""
        try:
            return cls._registry.get(str(widget.winfo_toplevel()))
        except Exception:
            return None

    def register(self, host):
        if host not in self._hosts:
            self._hosts.append(host)

    def unregister(self, host):
        try:
            self._hosts.remove(host)
        except ValueError:
            pass

    # --------------------------------------------------------------- 缩放 / 切片
    def _source(self):
        if self._src is None and not self._src_failed:
            if self._path is None:
                self._src_failed = True
                self._warn_missing("找不到背景图资源 %s（REQUIRED_ASSETS 缺件）"
                                   % BACKDROP_ASSET)
                return None
            try:
                self._src = tk.PhotoImage(file=self._path)
            except Exception as exc:
                self._src_failed = True
                self._src = None
                self._warn_missing("背景图加载失败：%s（%s）" % (self._path, exc))
        return self._src

    def _window_size(self):
        if self._root is None:
            return None
        try:
            w = self._root.winfo_width()
            h = self._root.winfo_height()
        except Exception:
            return None
        if w < 4 or h < 4:
            return None
        return w, h

    def _fit(self):
        """返回 (铺满窗口的缩放图, left, top)；尺寸变化时走重建节流。"""
        size = self._window_size()
        if size is None:
            return None
        if self._fit_cache is not None and self._fit_key == size:
            return self._fit_cache
        if self._fit_cache is not None:
            # 连续拖拽中：先用上一张，等定时器收尾后再精确重算，避免每帧重算卡顿
            self._rebuild_later()
            return self._fit_cache
        src = self._source()
        if src is None:
            return None
        iw, ih = src.width(), src.height()
        if iw < 2 or ih < 2:
            return None
        w, h = size
        z, s = best_ratio(max(w / float(iw), h / float(ih)))
        try:
            img = src.zoom(z, z).subsample(s, s)
        except Exception:
            img = src
        left = (img.width() - w) // 2
        top = (img.height() - h) // 2
        self._fit_cache = (img, left, top)
        self._fit_key = size
        return self._fit_cache

    def _origin(self, widget):
        """控件相对所在窗口左上角的偏移（像素）。"""
        try:
            top = widget.winfo_toplevel()
            return (widget.winfo_rootx() - top.winfo_rootx(),
                    widget.winfo_rooty() - top.winfo_rooty())
        except Exception:
            return (0, 0)

    def paint_into(self, canvas):
        """把底图在本窗口的对应片段画到 canvas 上（子控件与前景画布项照常显示其上）。"""
        try:
            if not canvas.winfo_exists():
                return
        except Exception:
            return
        w = canvas.winfo_width()
        h = canvas.winfo_height()
        if w < 2 or h < 2:
            return
        canvas.delete(_TEX_TAG)
        canvas.delete(_ERR_TAG)
        fit = self._fit()
        if fit is None:
            # 兜底色只允许垫在背景图之下；这里同时把「缺件」明确画出来（不静默）
            canvas.create_rectangle(0, 0, w, h, fill=self._t.bg,
                                    outline=self._t.bg, tags=_TEX_TAG)
            canvas.tag_lower(_TEX_TAG)
            canvas.create_text(w / 2, max(18, h / 2), text=_ERR_TEXT,
                               fill=self._t.error_text,
                               font=self._t.sans(TYPE["section_title"], bold=True),
                               anchor="center", tags=_ERR_TAG)
            return
        img, left, top = fit
        dx, dy = self._origin(canvas)
        canvas.create_image(-left - dx, -top - dy, anchor="nw",
                            image=img, tags=_TEX_TAG)
        canvas.tag_lower(_TEX_TAG)      # 底图必须在所有前景画布项之下

    # ------------------------------------------------------------------ 重绘
    def redraw(self):
        if self.canvas is None:
            return
        self.paint_into(self.canvas)

    def _repaint_hosts(self):
        for host in list(self._hosts):
            try:
                if host.winfo_exists():
                    self.paint_into(host)
                else:
                    self.unregister(host)
            except Exception:
                self.unregister(host)

    def _sweep_later(self):
        """窗口尺寸变化后统一重绘一次（各容器的相对位置可能都变了）。"""
        if self._sweep_job is not None or self._root is None:
            return
        try:
            self._sweep_job = self._root.after(60, self._sweep)
        except Exception:
            self._sweep_job = None

    def _sweep(self):
        self._sweep_job = None
        self.redraw()
        self._repaint_hosts()

    def _rebuild_later(self):
        if self._rebuild_job is not None or self._root is None:
            return
        try:
            self._rebuild_job = self._root.after(140, self._rebuild_now)
        except Exception:
            self._rebuild_job = None

    def _rebuild_now(self):
        self._rebuild_job = None
        self._fit_cache = None          # 强制按当前尺寸重新缩放
        self.redraw()
        self._repaint_hosts()


class TextureCanvas(tk.Canvas):
    """宣纸“透明”内容容器：底图画的是整张背景图在自己位置的那一块。

    用法与普通 Frame 一致（可 pack/grid、可放子控件）；子控件画在底图之上，
    没有子控件的空隙就透出宣纸纹理。子类可在 ``draw_foreground()`` 里用画布项
    画自己的文字/线条——画布上的文字没有底色块，视觉效果等同透明背景。
    """

    def __init__(self, master, edition="student", **kw):
        t = get_theme(edition)
        kw.setdefault("highlightthickness", 0)
        kw.setdefault("bd", 0)
        kw.setdefault("bg", t.bg)
        super().__init__(master, **kw)
        self._t = t
        self.bind("<Configure>", self._on_configure, add="+")
        self.bind("<Destroy>", self._on_destroy, add="+")

    def _on_configure(self, _event=None):
        self.repaint_texture()
        self.draw_foreground()

    def _on_destroy(self, _event=None):
        bd = Backdrop.for_widget(self)
        if bd is not None:
            bd.unregister(self)

    def repaint_texture(self):
        bd = Backdrop.for_widget(self)
        if bd is None:
            return
        bd.register(self)
        bd.paint_into(self)

    def draw_foreground(self):
        """子类钩子：在底图之上画自己的画布项（默认什么都不画）。"""
        return
