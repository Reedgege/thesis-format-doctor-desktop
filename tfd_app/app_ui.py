# -*- coding: utf-8 -*-
"""论文格式医生 · 新版界面外壳层（窗口 / 事件 / 线程 / 引擎编排）

职责边界（施工单 §1）
--------------------
  视图层 tfd_app/ui2/screen.py  只按状态绘制 + 登记热区；不含业务
  本模块                       唯一的业务编排者：事件 → 引擎 → 状态 → 重绘
  引擎层 tfd_app/engine.py     现成，不改

两条必须守住的规则（照抄旧 UI，不许自己发明）
--------------------------------------------
1. **检查免费、修正才扣**：`run_check` / `md_to_docx`（报告导出）都不扣次；
   只有 `run_fix_headings` 前面才走 `_guard_license()` + 试用扣次（gui.py L1561-1594）。
2. **试用版输出带水印**：未激活时给修正结果加水印（仅作效果预览）。

线程模型：引擎全部在后台线程跑，通过 queue 回主线程改界面 —— 界面绝不卡死。
"""
from __future__ import annotations

import os
import queue
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox

from . import engine, license as lic, trial, watermark
from .assetpath import find_asset
from .ui2 import screen

# ---------------------------------------------------------------------------
# 品牌口径（与 gui.py 完全一致，不改口径）
# ---------------------------------------------------------------------------
WECHAT_NAME = "芦苇不熬夜"
WECHAT_ID = "reedskill"
KEFU_WECHAT = "reedgege"          # 客服微信（个人号；海外版不展示）
ABOUT_MAIL = "hi@reedskill.com"
PRODUCT_TITLE = {"tfd-student": "论文格式医生 · 学生版",
                 "tfd-mentor": "论文格式医生 · 导师版"}

SIZES_NORMAL = (1280, 720)
SIZES_MIN = (1100, 690)
NAV_QR = ("客服微信", "官方公众号")
DOCX_EXT = (".docx", ".doc", ".wps")


def edition_of() -> str:
    """自己是学生版还是导师版 —— 由产品代号推，不写死（两版共用一份代码）。"""
    return "advisor" if "mentor" in getattr(lic, "DEFAULT_PRODUCT", "") else "student"


class App:
    def __init__(self):
        self.edition = edition_of()
        self.label = {"student": "学生版", "advisor": "导师版"}[self.edition]
        self.is_mentor = (self.edition == "advisor")

        # ---- 业务状态 ----
        self.src = ""              # 已选论文
        self.batch = []            # 导师版：已导入清单
        self.template = ""         # 已选模板
        self.profile = ""          # 模板画像 json
        self.remembered = False    # 是否已记住模板
        self.report_md = ""        # 最近一次检查报告（markdown）
        self.mode_idx = 0          # 交付方式（导师版）
        self.busy = False
        self.q = queue.Queue()

        # ---- 窗口 ----
        self.root = tk.Tk()
        self.root.title(PRODUCT_TITLE.get(getattr(lic, "DEFAULT_PRODUCT", ""), "论文格式医生"))
        self.root.overrideredirect(True)
        self.root.configure(bg=screen.TH.get_theme(self.edition).bg)
        self._maxed = False
        self._normal_geo = None
        w, h = SIZES_NORMAL
        self._center(w, h)
        # 顺序要紧：先把 T 换成该版本的主题实例，再套配色覆盖。
        # 反过来写 = 覆盖被新实例冲掉，两版都会退回 theme.py 里的旧学术蓝。
        screen.T = screen.TH.get_theme(self.edition)
        screen.apply_style(self.edition)
        screen.EDITION = self.edition
        screen.TRUST = screen.TRUST_VARIANTS["rec"]      # 老板 2026-09-13 定稿
        self.canvas = tk.Canvas(self.root, highlightthickness=0, bd=0,
                                bg=screen.T.bg)
        self.canvas.pack(fill="both", expand=True)
        self._apply_icon()

        # ---- 二维码浮层 ----
        self._qr = None
        self._qr_pin = False
        self._qr_which = None
        self._qr_after = None

        # ---- 事件 ----
        self.canvas.bind("<Button-1>", self._on_click)
        self.canvas.bind("<Motion>", self._on_motion)
        self.canvas.bind("<Leave>", self._on_leave)
        self.root.bind("<Escape>", self._on_esc)
        self.root.bind("<Configure>", self._on_configure)

        self.refresh_license()
        self.root.update_idletasks()      # 先让窗口拿到真实尺寸，再画第一帧
        self.render()
        self._pump()

    # ==================================================================
    # 窗口
    # ==================================================================
    def _center(self, w, h):
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        x = max(0, (sw - w) // 2)
        y = max(0, (sh - h) // 3)
        self._want_w, self._want_h = w, h
        self.root.geometry("%dx%d+%d+%d" % (w, h, x, y))

    def _apply_icon(self):
        p = find_asset("icon.png")
        if not p:
            return
        try:
            self._icon_img = tk.PhotoImage(file=p)     # 必须留引用，否则被 GC
            self.root.iconphoto(True, self._icon_img)
        except Exception:
            pass                                        # 图标失败不影响使用

    def _on_configure(self, ev):
        if ev.widget is not self.canvas:
            return
        w = self.canvas.winfo_width()
        h = self.canvas.winfo_height()
        if (w, h) != getattr(self, "_last_size", None) and w > 200 and h > 200:
            self._last_size = (w, h)
            self.render()

    def _toggle_max(self):
        if self._maxed:
            w, h = self._normal_geo or SIZES_NORMAL
            sw = self.root.winfo_screenwidth()
            sh = self.root.winfo_screenheight()
            self.root.geometry("%dx%d+%d+%d" % (w, h, max(0, (sw - w) // 2), max(0, (sh - h) // 3)))
            self._maxed = False
        else:
            self._normal_geo = (self.root.winfo_width(), self.root.winfo_height())
            # 只铺到工作区（不去动任务栏）
            sw = self.root.winfo_screenwidth()
            sh = self.root.winfo_screenheight()
            self.root.geometry("%dx%d+0+0" % (sw, int(sh * 0.94)))
            self._maxed = True

    def _minimize(self):
        # 无边框窗口的 iconify：先恢复系统边框，最小化，再在重新映射时抹掉边框
        self.root.overrideredirect(False)
        self.root.iconify()

        def _restore(_ev=None):
            self.root.overrideredirect(True)
            self.root.unbind("<Map>")
        self.root.bind("<Map>", _restore)

    def _close(self):
        if self.busy:
            if not messagebox.askyesno("正在处理", "还有任务在进行中，确定要关闭吗？"):
                return
        try:
            engine.cleanup_conv_dirs()
        except Exception:
            pass
        self.root.destroy()

    # ==================================================================
    # 渲染 + 事件派发
    # ==================================================================
    def render(self):
        # 首帧还没映射时 winfo_width() 只有 1px —— 用「想要的尺寸」兜底，
        # 否则第一帧会按 300x200 画一遍（闪一下错的画面）
        w = self.root.winfo_width()
        h = self.root.winfo_height()
        if w <= 200:
            w = self._want_w
        if h <= 200:
            h = self._want_h
        screen.STATUS = self._status_text()
        screen.STEP_INDEX = self.step_index
        screen.MAIN_BTN_TEXT = self.main_btn_text
        screen.MAIN_BTN_ENABLED = not self.busy
        screen.PICKED = self.src
        screen.BATCH_COUNT = len(self.batch)
        screen.TEMPLATE_NAME = self.template
        screen.TPL_REMEMBERED = self.remembered
        screen.render(self.canvas, w, h, self.label)

    @property
    def step_index(self):
        if self.report_md:
            return 1
        if self.src or self.batch:
            return 0
        return 0

    @property
    def main_btn_text(self):
        if self.busy:
            return "正在检查…"
        if self.report_md:
            return "查看检查报告"
        if self.is_mentor:
            return "开始检查"
        return "开始检查"

    def _status_text(self):
        if self.busy:
            return "处理中…"
        if self.report_md:
            return "检查完成"
        return "就绪"

    def _on_click(self, ev):
        self._hide_qr(force=True)
        r = screen.hit(ev.x, ev.y)
        if not r:
            return
        act, pl = r
        if act == "win":
            {"min": self._minimize, "max": self._toggle_max,
             "close": self._close}[pl]()
        elif act == "nav":
            if pl in NAV_QR:
                self._show_qr(pl, pinned=True)
            else:
                self._nav_other(pl)
        elif act == "license":
            self._open_activate()
        elif act == "mode":
            self.mode_idx = int(pl)
            self.render()
        elif act == "signer":
            self._edit_signer(ev)
        elif act.startswith("btn:"):
            self._on_button(act.split(":", 1)[1])
        elif act.startswith("zone:"):
            self._on_zone(act.split(":", 1)[1])

    def _on_button(self, text):
        if text in ("选择文件",):
            self._pick_files(multi=False)
        elif text == "选择文件夹":
            self._pick_folder()
        elif text == "选择模板":
            self._pick_template()
        elif text == "记住此模板":
            self._toggle_remember()
        elif text == "开始检查":
            if self.report_md and not self.busy:
                self._show_report_window()
            else:
                self._do_check()
        elif text == "一键修正":
            self._do_fix()

    def _on_zone(self, main):
        if main.startswith("批量导入论文"):
            self._pick_files(multi=True)
        elif "模板" in main:
            self._pick_template()
        else:
            self._pick_files(multi=False)

    # ==================================================================
    # 二维码浮层（悬停出、移开收、点击钉住）
    # ==================================================================
    def _on_motion(self, ev):
        r = screen.hit(ev.x, ev.y)
        which = r[1] if (r and r[0] == "nav" and r[1] in NAV_QR) else None
        if which == self._qr_which:
            return
        self._qr_which = which
        if self._qr_after:
            self.root.after_cancel(self._qr_after)
            self._qr_after = None
        if which:
            self._qr_after = self.root.after(260, lambda: self._show_qr(which))
        elif not self._qr_pin:
            self._hide_qr()

    def _on_leave(self, _ev=None):
        if not self._qr_pin:
            self._hide_qr()

    def _on_esc(self, _ev=None):
        if self._qr:
            self._qr_pin = False
            self._hide_qr()
        else:
            self._close()

    def _qr_anchor(self, which):
        for x0, y0, x1, y1, act, pl in screen.HOT:
            if act == "nav" and pl == which:
                return x0, y0, x1, y1
        return None

    def _show_qr(self, which, pinned=False):
        box = self._qr_anchor(which)
        if not box:
            return
        if pinned:
            self._qr_pin = True
        self._hide_qr(keep_pin=pinned or self._qr_pin)

        PW, PH = 268, 306
        top = tk.Toplevel(self.root)
        top.overrideredirect(True)
        top.configure(bg=screen.T.surface)
        cv = tk.Canvas(top, width=PW, height=PH, highlightthickness=1, bd=0,
                       bg=screen.T.surface, highlightbackground=screen.T.secondary_border)
        cv.pack()

        asset_name = ("kf_qrcode.png" if which == "客服微信" else "gz_qrcode.png")
        img_path = find_asset(asset_name)
        cv.create_text(PW / 2.0, 24, text=("客服微信" if which == "客服微信" else "官方公众号"),
                       fill=screen.T.navy, font=screen.T.sans(screen.TYPE["section_title"], bold=True))
        if img_path:
            try:
                self._qr_img = tk.PhotoImage(file=img_path)   # 留引用
                k = max(1, int(max(self._qr_img.width(), self._qr_img.height()) / 196) + 1)
                img = self._qr_img.subsample(k, k) if k > 1 else self._qr_img
                self._qr_img_shown = img
                cv.create_image(PW / 2.0, 140, image=img)
                caption = ("微信：%s" % KEFU_WECHAT) if which == "客服微信" \
                    else ("公众号：%s（ID：%s）" % (WECHAT_NAME, WECHAT_ID))
            except Exception as e:
                caption = "二维码读取失败：%s" % e
        else:
            # 不静默降级：明确告知缺图，并给出可核验的文字信息
            cv.create_rectangle(46, 60, PW - 46, 220, outline=screen.T.secondary_border,
                                dash=(4, 3))
            cv.create_text(PW / 2.0, 128, text="二维码待配置", fill=screen.T.muted,
                           font=screen.T.sans(screen.TYPE["body"], bold=True))
            cv.create_text(PW / 2.0, 154, text="（放一张 PNG 到 tfd_app/assets/%s）" % asset_name,
                           fill=screen.T.muted, font=screen.T.sans(screen.TYPE["caption"]))
            caption = "公众号：%s（ID：%s）" % (WECHAT_NAME, WECHAT_ID)
        cv.create_text(PW / 2.0, 254, text=caption, fill=screen.T.navy,
                       font=screen.T.sans(screen.TYPE["caption"]))
        cv.create_text(PW / 2.0, 278, text="点一下钉住，Esc 收起", fill=screen.T.muted,
                       font=screen.T.sans(screen.TYPE["footer"]))

        x = int(box[2] - PW)
        y = int(box[3] + 6)
        sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        wx = self.root.winfo_rootx() + x
        wy = self.root.winfo_rooty() + y
        wx = max(0, min(wx, sw - PW))
        if wy + PH > sh:
            wy = max(0, self.root.winfo_rooty() + int(box[1]) - PH - 6)
        top.geometry("%dx%d+%d+%d" % (PW, PH, wx, wy))
        self._qr = top

    def _hide_qr(self, force=False, keep_pin=False):
        if force:
            self._qr_pin = False
        if not keep_pin and not force and self._qr_pin:
            return
        if self._qr is not None:
            try:
                self._qr.destroy()
            except Exception:
                pass
            self._qr = None
        if force:
            self._qr_which = None

    def _nav_other(self, which):
        if which == "关于":
            messagebox.showinfo("关于", (
                "%s\n\n版本：新界面（施工中）\n"
                "官网：reedskill.com\n公众号：%s（ID：%s）\n合作邮箱：%s"
                % (PRODUCT_TITLE.get(getattr(lic, "DEFAULT_PRODUCT", ""), "论文格式医生"),
                   WECHAT_NAME, WECHAT_ID, ABOUT_MAIL)))
        else:
            messagebox.showinfo("帮助", (
                "1. 选择论文（.docx；旧格式 .doc/.wps 会自动转换）\n"
                "2. 可选：选择学校模板，按模板要求检查\n"
                "3. 点「开始检查」看报告；交付方式选好后点主按钮处理\n\n"
                "遇到问题：关注公众号【%s】留言，或致信 %s。"
                % (WECHAT_NAME, ABOUT_MAIL)))

    # ==================================================================
    # 文件 / 模板
    # ==================================================================
    def _pick_files(self, multi):
        if multi:
            paths = filedialog.askopenfilenames(
                title="选择论文（可多选）",
                filetypes=[("Word 文档", "*.docx *.doc *.wps"), ("所有文件", "*.*")])
            if paths:
                self.batch = [p for p in paths if p.lower().endswith(DOCX_EXT)]
                if not self.batch:
                    messagebox.showwarning("没选中 Word 文档", "请选择 .docx / .doc / .wps 文件。")
                self.src = self.batch[0] if self.batch else ""
        else:
            p = filedialog.askopenfilename(
                title="选择论文",
                filetypes=[("Word 文档", "*.docx *.doc *.wps"), ("所有文件", "*.*")])
            if p:
                self.src = p
                self.batch = []
        self.render()

    def _pick_folder(self):
        d = filedialog.askdirectory(title="选择放着论文的文件夹")
        if not d:
            return
        got = []
        for name in sorted(os.listdir(d)):
            if name.lower().endswith(DOCX_EXT) and not name.startswith("~$"):
                got.append(os.path.join(d, name))
        if not got:
            messagebox.showwarning("文件夹里没有 Word 文档",
                                   "该文件夹下没有 .docx / .doc / .wps 文件。")
            return
        self.batch = got
        self.src = got[0]
        self.render()

    def _pick_template(self):
        p = filedialog.askopenfilename(title="选择学校模板",
                                       filetypes=[("Word 文档", "*.docx *.doc"), ("所有文件", "*.*")])
        if not p:
            return
        self.template = p
        self.render()

        def work():
            out = os.path.join(self._cache_dir(), "profile.json")
            txt = engine.run_build_profile(p, out)
            return out, txt

        self._submit("正在提取模板要求…", work, self._after_profile)

    def _after_profile(self, res):
        if not res:
            return
        out, _txt = res
        self.profile = out
        self.render()

    def _toggle_remember(self):
        if not self.template:
            messagebox.showinfo("先选模板", "请先选择学校模板，再点「记住此模板」。")
            return
        self.remembered = not self.remembered
        if self.remembered:
            self._save_pref({"template": self.template})
        self.render()

    def _cache_dir(self):
        d = os.path.join(os.path.expanduser("~"), ".thesis_format_doctor")
        os.makedirs(d, exist_ok=True)
        return d

    def _pref_path(self):
        return os.path.join(self._cache_dir(), "ui_pref.json")

    def _save_pref(self, kv):
        import json
        p = self._pref_path()
        data = {}
        if os.path.isfile(p):
            try:
                data = json.load(open(p, encoding="utf-8"))
            except Exception:
                data = {}
        data.update(kv)
        try:
            json.dump(data, open(p, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
        except Exception:
            pass

    def _edit_signer(self, ev):
        if getattr(self, "_signer_win", None):
            return
        e = tk.Entry(self.root, bd=0, relief="flat", bg=screen.T.surface,
                     fg=screen.T.navy, font=screen.T.sans(screen.TYPE["body"]))
        e.place(x=ev.x, y=ev.y - 8, width=240, height=26)
        e.insert(0, getattr(self, "signer", "论文格式医生·导师版"))
        e.focus_set()

        def commit(_ev=None):
            self.signer = e.get().strip() or "论文格式医生·导师版"
            self._save_pref({"signer": self.signer})
            e.destroy()
            self._signer_win = None
            self.render()
        e.bind("<Return>", commit)
        e.bind("<FocusOut>", commit)
        self._signer_win = e

    # ==================================================================
    # 引擎：检查 / 修正（线程 + queue，界面不卡）
    # ==================================================================
    def _submit(self, status, work, done):
        """后台跑引擎；完成后回调 done(结果)。异常也要回主线程显示，不许静默。"""
        if self.busy:
            messagebox.showinfo("请稍等", "还有任务在处理中。")
            return
        self.busy = True
        self._busy_text = status
        self.render()

        def runner():
            try:
                res = work()
                self.q.put(("ok", done, res))
            except Exception as e:
                import traceback
                self.q.put(("err", done, (e, traceback.format_exc())))
        threading.Thread(target=runner, daemon=True).start()

    def _pump(self):
        try:
            while True:
                kind, done, payload = self.q.get_nowait()
                self.busy = False
                if kind == "ok":
                    try:
                        done(payload)
                    except Exception as e:
                        self._show_error(e)
                else:
                    exc, tb = payload
                    self._show_error(exc, tb)
                self.render()
        except queue.Empty:
            pass
        self.root.after(80, self._pump)

    def _show_error(self, exc, tb=None):
        sys.stderr.write("[app_ui] 处理失败: %s\n%s\n" % (exc, tb or ""))
        messagebox.showerror("处理失败", "%s\n\n%s" % (type(exc).__name__, exc))

    def _do_check(self):
        if not self.src:
            messagebox.showinfo("还没选论文", "请先选择要检查的论文文件。")
            return
        src, profile = self.src, self.profile or None

        def work():
            path, note = engine.normalize_input(src)
            md = engine.run_check(path, profile_path=profile)
            return path, note, md

        self._submit("正在检查格式…", work, self._after_check)

    def _after_check(self, res):
        path, note, md = res
        self.report_md = md or "（引擎没有返回报告内容）"
        self._note = note
        self.render()
        self._show_report_window()

    def _do_fix(self):
        if not self.src:
            messagebox.showinfo("还没选论文", "请先选择要修正的论文文件。")
            return
        src, profile = self.src, self.profile or None
        add_comments = self.is_mentor and self.mode_idx in (0, 2)   # 只批注 / 两者都要
        do_fix = (not self.is_mentor) or self.mode_idx in (1, 2)    # 一键修正 / 两者都要

        # 付费守卫 + 试用扣次（照抄 gui.py L1561-1594，规则一致）
        if not self._guard_license():
            return
        licensed = trial.is_licensed()
        if not licensed:
            left = trial.trials_left()
            if left <= 0:
                self._trial_exhausted()
                return
            if not messagebox.askyesno("试用提示", (
                    "试用版可修正 %d 次，本次会用掉 1 次。\n\n"
                    "修正结果会带水印、且为只读文档（仅作效果预览）；\n"
                    "激活正式版后输出可编辑、无水印的文件。\n\n继续吗？" % left)):
                return
            if not trial.consume_trial():
                self._trial_exhausted()
                return

        dst = filedialog.asksaveasfilename(
            title="保存修正后的论文", defaultextension=".docx",
            initialfile=os.path.splitext(os.path.basename(src))[0] + "_已修正.docx",
            filetypes=[("Word 文档", "*.docx")])
        if not dst:
            return
        base = os.path.splitext(dst)[0]

        def work():
            path, _note = engine.normalize_input(src)
            rep = engine.run_fix_headings(path, dst, profile_path=profile,
                                          report_docx=base + "_修改报告.docx",
                                          add_comments=add_comments or do_fix)
            note = ""
            if not licensed:
                try:
                    if watermark.apply_watermark(dst):
                        note = ("\n\n> 本预览版带水印且为只读文档（仅作效果预览）；"
                                "激活后输出可编辑无水印正式版。\n")
                except Exception as e:
                    note = "\n\n> 水印添加失败：%s（请反馈）\n" % e
            return dst, rep, note

        self._submit("正在修正…", work, self._after_fix)

    def _after_fix(self, res):
        dst, rep, note = res
        self.report_md = (rep or "") + note
        self.render()
        self._show_report_window()
        messagebox.showinfo("已生成", "修正完成：\n%s\n\n修改明细：\n%s"
                            % (os.path.basename(dst), os.path.basename(dst)[:-5] + "_修改报告.docx"))

    def _guard_license(self):
        """付费操作前守卫（照抄 gui.py）：本地失效立即拦；其余交给引擎端校验。"""
        try:
            if lic.is_revoked():
                messagebox.showerror("授权已失效", "当前授权已被撤销或到期，请重新激活。")
                return False
            st = lic.consume_use()
            if st in ("revoked", "expired"):
                messagebox.showerror("授权已失效", "授权状态：%s，请重新激活。" % st)
                return False
        except Exception as e:
            sys.stderr.write("[app_ui] 授权校验异常（离线放行）: %s\n" % e)
        return True

    def _trial_exhausted(self):
        messagebox.showinfo("试用次数已用完", (
            "免费试用次数已用完。\n\n"
            "点右上角「升级正式版 →」查看激活方式；\n"
            "或关注公众号【%s】回复「试用」获取说明。" % WECHAT_NAME))

    # ==================================================================
    # 报告窗
    # ==================================================================
    def _show_report_window(self):
        if getattr(self, "_rep_win", None) and self._rep_win.winfo_exists():
            self._rep_win.lift()
            return
        win = tk.Toplevel(self.root)
        self._rep_win = win
        win.title("检查报告")
        win.configure(bg=screen.T.bg)
        win.geometry("820x620")
        head = tk.Frame(win, bg=screen.T.surface, height=54)
        head.pack(fill="x")
        tk.Label(head, text="检查报告", bg=screen.T.surface, fg=screen.T.navy,
                 font=screen.T.serif(screen.TYPE["page_title"], bold=True)).pack(side="left", padx=18, pady=10)
        body = tk.Text(win, wrap="word", bd=0, bg=screen.T.bg, fg=screen.T.navy,
                       font=screen.T.sans(screen.TYPE["body"]), padx=18, pady=12)
        body.pack(fill="both", expand=True)
        body.insert("1.0", self.report_md)
        body.configure(state="disabled")
        foot = tk.Frame(win, bg=screen.T.surface)
        foot.pack(fill="x")
        tk.Button(foot, text="导出 Word 报告", command=self._export_report,
                  bg=screen.T.primary, fg="#FFFFFF", bd=0, padx=16, pady=6).pack(side="right", padx=12, pady=10)
        tk.Button(foot, text="关闭", command=win.destroy, bg=screen.T.surface,
                  fg=screen.T.muted, bd=0, padx=16, pady=6).pack(side="right", pady=10)

    def _export_report(self):
        if not self.report_md:
            return
        base = os.path.splitext(os.path.basename(self.src or "论文"))[0]
        dst = filedialog.asksaveasfilename(title="导出检查报告", defaultextension=".docx",
                                           initialfile=base + "_检查报告.docx",
                                           filetypes=[("Word 文档", "*.docx")])
        if not dst:
            return
        try:
            engine.md_to_docx(self.report_md, dst)
        except Exception as e:
            self._show_error(e)
            return
        messagebox.showinfo("已导出", os.path.basename(dst))

    # ==================================================================
    # 授权
    # ==================================================================
    def refresh_license(self):
        """把真实授权状态映射到界面四态（文案取自视图层 LIC_STATES，不自编）。

        license_summary() 实测返回 dict：{'kind':'永久版','desc':...,'badge':...,'type':'lifetime'}
        → 用 type 字段判断，不猜字符串。
        """
        tag = "trial_new"
        try:
            if trial.is_licensed():
                info = {}
                try:
                    info = lic.license_summary()
                except Exception:
                    info = {}
                if isinstance(info, dict):
                    t = str(info.get("type") or "")
                    kind = str(info.get("kind") or "")
                else:
                    t, kind = "", str(info)
                if t == "lifetime" or "永久" in kind:
                    tag = "permanent"
                elif t or "月" in kind:
                    tag = "monthly"
                else:
                    tag = "permanent"
            else:
                tag = "trial_new" if trial.trials_left() > 0 else "trial_used"
        except Exception as e:
            sys.stderr.write("[app_ui] 读授权状态失败: %s\n" % e)
        screen.LIC_TAG = tag
        screen.LIC = screen.LIC_STATES[tag]

    def _open_activate(self):
        win = tk.Toplevel(self.root)
        win.title("激活正式版")
        win.configure(bg=screen.T.bg)
        win.geometry("460x300")
        tk.Label(win, text="激活正式版", bg=screen.T.bg, fg=screen.T.navy,
                 font=screen.T.sans(screen.TYPE["section_title"], bold=True)).pack(pady=(16, 4))
        try:
            mc = lic.get_machine_code()
        except Exception as e:
            mc = "读取失败：%s" % e
        tk.Label(win, text="机器码：%s" % mc, bg=screen.T.bg, fg=screen.T.muted,
                 font=screen.T.sans(screen.TYPE["caption"])).pack()
        ent = tk.Entry(win, width=44, bd=1, relief="solid")
        ent.pack(pady=12)
        tip = tk.Label(win, text="粘贴激活码后点「激活」", bg=screen.T.bg, fg=screen.T.muted,
                       font=screen.T.sans(screen.TYPE["caption"]))
        tip.pack()

        def do():
            card = ent.get().strip()
            if not card:
                tip.config(text="请先填写激活码")
                return
            tip.config(text="正在校验…")

            def work():
                # 离线码优先（无需联网），否则走在线激活
                if "|" in card:
                    ok, msg = lic.verify_offline_code(card, mc)
                    if ok:
                        lic.save_offline_license(card, mc)
                    return bool(ok), str(msg or ("已激活（离线码）" if ok else "离线码校验失败"))
                info = lic.activate_online(card, mc)
                return bool(info), str(info)

            def runner():
                try:
                    res = work()
                except Exception as e:
                    res = (False, "%s: %s" % (type(e).__name__, e))
                try:
                    win.after(0, lambda: self._after_activate(win, tip, res))
                except Exception:
                    pass
            threading.Thread(target=runner, daemon=True).start()
        tk.Button(win, text="激活", command=do, bg=screen.T.primary, fg="#FFFFFF",
                  bd=0, padx=18, pady=6).pack(pady=8)

    def _after_activate(self, win, tip, res):
        ok, msg = res
        tip.config(text=("激活成功 ✓" if ok else "未通过：%s" % msg))
        if ok:
            try:
                lic.start_heartbeat(on_revoked=lambda: self.root.after(0, self._on_revoked))
            except Exception:
                pass
            self.refresh_license()
            self.render()
            win.after(900, win.destroy)

    def _on_revoked(self):
        self.refresh_license()
        self.render()
        messagebox.showwarning("授权状态变化", "当前授权已被撤销或到期，请重新激活。")

    # ==================================================================
    def run(self):
        self.root.mainloop()


def main():
    try:
        App().run()
    except Exception:
        import traceback
        traceback.print_exc()
        raise


if __name__ == "__main__":
    main()
