# -*- coding: utf-8 -*-
"""学生版滚轮回归（2026-11-17 跨平台同步检查）。

学生版操作界面不受影响，但「手动录入」与「确认页」两个弹窗虽然绑了
Button-4/5，处理函数只读 event.delta —— X11 下这两个事件没有 delta，
旧写法会把「向上」也滚成「向下」。现统一走 _wheel_steps_event。

用法：python tests/test_wheel_cross_platform.py（需要带 tkinter 的系统 Python）
"""
from __future__ import annotations

import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, ".."))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

import tkinter as tk                                            # noqa: E402
from tfd_app.gui import WHEEL_UNITS, _wheel_steps_event         # noqa: E402

FAILED = []


def ck(label, cond, extra=""):
    print(("  PASS " if cond else "  FAIL ") + label + ("  | " + extra if extra else ""))
    if not cond:
        FAILED.append(label)


class E:
    def __init__(self, delta=0, num=None):
        self.delta = delta
        self.num = num


print("=== 1. 事件归一化 ===")
ck("X11 Button-4 向上", _wheel_steps_event(E(num=4)) == -WHEEL_UNITS)
ck("X11 Button-5 向下", _wheel_steps_event(E(num=5)) == WHEEL_UNITS)
ck("Windows 上滚 delta=+120", _wheel_steps_event(E(delta=120)) == -WHEEL_UNITS)
ck("Windows 下滚 delta=-120", _wheel_steps_event(E(delta=-120)) == WHEEL_UNITS)
ck("macOS 上滚 delta=+1 也按一格算", _wheel_steps_event(E(delta=1)) == -WHEEL_UNITS)
ck("非滚轮事件 → 0", _wheel_steps_event(E()) == 0)

print("\n=== 2. 源码层面：两处弹窗都绑了三种事件 ===")
src = open(os.path.join(ROOT, "tfd_app", "gui.py"), encoding="utf-8").read()
for seq in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
    ck("%s 有绑定" % seq, seq in src)

print("\n=== 3. 真 Tk 事件打到 Canvas ===")
root = tk.Tk()
root.geometry("300x220+0+0")
root.update()
try:
    cv = tk.Canvas(root, width=200, height=100, highlightthickness=0,
                   yscrollincrement=20)
    cv.pack()
    cv.configure(scrollregion=(0, 0, 200, 1000))
    cv.update_idletasks()

    def handler(e):
        steps = _wheel_steps_event(e)
        if steps:
            cv.yview_scroll(steps, "units")
        return "break"

    cv.bind("<MouseWheel>", handler)
    cv.bind("<Button-4>", handler)
    cv.bind("<Button-5>", handler)

    def fire(seq, **kw):
        root.update()
        cv.event_generate(seq, x=50, y=50, when="now", **kw)
        root.update()
        return cv.yview()[0]

    cv.yview_moveto(0.5)
    base = cv.yview()[0]
    ck("Button-5 真的向下滚", fire("<Button-5>") > base)
    cv.yview_moveto(0.5)
    ck("Button-4 真的向上滚", fire("<Button-4>") < base)
    cv.yview_moveto(0.5)
    ck("MouseWheel delta=-120 向下滚", fire("<MouseWheel>", delta=-120) > base)
finally:
    root.destroy()

print("\n" + "=" * 66)
if FAILED:
    print("RESULT: FAIL (%d)" % len(FAILED))
    for f in FAILED:
        print("  - " + f)
    sys.exit(1)
print("RESULT: ALL PASS")
