# -*- coding: utf-8 -*-
"""运行时资源回归（学生版）
==================================================================
背景：build.py 若没把 ``tfd_app/assets`` 带进产物，而读取方都写了
``if os.path.isfile(...)`` 守卫 → 打包版会「**二维码不显示、水印图丢失、
主界面背景图缺失**」，而且**全程静默**。本测试在**打包前**就把这类缺件钉住：

  1. ``REQUIRED_ASSETS`` 每个文件在源码树里真实存在（缺一个 = 会打个缺件的包）；
  2. ``assetpath.find_asset`` 能定位到它们（多候选路径逻辑没写坏）；
  3. ``build.py`` 以 ``REQUIRED_ASSETS`` 为唯一事实来源、真的生成
     ``--include-data-files=``，且缺件时红灯退出（BRIEF 第七节第 3 条）；
  4. gui 的 ICON / QRCODE / MINIAPP_QRCODE 与 watermark 的两张水印图路径，
     都指向真实文件（导入即求值，等于校验了资源接线）；
  5. 学生版专属：主界面唯一背景层 = ``student_background.png``（与装修包
     ``STUDENT_BACKGROUND.png`` 逐字节一致），且导师版专属的
     ``advisor_background.png`` **不在本仓库**。

用法：``python tests/test_assets.py``（无 tkinter 的环境会自动跳过第 4 节）
"""
import hashlib
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from tfd_app import assetpath          # noqa: E402

fails = []


def check(cond, label):
    print(("  ok  " if cond else "  FAIL ") + label)
    if not cond:
        fails.append(label)


print("=== 1. 必需资源清单 ===")
missing = assetpath.missing_assets()
check(not missing, "REQUIRED_ASSETS 全部存在（缺失：%s）" % (missing or "无"))
for name in assetpath.REQUIRED_ASSETS:
    p = assetpath.find_asset(name)
    check(p is not None and os.path.isfile(p), "find_asset(%r) -> %s" % (name, p))

print("=== 2. 构建脚本的清单与资源一致 ===")
build_py = os.path.join(ROOT, "build.py")
src = open(build_py, encoding="utf-8").read()
check("REQUIRED_ASSETS" in src, "build.py 以 assetpath.REQUIRED_ASSETS 为唯一事实来源")
check("--include-data-files=" in src, "build.py 会把资源带进产物")
check("构建中止" in src, "build.py 缺件时红灯退出（不静默发缺件包）")

print("=== 3. 学生版专属：背景图唯一 + 不留导师版背景 ===")
for name in assetpath.FORBIDDEN_ASSETS:
    check(not os.path.isfile(os.path.join(assetpath.asset_dir(), name)),
          "导师版专属素材不在本仓库：%s" % name)
bg = assetpath.find_asset("student_background.png")
check(bg is not None, "学生版背景图存在：student_background.png")
# BRIEF §7.7：这张底图已由我方擦掉旧界面残影，**禁止还原成装修包原图**。
# 原图指纹 md5[0:16] = 089372a6b3f2434a；擦净版 = bcea1471a3d5a06c。
CLEAN_BG_MD5_16 = "bcea1471a3d5a06c"
if bg:
    with open(bg, "rb") as fh:
        raw = fh.read()
    check(raw[:8] == b"\x89PNG\r\n\x1a\n", "背景图是合法 PNG（可被 Tk PhotoImage 读取）")
    md5_16 = hashlib.md5(raw).hexdigest()[:16]
    check(md5_16 == CLEAN_BG_MD5_16,
          "背景图是擦净旧 UI 残影的版本（md5[0:16]=%s，期望 %s）"
          % (md5_16, CLEAN_BG_MD5_16))
    mock = os.path.join(ROOT, "_ui_ref", "01_STUDENT", "background",
                        "STUDENT_BACKGROUND.png")
    if os.path.isfile(mock):
        with open(mock, "rb") as fh:
            pack_raw = fh.read()
        check(raw != pack_raw,
              "背景图未被还原成装修包原图（原图烘着旧 UI 残影，不得还原）")
    else:
        print("  skip 装修包原图不在本仓库，跳过「未还原」比对")

print("=== 4. gui / watermark / backdrop 的资源接线 ===")
try:
    from tfd_app import gui as gui_mod          # 需要 tkinter
    from tfd_app import watermark as wm_mod
except ImportError as e:                        # pragma: no cover - 无 tkinter 的环境
    print("  skip 本机无 tkinter，跳过资源接线检查：%s" % e)
else:
    check(os.path.isfile(gui_mod.ICON), "gui.ICON -> %s" % gui_mod.ICON)
    check(os.path.isfile(gui_mod.QRCODE), "gui.QRCODE -> %s" % gui_mod.QRCODE)
    check(os.path.isfile(gui_mod.MINIAPP_QRCODE),
          "gui.MINIAPP_QRCODE -> %s" % gui_mod.MINIAPP_QRCODE)
    check(os.path.isfile(wm_mod._IMG_SRC), "watermark 水印图 -> %s" % wm_mod._IMG_SRC)
    check(os.path.isfile(wm_mod._BG_IMG_SRC),
          "watermark 背景水印图 -> %s" % wm_mod._BG_IMG_SRC)
    from tfd_app.ui import backdrop as bd_mod
    check(bd_mod.BACKDROP_ASSET == "student_background.png",
          "backdrop.BACKDROP_ASSET = student_background.png（不当视觉稿/水印当背景）")
    bd_src = open(bd_mod.__file__, encoding="utf-8").read()
    check("STUDENT_UI_REFERENCE" not in bd_src and "watermark_bg" not in bd_src,
          "backdrop 不引用整屏视觉稿 / 输出水印图作背景")

print("=== 5. 版本号一致性（VERSION 文件 / buildinfo / gui.APP_VERSION）===")
import re as _re                                          # noqa: E402
from tfd_app import buildinfo                             # noqa: E402
ver_file = open(os.path.join(ROOT, "VERSION"), encoding="utf-8").read().strip()
check(ver_file.lstrip("vV") == buildinfo.APP_VERSION,
      "VERSION 文件(%s) 与 buildinfo.APP_VERSION(%s) 一致" % (ver_file, buildinfo.APP_VERSION))
_screen = open(os.path.join(ROOT, "tfd_app", "ui2", "screen.py"), encoding="utf-8").read()
_lit = _re.findall(r"v1\.[0-9]+\.[0-9]+", _screen)
check(not _lit, "ui2/screen.py 里没有写死的版本号（找到：%s）" % (_lit or "无"))
try:
    from tfd_app import gui as _gui_v                       # noqa: E402
except ImportError:                                          # pragma: no cover
    print("  skip 本机无 tkinter，跳过 gui.APP_VERSION 比对")
else:
    check(_gui_v.APP_VERSION == buildinfo.APP_VERSION,
          "gui.APP_VERSION(%s) 与 buildinfo.APP_VERSION(%s) 一致"
          % (_gui_v.APP_VERSION, buildinfo.APP_VERSION))

print()
if fails:
    print("==== 资源回归失败 %d 项 ====" % len(fails))
    for f in fails:
        print("  - " + f)
    raise SystemExit(1)
print("==== 资源回归全部通过 ====")
