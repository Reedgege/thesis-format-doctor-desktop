# -*- coding: utf-8 -*-
"""命令行入口 —— 让打包后的产物能被自动化验收，也让客服有排障手段。

为什么要它
----------
本产品一直是纯 GUI：双击进界面，上传论文、点检查、点修正。这没问题，但带来两个
盲区：

1. **CI 只能验「能启动」**。`build.yml` 里那些 smoke test 只证明进程活着，证明不了
   「传一份论文进去，出来的文档真的改好了」。而打包恰恰是最容易坏在这一环的
   （海外版就出过：macOS/Linux 打包后读不到 python-docx 自带模板，页码静默失效，
   报告还说改好了 —— 见 thesis-format-doctor-global 的 src/docx_boot.py）。
2. **客户报「改了没用」时无据可查**。截图看不出文档内部属性，客服只能猜。

有了这个入口，CI 可以在三个平台的真产物上跑「检查 → 修正 → 复检」，
排障时也能让客户跑一条命令把现场打出来。

设计约束
--------
- **不改变默认行为**：不带参数仍然直接进 GUI。带子命令才走命令行。
- **输出双通道**：既 print，也支持 ``--json-out`` 写文件。Windows 打包是窗口子系统
  （``--windows-console-mode=disable``），double-click 时没有 stdout，
  写文件是唯一可靠的取数方式。
- 只调用 ``tfd_app.engine`` 里既有的入口函数，不复制、不绕过任何业务逻辑 ——
  验的必须是客户那条路。

用法
----
    <产品>.exe check 论文.docx [--profile 模板.json] [--json-out 结果.json]
    <产品>.exe fix   论文.docx --out 修正稿.docx [--profile 模板.json] [--no-refs]
    <产品>.exe selfcheck [--json-out 结果.json]
"""

from __future__ import annotations

import json
import os
import re
import sys
import traceback

# 子命令白名单：只有落到这里才走命令行，其余一律进 GUI（保证双击行为不变）
_COMMANDS = ("check", "fix", "profile", "selfcheck")


def _reconfig_out() -> None:
    """管道/重定向下强制 UTF-8，避免中文报告在 cp1252 管道里炸掉。"""
    for stream in (sys.stdout, sys.stderr):
        try:
            if stream is not None and not stream.isatty():
                stream.reconfigure(encoding="utf-8", errors="replace")
            elif stream is not None:
                stream.reconfigure(errors="replace")
        except Exception:
            pass


def _emit(payload: dict, json_out: str | None) -> None:
    text = json.dumps(payload, ensure_ascii=False, indent=2)
    if json_out:
        try:
            with open(json_out, "w", encoding="utf-8") as f:
                f.write(text)
        except OSError as e:
            print("cannot write %s: %s" % (json_out, e), file=sys.stderr)
    try:
        print(text)
    except Exception:
        pass


def _opt(argv: list, name: str, default=None):
    """取 `--name value` 的值；没有就返回 default。"""
    if name in argv:
        i = argv.index(name)
        if i + 1 < len(argv):
            return argv[i + 1]
    return default


def _metrics(markdown: str) -> dict:
    """从报告正文里抠出结构化指标（跨平台应当逐字一致的那几个数）。

    只取与平台无关的统计；报告里出现的文件名、路径一律不进指纹。
    """
    out = {"paragraphs": None, "tables": None, "issues": None,
           "high": None, "mid": None, "low": None}
    m = re.search(r"正文段落：\*\*(\d+)\*\*.*?表格：\*\*(\d+)\*\*.*?检出问题：\*\*(\d+)\*\*",
                  markdown, re.S)
    if m:
        out["paragraphs"], out["tables"], out["issues"] = (int(x) for x in m.groups())
    m = re.search(r"高危\s*\*\*(\d+)\*\*.*?中危\s*\*\*(\d+)\*\*.*?低危\s*\*\*(\d+)\*\*",
                  markdown, re.S)
    if m:
        out["high"], out["mid"], out["low"] = (int(x) for x in m.groups())
    return out


def _doc_facts(path: str) -> dict:
    """输出文档的内部属性 —— 「文档到底改好了没」就看这些。

    读不到（没装 python-docx / 文件坏了）不抛，返回 {"error": ...}，
    让调用方判断，而不是让整条验收挂掉。
    """
    try:
        import docx
        from docx.oxml.ns import qn
    except Exception as e:
        return {"error": "cannot import python-docx: %r" % (e,)}
    try:
        d = docx.Document(path)
    except Exception as e:
        return {"error": "cannot open docx: %r" % (e,)}

    heading_style = heading_custom = 0
    fonts, sizes = set(), set()
    body_indents = []
    for p in d.paragraphs:
        t = (p.text or "").strip()
        if not t:
            continue
        try:
            sid = (p.style.style_id or "")
            name = (p.style.name or "")
        except Exception:
            sid = name = ""
        if sid.lower().startswith("heading") or name.lower().startswith("heading "):
            heading_style += 1
        elif name and name != "Normal":
            heading_custom += 1
        for r in p.runs:
            if r.font.name:
                fonts.add(r.font.name)
            if r.font.size:
                sizes.add(round(r.font.size.pt, 1))
        if name == "Normal" and len(t) > 20:
            fli = p.paragraph_format.first_line_indent
            if fli is None:
                body_indents.append(None)
            else:
                body_indents.append(round(fli / 914400.0, 2))

    table_borders = []
    for tb in d.tables:
        try:
            tbl_pr = tb._tbl.find(qn("w:tblPr"))
            style = None
            if tbl_pr is not None:
                st = tbl_pr.find(qn("w:tblStyle"))
                if st is not None:
                    style = st.get(qn("w:val"))
            table_borders.append(style)
        except Exception:
            table_borders.append(None)

    try:
        sec = d.sections[0]
        margins_cm = {k: round(getattr(sec, "%s_margin" % k).cm, 2)
                      for k in ("top", "bottom", "left", "right")}
    except Exception:
        margins_cm = {}

    return {
        "paragraphs_nonempty": sum(1 for p in d.paragraphs if (p.text or "").strip()),
        "heading_style_count": heading_style,
        "custom_style_count": heading_custom,
        "fonts": sorted(fonts),
        "sizes_pt": sorted(sizes),
        "body_first_line_indents_in": body_indents[:8],
        "table_styles": table_borders,
        "margins_cm": margins_cm,
        "has_comments_part": any(n.startswith("word/comments")
                                 for n in _zip_names(path)),
    }


def _zip_names(path: str) -> list:
    try:
        import zipfile

        with zipfile.ZipFile(path) as z:
            return z.namelist()
    except Exception:
        return []


def cmd_check(argv: list) -> int:
    from . import engine

    if not argv:
        print("usage: check <docx> [--profile p.json] [--json-out f]", file=sys.stderr)
        return 2
    src, profile = argv[0], _opt(argv, "--profile")
    if not os.path.isfile(src):
        _emit({"ok": False, "cmd": "check", "error": "not found: %s" % src},
              _opt(argv, "--json-out"))
        return 2
    report = engine.run_check(src, profile_path=profile)
    _emit({"ok": True, "cmd": "check", "input": src,
           "profile": profile, "metrics": _metrics(report), "report": report},
          _opt(argv, "--json-out"))
    return 0


def cmd_fix(argv: list) -> int:
    from . import engine

    if not argv:
        print("usage: fix <docx> --out dst [--profile p.json] [--no-refs] [--json-out f]",
              file=sys.stderr)
        return 2
    src, dst = argv[0], _opt(argv, "--out")
    profile = _opt(argv, "--profile")
    if not dst:
        _emit({"ok": False, "cmd": "fix", "error": "--out is required"},
              _opt(argv, "--json-out"))
        return 2

    payload = {"ok": True, "cmd": "fix", "input": src, "out": dst, "profile": profile}
    # 三线表触发：不给就跟随模板画像（profile["tables"]["threeLine"]），
    # 与 GUI 的默认行为完全一致；验收时可用 --three-line / --no-three-line 强制。
    three_line = None
    if "--three-line" in argv:
        three_line = True
    elif "--no-three-line" in argv:
        three_line = False
    try:
        payload["fix_log"] = engine.run_fix_headings(
            src, dst, profile_path=profile, three_line=three_line)
    except TypeError:
        # 导师版/旧版没这个参数时退回基础调用，不让验收挂在一个可选开关上
        payload["fix_log"] = engine.run_fix_headings(src, dst, profile_path=profile)
    except Exception as e:      # noqa: BLE001  引擎报错要原样带回给调用方
        payload.update({"ok": False,
                        "error": "%s: %s" % (type(e).__name__, e),
                        "traceback": traceback.format_exc(limit=6)})
        _emit(payload, _opt(argv, "--json-out"))
        return 1

    if "--no-refs" not in argv:
        refs_out = _opt(argv, "--refs-out") or (os.path.splitext(dst)[0] + "_refs.docx")
        try:
            payload["refs_log"] = engine.run_reformat_refs(dst, refs_out)
            payload["refs_out"] = refs_out
        except Exception as e:  # noqa: BLE001
            payload["refs_error"] = "%s: %s" % (type(e).__name__, e)

    final = payload.get("refs_out") or dst
    payload["final"] = final
    payload["doc"] = _doc_facts(final)
    # 复检：输出文档到底还剩几个问题（客户视角的最终答案）
    try:
        payload["recheck"] = _metrics(engine.run_check(final, profile_path=profile))
    except Exception as e:  # noqa: BLE001
        payload["recheck_error"] = "%s: %s" % (type(e).__name__, e)

    _emit(payload, _opt(argv, "--json-out"))
    return 0


def cmd_selfcheck(argv: list) -> int:
    """打印运行环境与内置资源定位情况（排障用，客户可以一条命令把现场发过来）。"""
    info = {
        "ok": True,
        "cmd": "selfcheck",
        "python": sys.version.split()[0],
        "platform": sys.platform,
        "executable": sys.executable,
        "argv0": (sys.argv[0] if getattr(sys, "argv", None) else ""),
        "cwd": os.getcwd(),
        "frozen": bool(getattr(sys, "frozen", False)),
    }
    # python-docx 自带模板在产物里的落点与接管情况（mac/Linux 崩这里时一眼可见）
    try:
        from . import docx_boot

        info["docx_boot"] = docx_boot.report()
    except Exception:
        info["docx_boot"] = traceback.format_exc(limit=4)
    try:
        import docx

        info["python_docx"] = getattr(docx, "__version__", "?")
        info["docx_dir"] = os.path.dirname(os.path.abspath(docx.__file__))
        tpl = os.path.join(info["docx_dir"], "templates")
        info["docx_templates_ok"] = os.path.isfile(os.path.join(tpl, "default-header.xml"))
        # 真建一次页眉：打包后路径失效会在这里暴露（异常原文带回去）
        try:
            doc = docx.Document()
            _ = doc.sections[0].header.paragraphs
            info["header_create"] = "ok"
        except Exception:
            info["header_create"] = traceback.format_exc(limit=8)
    except Exception:
        info["docx_error"] = traceback.format_exc(limit=4)
    try:
        from . import engine  # noqa: F401

        info["engine_import"] = "ok"
    except Exception:
        info["engine_import"] = traceback.format_exc(limit=4)
    _emit(info, _opt(argv, "--json-out"))
    return 0


def cmd_profile(argv: list) -> int:
    """从学校模板 docx 提取格式画像（客户流程的第一步：上传模板）。"""
    from . import engine

    if not argv:
        print("usage: profile <template.docx> --out p.json [--json-out f]", file=sys.stderr)
        return 2
    tpl, out_json = argv[0], _opt(argv, "--out")
    if not out_json:
        _emit({"ok": False, "cmd": "profile", "error": "--out is required"},
              _opt(argv, "--json-out"))
        return 2
    if not os.path.isfile(tpl):
        _emit({"ok": False, "cmd": "profile", "error": "not found: %s" % tpl},
              _opt(argv, "--json-out"))
        return 2
    try:
        text = engine.run_build_profile(tpl, out_json)
    except Exception as e:      # noqa: BLE001
        _emit({"ok": False, "cmd": "profile",
               "error": "%s: %s" % (type(e).__name__, e),
               "traceback": traceback.format_exc(limit=6)},
              _opt(argv, "--json-out"))
        return 1
    try:
        with open(out_json, encoding="utf-8") as f:
            prof = json.load(f)
    except Exception:
        prof = None
    # 画像里到底抠到了什么：标题级数、正文要素有没有值 —— 一眼看出模板能不能用
    brief = {}
    if isinstance(prof, dict):
        brief = {
            "heading_levels": sorted((prof.get("headingStyles") or {}).keys()),
            "has_body": bool(prof.get("body")),
            "levels": sorted((prof.get("levels") or {}).keys()),
            "annotation_count": len(prof.get("annotations") or []),
        }
    _emit({"ok": True, "cmd": "profile", "template": tpl, "out": out_json,
           "brief": brief, "profile_text": text},
          _opt(argv, "--json-out"))
    return 0


def main(argv: list) -> int:
    """命令行入口。返回退出码（0 成功 / 1 执行失败 / 2 参数错）。"""
    _reconfig_out()
    cmd = argv[0]
    rest = argv[1:]
    if cmd == "check":
        return cmd_check(rest)
    if cmd == "fix":
        return cmd_fix(rest)
    if cmd == "profile":
        return cmd_profile(rest)
    if cmd == "selfcheck":
        return cmd_selfcheck(rest)
    print("unknown command: %s (available: %s)" % (cmd, ", ".join(_COMMANDS)),
          file=sys.stderr)
    return 2


def wants_cli(argv: list) -> bool:
    """是否需要走命令行。GUI 默认行为不能变，所以只在第一个参数是已知子命令时才拦。"""
    return bool(argv) and argv[0] in _COMMANDS
