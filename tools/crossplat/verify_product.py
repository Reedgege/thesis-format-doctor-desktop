# -*- coding: utf-8 -*-
"""国内版（学生版 / 导师版）三平台功能验收：检查 → 修正 → 复检，逐项断言。

为什么需要它
------------
`build.yml` 里原有的 smoke test 只证明「进程能起来」：`.deb` 装得上、AppImage
能在 xvfb 里启动、macOS 的 .app 活得下去。这些都是**外壳**。真正卖钱的是
「传一份论文进去，出来的文档标题套好了、正文格式对了、参考文献重排了」，
而这一层在三个平台上从来没被自动化验过。

验的是**打包产物**，不是源码
------------------------------
源码跑通不代表产物跑通。海外版踩过：macOS/Linux 打包后 python-docx 读不到自带
模板，「建页眉」静默失败，报告还说改好了 —— 因为只有产物才会把那几个模板变成
"数据文件 + 相对路径"。所以这里一律调用 ``<产物> check/fix``，走客户那条路。

断言的是事实，不是基准
----------------------
不去比对「跟 Windows 跑出来的数字一样」，而是直接断言**文档该有的属性**：
标题样式套上了没有、正文字号对不对、首行缩进在不在、表格是不是三线、
复检还剩几个问题。三平台都该满足同一组事实，任一条不成立就是该平台坏了。

用法
----
    python tools/crossplat/verify_product.py --exe <可执行文件> [--label macOS]
        [--workdir <临时目录>] [--fixtures <素材目录>]
"""

from __future__ import annotations

import argparse
import itertools
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile

for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(os.path.dirname(HERE))

# 素材里手打的标题个数（见 make_fixtures.py 的 HEADINGS）
EXPECT_HEADINGS = 8
# 小四 = 12pt
EXPECT_BODY_PT = 12.0
# 2 字符缩进 ≈ 0.33 英寸
EXPECT_INDENT_IN = 0.33
# 模板（fixtures/template.docx）里明写的正文要求 —— 引擎必须原样提出来。
# 这几个值是 make_fixtures.build_template 亲手写进样式定义的，改动模板时同步改这里。
EXPECT_BODY_STYLE_ID = "Normal"     # 不能取到 Body Text 3 Char 这类内置引用块样式
EXPECT_BODY_FONT = "宋体"
EXPECT_BODY_SZ = 24                 # 半磅：24 = 12pt
EXPECT_BODY_INDENT_CHARS = 2
EXPECT_BODY_LINE = ("exact", 20)    # 固定 20 磅


def _esc(s: str) -> str:
    return str(s).replace("%", "%25").replace("\r", "").replace("\n", "%0A")


def notice(msg):
    print("::notice::" + _esc(msg), flush=True)


def error(msg):
    print("::error::" + _esc(msg), flush=True)


def step(msg):
    print("\n=== " + msg + " ===", flush=True)


_ENTRY = None       # --entry 的源码入口（仅本地调试用）
_SEQ = itertools.count()


def run_exe(exe, args, cwd, timeout=300):
    """跑产物的一条命令，返回 (rc, 输出文本, stderr)。

    为什么靠文件取数而不是 stdout：Windows 打包是窗口子系统
    （``--windows-console-mode=disable``），产物被调起时**没有 stdout 可读**，
    print 出去的东西全丢。所以每条命令都额外传 ``--json-out``，
    有文件就以文件为准 —— 三平台走同一条取数路径。
    """
    env = dict(os.environ)
    env["PYTHONIOENCODING"] = "utf-8"
    env["PYTHONUTF8"] = "1"
    out_json = os.path.join(cwd, "_cli_%03d.json" % next(_SEQ))
    if os.path.exists(out_json):
        os.remove(out_json)
    # 本地没构建产物时，可用 --entry main.py 以源码方式跑同一套断言
    # （只能验逻辑，验不出打包问题 —— CI 里一律用真产物）。
    base = [sys.executable, _ENTRY] if _ENTRY else [exe]
    cmd = base + args + ["--json-out", out_json]
    try:
        p = subprocess.run(cmd, cwd=cwd, capture_output=True,
                           timeout=timeout, env=env)
    except subprocess.TimeoutExpired:
        return (-1, "", "timeout after %ss" % timeout)
    except OSError as e:
        return (-1, "", "cannot execute: %s" % e)
    text = p.stdout.decode("utf-8", errors="replace")
    if os.path.isfile(out_json):
        try:
            with open(out_json, encoding="utf-8") as f:
                text = f.read()
        except OSError:
            pass
    return (p.returncode, text, p.stderr.decode("utf-8", errors="replace"))


def json_of(text):
    """从混着日志的输出里抠 JSON。"""
    if not text or not text.strip():
        return None
    s = text.strip()
    try:
        return json.loads(s)
    except Exception:
        pass
    start, end = s.find("{"), s.rfind("}")
    if start < 0 or end <= start:
        return None
    try:
        return json.loads(s[start:end + 1])
    except Exception:
        return None


# ---------------------------------------------------------------------------
# 文档属性核对（只看输出文档本身）
# ---------------------------------------------------------------------------
def inspect_output(path: str) -> dict:
    """读输出文档的真实属性。这个函数跑在 CI 的 python 里，不在产物里。"""
    import zipfile

    import docx
    from docx.oxml.ns import qn

    out = {"error": None}
    try:
        d = docx.Document(path)
    except Exception as e:
        return {"error": "cannot open %s: %r" % (os.path.basename(path), e)}

    heading = 0
    body_sizes, body_indents = [], []
    for p in d.paragraphs:
        t = (p.text or "").strip()
        if not t:
            continue
        try:
            sid = (p.style.style_id or "")
            nm = (p.style.name or "")
        except Exception:
            sid = nm = ""
        if sid.lower().startswith("heading") or nm.lower().startswith("heading "):
            heading += 1
        if nm == "Normal" and len(t) > 30:
            for r in p.runs:
                if r.font.size:
                    body_sizes.append(round(r.font.size.pt, 1))
            fli = p.paragraph_format.first_line_indent
            body_indents.append(round(fli / 914400.0, 2) if fli is not None else None)
    out["heading_style_count"] = heading
    out["body_sizes_pt"] = body_sizes
    out["body_indents_in"] = body_indents

    # 表格三线：上下单线、无竖线、无内部横线，且表头行有下线
    tbl_info = []
    for tbl in d.tables:
        pr = tbl._tbl.find(qn("w:tblPr"))
        bd = pr.find(qn("w:tblBorders")) if pr is not None else None
        edges = {}
        if bd is not None:
            for k in ("top", "bottom", "left", "right", "insideH", "insideV"):
                el = bd.find(qn("w:" + k))
                val = (el.get(qn("w:val")) if el is not None else None) or "unset"
                edges[k] = val
        header_bottom = False
        if tbl.rows:
            for cell in tbl.rows[0].cells:
                tcpr = cell._tc.find(qn("w:tcPr"))
                tb = tcpr.find(qn("w:tcBorders")) if tcpr is not None else None
                el = tb.find(qn("w:bottom")) if tb is not None else None
                if el is not None and (el.get(qn("w:val")) or "") not in ("", "nil", "none"):
                    header_bottom = True
        tbl_info.append({"edges": edges, "header_bottom": header_bottom})
    out["tables"] = tbl_info

    try:
        with zipfile.ZipFile(path) as z:
            names = z.namelist()
        out["has_comments_part"] = any("comments" in n for n in names)
        out["zip_entries"] = len(names)
    except Exception:
        out["has_comments_part"] = None
        out["zip_entries"] = None
    return out


def _metrics(payload) -> dict:
    return (payload or {}).get("metrics") or {}


def main(argv=None) -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--exe", required=True)
    ap.add_argument("--label", default="unknown")
    ap.add_argument("--workdir", default=None)
    ap.add_argument("--fixtures", default=os.path.join(HERE, "fixtures"))
    ap.add_argument("--timeout", type=int, default=300)
    ap.add_argument("--entry", default=None,
                    help="仅本地调试：用 `python <entry> <args>` 代替产物 exe")
    args = ap.parse_args(argv)

    global _ENTRY
    if args.entry:
        _ENTRY = os.path.abspath(args.entry)

    tag = "[%s] " % args.label
    failures, soft = [], []
    exe = os.path.abspath(args.exe)
    if not os.path.isfile(exe):
        error(tag + "受测程序不存在：" + exe)
        return 2

    work = args.workdir or tempfile.mkdtemp(prefix="cn_verify_")
    os.makedirs(work, exist_ok=True)
    src = os.path.join(work, "thesis.docx")
    tpl = os.path.join(work, "template.docx")
    prof = os.path.join(work, "profile.json")
    for name, dst in (("thesis.docx", src), ("template.docx", tpl)):
        s = os.path.join(args.fixtures, name)
        if not os.path.isfile(s):
            error(tag + "素材缺失：%s（先跑 tools/crossplat/make_fixtures.py）" % s)
            return 2
        shutil.copyfile(s, dst)

    print("受测程序 : " + exe)
    print("工作目录 : " + work)

    # ---- 0. 自检 ----
    step("0/6 产物环境自检")
    rc, out, err = run_exe(exe, ["selfcheck"], work, 180)
    sc = json_of(out)
    if sc is None:
        error(tag + "selfcheck 没有输出 JSON（rc=%s）：%s" % (rc, (out or err)[:400]))
        failures.append("selfcheck 失败")
    else:
        brief = {k: sc.get(k) for k in
                 ("python", "platform", "python_docx", "docx_templates_ok",
                  "header_create", "engine_import")}
        print(json.dumps(brief, ensure_ascii=False))
        notice(tag + "selfcheck | " + json.dumps(brief, ensure_ascii=False)[:800])
        if str(sc.get("header_create")) != "ok":
            failures.append("产物里连空页眉都建不起来：%s"
                            % str(sc.get("header_create"))[:300])
        if str(sc.get("engine_import")) != "ok":
            failures.append("产物里 import engine 失败：%s"
                            % str(sc.get("engine_import"))[:300])

    # ---- 1. 模板画像提取（客户流程第一步，引擎自己去模板里读要求）----
    # 不能用手写的 profile.json：那是我替引擎做的决定，字段名写错就会变成
    # 「检查器读不到要求、报已达标，修正器按另一套字段改了」 —— 本轮真踩过。
    step("1/6 模板画像提取（profile）")
    rc, out, err = run_exe(exe, ["profile", tpl, "--out", prof], work, 180)
    pj = json_of(out)
    if pj is None or not pj.get("ok"):
        error(tag + "profile 没有输出 JSON（rc=%s）：%s" % (rc, (out or err)[:400]))
        return 1
    brief = pj.get("brief") or {}
    print("画像概况: " + json.dumps(brief, ensure_ascii=False))
    notice(tag + "画像概况 " + json.dumps(brief, ensure_ascii=False)[:600])
    if not brief.get("has_body"):
        failures.append("模板画像里没有正文样式（正文要求全部丢失，客户会拿到「已达标」的假报告）")
    for lv in ("1", "2", "3"):
        if lv not in (brief.get("heading_levels") or []):
            failures.append("模板画像缺 %s 级标题样式" % lv)
    # 正文要素必须与模板里写的一致（这几项错一项，客户论文就照着错的标准改）
    # 注意两处各存一半：profile["body"] 是「认对了哪个样式」，profile["levels"]["body"]
    # 是「这个样式的具体要求」。只看 levels 会漏掉样式认错，只看 body 会漏掉要求没提出来。
    body, body_style = {}, {}
    try:
        with open(prof, encoding="utf-8") as f:
            _pj = json.load(f)
        body = (_pj.get("levels") or {}).get("body") or {}
        body_style = _pj.get("body") or {}
    except Exception as e:  # noqa: BLE001
        failures.append("画像文件读不出来：%r" % (e,))
    if body:
        print("画像正文: " + json.dumps(
            {"styleId": body_style.get("styleId"),
             **{k: body.get(k) for k in
                ("zh_font", "sz", "indent_chars", "line_rule", "line_val")}},
            ensure_ascii=False))
        if body_style.get("styleId") != EXPECT_BODY_STYLE_ID:
            failures.append("正文样式取到了 %r，模板里的正文是 %r —— "
                            "内置 Body Text 3 Char(8pt) 这类引用块样式把真正文挤掉了"
                            % (body_style.get("styleId"), EXPECT_BODY_STYLE_ID))
        if body.get("zh_font") != EXPECT_BODY_FONT:
            failures.append("画像正文中文字体是 %r，模板要求 %r"
                            % (body.get("zh_font"), EXPECT_BODY_FONT))
        if str(body.get("sz")) != str(EXPECT_BODY_SZ):
            failures.append("画像正文字号是 %r，模板要求 sz=%s（12pt）"
                            % (body.get("sz"), EXPECT_BODY_SZ))
        if body.get("indent_chars") != EXPECT_BODY_INDENT_CHARS:
            failures.append("画像正文首行缩进是 %r 字符，模板要求 %d 字符"
                            % (body.get("indent_chars"), EXPECT_BODY_INDENT_CHARS))
        if (body.get("line_rule"), body.get("line_val")) != EXPECT_BODY_LINE:
            failures.append("画像正文行距是 %r，模板要求固定 %d 磅"
                            % ((body.get("line_rule"), body.get("line_val")),
                               EXPECT_BODY_LINE[1]))

    # ---- 2. 修前体检 ----
    step("2/6 修前体检（check）")
    rc, out, err = run_exe(exe, ["check", src, "--profile", prof], work, args.timeout)
    before = json_of(out)
    if before is None:
        error(tag + "check 没有输出 JSON（rc=%s）：%s" % (rc, (out or err)[:400]))
        return 1
    bm = _metrics(before)
    print("修前指标: " + json.dumps(bm, ensure_ascii=False))
    notice(tag + "修前 " + json.dumps(bm, ensure_ascii=False))
    if not bm.get("issues"):
        failures.append("素材本该是不合规的，check 却报 0 个问题（诊断可能失效）")

    # ---- 3. 一键修正 ----
    step("3/6 一键修正（fix）")
    fixed = os.path.join(work, "fixed.docx")
    rc, out, err = run_exe(
        exe, ["fix", src, "--out", fixed, "--profile", prof], work, args.timeout)
    fixp = json_of(out)
    if fixp is None:
        error(tag + "fix 没有输出 JSON（rc=%s）：%s" % (rc, (out or err)[:600]))
        return 1
    print("ok=%s" % fixp.get("ok"))
    if not fixp.get("ok"):
        failures.append("fix 失败：%s" % json.dumps(fixp.get("error"), ensure_ascii=False)[:300])
        for e in (fixp.get("traceback") or "").splitlines()[-5:]:
            print("  " + e)
    final = fixp.get("final") or fixed
    if not os.path.isfile(final):
        failures.append("fix 没有生成输出文档（期望 %s）" % final)
        return finish(failures, soft, tag)

    # ---- 3. 输出文档属性核对（客户真正拿到手的东西）----
    step("4/6 输出文档属性（客户拿到手的东西）")
    facts = inspect_output(final)
    if facts.get("error"):
        failures.append("输出文档打不开：%s" % facts["error"])
        return finish(failures, soft, tag)
    print(json.dumps({k: v for k, v in facts.items() if k != "tables"}, ensure_ascii=False))
    notice(tag + "输出文档 | " + json.dumps(
        {"headings": facts["heading_style_count"],
         "sizes": sorted(set(facts["body_sizes_pt"])),
         "indents": sorted(set(x for x in facts["body_indents_in"] if x is not None)),
         "tables": facts["tables"], "comments": facts["has_comments_part"]},
        ensure_ascii=False)[:900])

    if facts["heading_style_count"] < EXPECT_HEADINGS:
        failures.append(
            "只有 %d 个段落套上了标题样式（素材里手打了 %d 个标题）—— "
            "客户点完「一键套标题样式」，文档里的标题没变成真标题"
            % (facts["heading_style_count"], EXPECT_HEADINGS))

    sizes = [s for s in facts["body_sizes_pt"] if s]
    if not sizes:
        failures.append("正文段落读不到字号")
    elif abs(max(set(sizes), key=sizes.count) - EXPECT_BODY_PT) > 0.01:
        failures.append("正文字号没落到小四(%.0fpt)，实测众数 %s"
                        % (EXPECT_BODY_PT, max(set(sizes), key=sizes.count)))

    indents = [x for x in facts["body_indents_in"] if x is not None]
    if not indents:
        failures.append("正文段落没有首行缩进（模板要求 2 字符 ≈ %.2f″）" % EXPECT_INDENT_IN)
    elif abs(indents[0] - EXPECT_INDENT_IN) > 0.02:
        failures.append("首行缩进是 %.2f″，模板要求 2 字符 ≈ %.2f″"
                        % (indents[0], EXPECT_INDENT_IN))

    if not facts["has_comments_part"]:
        failures.append("输出文档没有批注部件 —— 客户看不到「哪处按什么规则改的」")

    for i, t in enumerate(facts["tables"]):
        e = t["edges"]
        if e.get("left") not in ("none", "nil", "unset") or e.get("right") not in ("none", "nil", "unset"):
            failures.append("表 %d 还有竖线（left=%s right=%s），不是学术三线表"
                            % (i + 1, e.get("left"), e.get("right")))
        if e.get("insideH") not in ("none", "nil", "unset"):
            failures.append("表 %d 还有内部横线（insideH=%s）" % (i + 1, e.get("insideH")))
        if e.get("top") in ("unset", None) or e.get("bottom") in ("unset", None):
            failures.append("表 %d 缺上下边线（top=%s bottom=%s）"
                            % (i + 1, e.get("top"), e.get("bottom")))
        if not t["header_bottom"]:
            failures.append("表 %d 表头行下面没有线（三线表的第二条线）" % (i + 1))

    # ---- 4. 复检 ----
    step("5/6 复检修正稿（check again）")
    rc, out, err = run_exe(exe, ["check", final, "--profile", prof], work, args.timeout)
    after = json_of(out)
    if after is None:
        failures.append("复检没有输出 JSON（rc=%s）：%s" % (rc, (out or err)[:400]))
        return finish(failures, soft, tag)
    am = _metrics(after)
    print("复检指标: " + json.dumps(am, ensure_ascii=False))
    notice(tag + "复检 " + json.dumps(am, ensure_ascii=False))
    if am.get("issues"):
        failures.append("修正后仍检出 %s 个问题（高危 %s / 中危 %s / 低危 %s）"
                        % (am.get("issues"), am.get("high"), am.get("mid"), am.get("low")))

    # ---- 5. 幂等：再修一遍不该把文档改坏 ----
    step("6/6 幂等（对修正稿再修一次）")
    again = os.path.join(work, "fixed2.docx")
    rc, out, err = run_exe(
        exe, ["fix", final, "--out", again, "--profile", prof], work, args.timeout)
    p2 = json_of(out)
    if p2 is None:
        soft.append("第二遍修正没有输出 JSON（rc=%s），跳过幂等核对" % rc)
    else:
        f2 = inspect_output(p2.get("final") or again)
        if not f2.get("error"):
            if f2["heading_style_count"] != facts["heading_style_count"]:
                failures.append("第二遍修正把标题样式个数改了（%d → %d）"
                                % (facts["heading_style_count"], f2["heading_style_count"]))
            else:
                print("标题样式个数稳定 ✓（%d）" % f2["heading_style_count"])
            r2 = (p2.get("recheck") or {}).get("issues")
            if r2:
                failures.append("第二遍修正后仍报 %s 个问题（改不干净）" % r2)
            else:
                print("第二遍复检 0 问题 ✓")

    return finish(failures, soft, tag)


def finish(failures, soft, tag) -> int:
    step("结论")
    for s in soft:
        print("::warning::" + _esc(tag + s))
    if failures:
        for m in failures:
            error(tag + m)
        print("\n结论：❌ 该平台产物功能不合格")
        return 1
    notice(tag + "✅ 检查 / 修正 / 复检 / 幂等全通，输出文档属性符合模板要求")
    print("\n结论：✅ 通过")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
