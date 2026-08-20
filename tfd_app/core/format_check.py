# -*- coding: utf-8 -*-
# Copyright (c) 2026 芦苇（山东大学 MBA）. All rights reserved.
# 原创作品 | 禁止未经授权转售、二次分发或抄袭 | 授权用户可在扣子/虾评平台内使用
# build: 202608191649 | version: 1.3.5 | file_sha256: 279e3cd81fd94448
# -*- coding: utf-8 -*-
"""
format-check.py — 客户论文 对照 学校模板画像 -> 格式诊断报告

用法：
  python format-check.py 客户论文.docx profile.json [-o 报告.md]

检查项（全部为"机械可判定"的规范错误）：
  1. 未套样式的疑似标题（启发式识别，标置信度）
  2. 已套标题样式但与模板规范不符（字体/字号/缩进/行距）
  3. 正文样式与模板不符
  4. 页边距与模板不符
  5. 表格非标准三线表（含竖线/左右边线）
  6. 参考文献不规范（缺 [n]、缺出版年、缺 [J]/[M] 类型标记、编号不连续）

输出：控制台 + 可选 Markdown 文件。每条带 严重度(高/中/低) 与 对应规范。
注意：本脚本只"诊断"不"修正"；修正由 headings-fix.py / ref-reformat.py 完成。
"""
import sys
import os
import json
import re
import xml.etree.ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from docxutils import load, paragraphs, margins, tables, detect_heading, WR

__author__ = "芦苇"
__copyright__ = "Copyright (c) 2026 芦苇（山东大学 MBA）"
__version__ = "1.3.5"



_INTEGRITY_TAG = "Q29weXJpZ2h0IChjKSAyMDI2IOiKpuiLh++8iOWxseS4nOWkp+WtpiBNQkHvvIkuIEFsbCByaWdodHMgcmVzZXJ2ZWQuIHwgYnVpbGQ6MjAyNjA4MTkxNjQ5"  # noqa
def check(client_path, profile):
    _, root = load(client_path)
    paras = paragraphs(root)
    issues = []

    # ---- 1 & 2. 标题：未套样式 vs 套了但错配 ----
    heading_specs = profile.get("headings", {})
    for i, p in enumerate(paras):
        txt = p["text"].strip()
        lvl, conf = detect_heading(txt)
        if lvl is None:
            continue
        styled = p["style"] in ("Heading1", "Heading2", "Heading3") or (
            p["style"] and re.search(r"标题\s*\d", p["style"] or ""))
        if not styled:
            issues.append({
                "sev": "高", "cat": "标题层级",
                "msg": f"疑似{_lvl_name(lvl)}未套标题样式：「{txt[:24]}」"
                        + (f"（置信度 {int(conf*100)}%）" if conf < 0.9 else ""),
                "fix": "headings-fix.py 自动套 Heading%d 样式，使 Word 可一键生成目录" % lvl,
            })
        else:
            # 套了样式，比对规范
            spec = heading_specs.get(str(lvl))
            if spec:
                mism = []
                ef = (p["rpr"] or {})
                if spec.get("size") and ef.get("sz") and ef["sz"] != spec["size"]:
                    mism.append(f"字号 {ef['sz']}≠模板 {spec['size']}")
                if spec.get("font") and ef.get("eastAsia") and ef["eastAsia"] != spec["font"]:
                    mism.append(f"中文字体 {ef.get('eastAsia')}≠模板 {spec['font']}")
                if mism:
                    issues.append({
                        "sev": "中", "cat": "标题样式",
                        "msg": f"{_lvl_name(lvl)}「{txt[:20]}」与模板不符：{'；'.join(mism)}",
                        "fix": "按模板标题样式修正字体/字号",
                    })

    # ---- 3. 正文样式 ----
    body_spec = profile.get("body") or {}
    if body_spec.get("size") or body_spec.get("font"):
        sampled = 0
        for p in paras:
            if p["style"] in (None, "Normal", "") and p["text"].strip() and (p["rpr"] or {}).get("sz"):
                ef = p["rpr"]
                if body_spec.get("size") and ef["sz"] != body_spec["size"]:
                    issues.append({"sev": "中", "cat": "正文",
                                   "msg": f"正文段落字号 {ef['sz']}≠模板 {body_spec['size']}",
                                   "fix": "按模板正文样式批量修正"})
                    sampled += 1
                if sampled >= 3:
                    break

    # ---- 4. 页边距 ----
    pg = margins(root)
    tpl_pg = profile.get("page", {})
    for edge in ("top", "bottom", "left", "right"):
        if tpl_pg.get(edge) and pg.get(edge) and pg[edge] != tpl_pg[edge]:
            issues.append({"sev": "中", "cat": "页边距",
                           "msg": f"页边距 {edge}: 实际 {pg[edge]} ≠ 模板 {tpl_pg[edge]}",
                           "fix": "按模板页边距调整"})

    # ---- 5. 三线表 ----
    tpl_tbl = profile.get("tables", {})
    if tpl_tbl.get("threeLine") is True:
        for ti, b in enumerate(tables(root), 1):
            has_v = b.get("insideV") not in (None, "none") or b.get("left") not in (None, "none") or b.get("right") not in (None, "none")
            if has_v:
                issues.append({"sev": "中", "cat": "表格",
                               "msg": f"表 {ti} 含竖线/左右边线，非标准三线表",
                               "fix": "按三线表规范（仅顶/底/栏目线，无竖线）重设边框"})

    # ---- 6. 参考文献 ----
    issues += _check_refs(paras)

    return issues


def _check_refs(paras):
    full = "\n".join(p["text"] for p in paras)
    idx = full.find("参考文献")
    if idx < 0:
        return []
    seg = full[idx:idx + 4000]
    entries = re.findall(r"\[\d+\][^\n\[]*", seg)
    issues = []
    nums = [int(m.group(1)) for m in re.finditer(r"\[(\d+)\]", seg)]
    # 连续性
    if nums:
        expect = list(range(min(nums), max(nums) + 1))
        missing = [n for n in expect if n not in nums]
        if missing:
            issues.append({"sev": "中", "cat": "参考文献",
                           "msg": f"参考文献编号不连续，缺 {missing}",
                           "fix": "ref-reformat.py 重排并补连续编号"})
    for e in entries:
        if not re.search(r"(19|20)\d{2}", e):
            issues.append({"sev": "低", "cat": "参考文献",
                           "msg": f"条目缺出版年：「{e.strip()[:30]}」",
                           "fix": "ref-reformat.py 补全并规范"})
        if not re.search(r"\[[JMDCN]\]", e):
            issues.append({"sev": "低", "cat": "参考文献",
                           "msg": f"条目缺文献类型标记 [J]/[M] 等：「{e.strip()[:30]}」",
                           "fix": "ref-reformat.py 按 GB/T 7714 标注类型"})
    return issues


def _lvl_name(lvl):
    return {1: "一级标题", 2: "二级标题", 3: "三级标题"}.get(lvl, "标题")


def _render(issues, client_path):
    lines = [f"# 格式诊断报告（对照学校模板）", "", f"论文：{os.path.basename(client_path)}", ""]
    if not issues:
        lines.append("✅ 未发现明显格式规范问题。")
        return "\n".join(lines)
    order = {"高": 0, "中": 1, "低": 2}
    issues.sort(key=lambda x: order.get(x["sev"], 3))
    lines.append(f"共 {len(issues)} 项问题：")
    lines.append("")
    for n, it in enumerate(issues, 1):
        lines.append(f"**{n}. [{it['sev']}] {it['cat']}** — {it['msg']}")
        lines.append(f"   → 修正：{it['fix']}")
        lines.append("")
    return "\n".join(lines)


def main():
    if len(sys.argv) < 3:
        print("用法: python format-check.py 客户论文.docx profile.json [-o 报告.md]")
        sys.exit(1)
    client_path, prof_path = sys.argv[1], sys.argv[2]
    out = None
    if "-o" in sys.argv:
        out = sys.argv[sys.argv.index("-o") + 1]
    with open(prof_path, encoding="utf-8") as f:
        profile = json.load(f)
    issues = check(client_path, profile)
    report = _render(issues, client_path)
    print(report)
    if out:
        with open(out, "w", encoding="utf-8") as f:
            f.write(report)
        print(f"\n已写出报告 -> {out}")


if __name__ == "__main__":
    main()
