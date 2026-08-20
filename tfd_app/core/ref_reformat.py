# -*- coding: utf-8 -*-
# Copyright (c) 2026 芦苇（山东大学 MBA）. All rights reserved.
# 原创作品 | 禁止未经授权转售、二次分发或抄袭 | 授权用户可在扣子/虾评平台内使用
# build: 202608191649 | version: 1.3.5 | file_sha256: 7738cb6a9a2019e0
# -*- coding: utf-8 -*-
"""
ref-reformat.py — 参考文献按 GB/T 7714 范式重排（含异常标注）

用法：
  python ref-reformat.py 输入.docx [-o 输出.docx]

能力：
  - 抽取"参考文献"区条目（[n] 标记）
  - 解析字段：作者 / 题名 / 文献类型[J][M][C][D] / 出处 / 出版年
  - 重排为标准范式：[n] 作者. 题名[类型]. 出处, 年.
  - 低置信度条目（字段缺失/解析不确定）明确标红，需人工复核
  - 可选写回新 docx（仅替换参考文献区文字，不碰正文其余部分）

说明：中文文献格式杂，纯规则解析必有误差；本工具对"干净条目"自动改、
"异常条目"标红，不追求 100% 无人值守。符合"调样式/文字排版、不改论点内容"。
"""
import sys
import os
import re
import xml.etree.ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import docxutils
from docxutils import load, write_docx, to_doc_xml, WR

__author__ = "芦苇"
__copyright__ = "Copyright (c) 2026 芦苇（山东大学 MBA）"
__version__ = "1.3.5"



_INTEGRITY_TAG = "Q29weXJpZ2h0IChjKSAyMDI2IOiKpuiLh++8iOWxseS4nOWkp+WtpiBNQkHvvIkuIEFsbCByaWdodHMgcmVzZXJ2ZWQuIHwgYnVpbGQ6MjAyNjA4MTkxNjQ5"  # noqa
def extract_refs(paras):
    """从段落列表中抽取参考文献条目。遇到表格或新标题即停止，避免吞入表格文字。"""
    out = []
    started = False
    for p in paras:
        if not started:
            if docxutils.is_reference_heading(p["text"]):
                started = True
            continue
        if p.get("in_table"):
            break
        if p["style"] and re.search(r"标题\s*\d", p["style"] or ""):
            break
        t = p["text"].strip()
        if not t:
            continue
        m = re.match(r"^\[(\d+)\]\s*(.*)", t, re.S)
        if m:
            out.append({"num": int(m.group(1)), "raw": m.group(2).strip()})
        elif out:
            # 续行，并入上一条
            out[-1]["raw"] += " " + t
    return out


def parse_ref(raw):
    """尽力解析单条参考文献，返回字段字典 + 置信度。"""
    text = raw.replace("\n", " ").strip()
    res = {"author": "", "title": "", "type": "", "source": "", "year": ""}
    conf = 0.5
    # 类型
    tm = re.search(r"\[([JMDCN])\]", text)
    if tm:
        res["type"] = tm.group(1)
        conf += 0.15
    # 年
    ym = re.search(r"(19|20)\d{2}", text)
    if ym:
        res["year"] = ym.group(0)
        conf += 0.15
    # 作者（题名前的第一段，通常以 . 或空格结束）
    am = re.match(r"^([^.。]+?)[.\s]", text)
    if am:
        res["author"] = am.group(1).strip()
        conf += 0.1
    # 题名（作者后到类型标记前）
    if res["author"]:
        rest = text[len(res["author"]):].lstrip(". ")
        tm2 = re.search(r"\[[JMDCN]\]", rest)
        if tm2:
            res["title"] = rest[:tm2.start()].strip().rstrip(".。")
            res["source"] = rest[tm2.end():].strip().rstrip(".。")
            conf += 0.1
    if res["author"] and res["title"]:
        conf = min(conf + 0.1, 1.0)
    return res, round(conf, 2)


def reformat(refs):
    out = []
    for r in refs:
        fields, conf = parse_ref(r["raw"])
        if fields["author"] and fields["type"] and fields["year"]:
            src = fields["source"].strip().lstrip(".。 ").rstrip(".。 ")
            new = f"[{r['num']}] {fields['author']}. {fields['title']}[{fields['type']}]. {src}."
        else:
            new = f"[{r['num']}] {r['raw']}  ⚠需人工核"
        out.append({"num": r["num"], "old": r["raw"], "new": new, "conf": conf})
    return out


def _render(results):
    lines = ["# 参考文献重排报告（GB/T 7714）", ""]
    for r in results:
        tag = "✅" if "⚠" not in r["new"] else "⚠"
        lines.append(f"{tag} [{r['num']}] (置信度 {r['conf']})")
        lines.append(f"   原：{r['old'][:60]}")
        lines.append(f"   新：{r['new'][:70]}")
        lines.append("")
    return "\n".join(lines)


def _replace_in_docx(src, dst, results):
    """把参考文献区 [n] 段落文字替换为重排结果（保留首 run 的字体样式）。"""
    _, root = load(src)
    mapping = {r["num"]: r["new"] for r in results}
    in_ref = False
    for p in root.iter(WR + "p"):
        txt = "".join(t.text or "" for t in p.iter(WR + "t"))
        if docxutils.is_reference_heading(txt):
            in_ref = True
            continue
        if not in_ref:
            continue
        m = re.match(r"^\s*\[(\d+)\]", txt)
        if m and int(m.group(1)) in mapping:
            _set_para_text(p, mapping[int(m.group(1))])
    xml = to_doc_xml(root)
    write_docx(src, dst, xml)


def _set_para_text(p, text):
    runs = p.findall(WR + "r")
    if not runs:
        # 无 run，构造一个
        r = ET.SubElement(p, WR + "r")
        t = ET.SubElement(r, WR + "t")
        t.text = text
        return
    # 保留首个 run 的 rPr，清空其文字，删除其余 run
    first = runs[0]
    t0 = first.find(WR + "t")
    if t0 is None:
        t0 = ET.SubElement(first, WR + "t")
    t0.text = text
    for r in runs[1:]:
        p.remove(r)


def main():
    if len(sys.argv) < 2:
        print("用法: python ref-reformat.py 输入.docx [-o 输出.docx]")
        sys.exit(1)
    src = sys.argv[1]
    out = None
    if "-o" in sys.argv:
        out = sys.argv[sys.argv.index("-o") + 1]
    from docxutils import paragraphs
    _, root = load(src)
    refs = extract_refs(paragraphs(root))
    if not refs:
        print("未检测到参考文献区。")
        return
    results = reformat(refs)
    print(_render(results))
    if out:
        _replace_in_docx(src, out, results)
        print(f"\n已写出重排后文档 -> {out}")


if __name__ == "__main__":
    main()
