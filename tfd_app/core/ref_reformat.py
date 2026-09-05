# -*- coding: utf-8 -*-
# Copyright (c) 2026 芦苇不熬夜. All rights reserved.
# 原创作品 | 禁止未经授权转售、二次分发或抄袭 | 授权用户可在扣子/虾评平台内使用
# build: 202609051150 | version: 1.3.8 | file_sha256: 7738cb6a9a2019e0
# -*- coding: utf-8 -*-
"""
ref-reformat.py — 参考文献按 GB/T 7714 范式重排（逐条挂 Word 批注）

用法：
  python ref-reformat.py 输入.docx [-o 输出.docx]

能力：
  - 抽取"参考文献"区条目（[n] 标记，**编号保留原文、断档不强制重排**——
    正文引用与文后编号的对应关系属"内容语义"，引擎不擅动）
  - 解析字段：作者 / 题名 / 文献类型 / 出处 / 出版年（GB/T 7714 全类型码，含 [C]// 会议论文集）
  - 可解析条目按 GB/T 7714 范式排布（标点/空格/作者污染归一），缺项不编造；
    低置信条目仅做安全标点归一，无法确认的内容绝不擅动
  - **凡被引擎修改过的条目，一律在该条目上挂 Word 批注说明**（作者
    "论文格式医生·提示"，Word 中可见、可删）：干净条目挂"已按 GB/T 7714
    规范化标点与空格"告知改动；需人工核对的条目（缺年份 / 解析不确定）
    挂具体提示。**条目正文本身干净、不带任何 ⚠ 等工具文字** —— 符合
    "只改格式、绝不动内容、可加批注"
  - 写回新 docx（仅替换参考文献区文字，不碰正文其余部分）

说明：中文文献格式杂，纯规则解析必有误差；本工具只做有把握的格式归一，
凡有改动逐条以批注留痕、拿不准的以批注请人工核对，绝不把工具提示写进正文。
"""
import sys
import os
import re
import xml.etree.ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import docxutils
from docxutils import load, write_docx_files, to_doc_xml, WR

__author__ = "芦苇"
__copyright__ = "Copyright (c) 2026 芦苇不熬夜"
__version__ = "1.3.8"

# 干净条目（可解析、无 warn）被引擎改写后，挂到该条目上的统一告知文案
_CLEAN_NOTE = "此条已按 GB/T 7714 规范化标点与空格，文献内容未改动（如需逐处对照见修改明细）"



_INTEGRITY_TAG = "Q29weXJpZ2h0IChjKSAyMDI2IOiKpuiLh+S4jeeGrOWknC4gQWxsIHJpZ2h0cyByZXNlcnZlZC4gfCBidWlsZDoyMDI2MDgxOTE2NDk="  # noqa
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
        # 防御：参考文献区后面跟着的"章节标题"（无标题样式时），不能吞入上一条
        # 判定：去掉空白后是「编号 + 结构词」或「结构词」开头的短行（标题短，文献条目长）
        flat = re.sub(r"\s+", "", (p["text"] or "").strip())
        if flat and len(flat) <= 24 and re.match(
                r"^(?:第?[0-9一二三四五六七八九十百]+[\.、\s]*)?(?:结论|致谢|附录|目录|摘要|"
                r"abstract|conclusion|acknowledg|appendix|reference|bibliograph)",
                flat, re.I):
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


def _clean(text):
    """安全格式归一（不动内容语义）：
    - 全角/半角空白压缩为单个半角空格
    - 清理 ",." / "，." / "、." 这类逗号后句点污染（作者缩写解析 bug 来源）
    - 去掉行尾冗余句点（重排输出时会统一补回）
    """
    if not text:
        return ""
    text = re.sub(r"\s+", " ", text.replace("\n", " ")).strip()
    text = re.sub(r"([,，、;；:：])\s*[.。]+", r"\1", text)
    return text.strip()


def _clean_author(author):
    """作者净化：归一空白 + 逗号前不留空格（Smith, X -> Smith, X 保留，但 'Chen , X' -> 'Chen, X'）。"""
    a = _clean(author)
    a = re.sub(r"\s+([,，;；])", r"\1", a)
    a = re.sub(r"([,，;；])\s+", r"\1 ", a)
    return a.strip()


def parse_ref(raw):
    """尽力解析单条参考文献，返回字段字典 + 置信度。

    以文献类型标记 [JMDCN] 为锚点三段切分（比旧版只认"作者.题名"更稳）：
      头部(作者.题名) + [类型] + 尾部(出处, 年.)
    无法可靠切分时返回低置信，交由调用方做"安全归一 + 批注提示"，绝不往正文写 ⚠。
    """
    text = re.sub(r"\s+", " ", (raw or "").replace("\n", " ")).strip()
    res = {"author": "", "title": "", "type": "", "source": "", "year": "", "joiner": ""}
    conf = 0.5
    # 类型标记：GB/T 7714 全类型码（J期刊 M专著 C会议 D学位 N报纸 R报告 S标准 P专利
    # G汇编 Z其他 K词典 EB电子文献，及复合 /OL /DB /CP 等）；
    # 会议论文集格式 [C]// 论文集 的 "//" 引导符一并捕获，保持国标连写不被拆散。
    tm = re.search(r"\[((?:J|M|C|D|N|R|S|P|G|Z|K|EB)(?:/[A-Z]{1,3})*)\][ ]*(//)?", text)
    if tm:
        res["type"] = tm.group(1)
        res["joiner"] = tm.group(2) or ""
        conf += 0.15
        head = text[:tm.start()].strip()
        tail = text[tm.end():].strip()
    else:
        head, tail = text, ""
    # 年：只在"类型标记之后的出处段"找（避免把题名里的年份如"2025年中国…"误当出版年）；
    # 无类型标记时才回退整段扫描辅助判断。
    ym = re.search(r"(?:19|20)\d{2}", tail if tm else text)
    if ym:
        res["year"] = ym.group(0)
        conf += 0.15
    # 切分头部：以"句点"为 作者 | 题名 分界。
    # 分界点判定：句点后跟（0~n 空白 +）一个【非小写字母、非空白、非标点】字符才切——
    #   中文标准式"张三,李四.基于…"（句点后无空格）能切；
    #   英文作者缩写 "X. and Liu"（句点后接 and 小写）不切、连写姓名缩写 "Y. The"(T大写) 正常切。
    if head:
        m = re.split(r"[.。](?=\s*[^a-z\s.。])", head, maxsplit=1)
        if len(m) == 2:
            a, t = m[0].strip(), m[1].strip()
            if a and t:
                res["author"] = _clean_author(a)
                res["title"] = t.rstrip(".。 ")
                conf += 0.15
        else:
            # 没有句点分界（如"张三 基于…"或整段是题名），整段视为题名、作者留空
            res["title"] = _clean(head)
    if tail:
        # 尾部 = 出处（含年、卷期页码、出版地）；中文条目常带"出处: 地点"需处理
        res["source"] = _clean(tail)
    if res["author"] and res["title"] and res["type"]:
        conf = min(conf + 0.2, 1.0)
    elif res["title"] and res["type"]:
        conf = min(conf + 0.1, 1.0)
    return res, round(conf, 2)


def _fallback_normalize(raw):
    """低置信条目的安全归一：压缩空白、清理 ",." 污染、编号后统一单空格、
    句点后接 ASCII 字母/数字时补一个空格（排版归一，不动内容语义）。

    原则：能确定的格式问题（空格/标点/编号样式）直接改；内容解析不确定的保留原文。
    正文不带任何提示文字 —— 需人工核的条目由批注提示，绝不往正文写 ⚠。
    """
    t = _clean(raw)
    # 句点后紧跟英文/数字（如 "corporation.Harvard" -> "corporation. Harvard"）
    t = re.sub(r"([.。])(?=[A-Za-z0-9])", r"\1 ", t)
    return t.strip()


def reformat(refs):
    """按 GB/T 7714 范式归一文本，**编号保留原文（断档不重排）**。

    正文输出必须干净（绝不含 ⚠ 等工具标记）；确需人工处理的条目
    （缺年份 / 解析不确定）通过 `warn` 字段传出，由写回层在对应条目上挂 Word 批注。

    返回每条：{num(原文编号), old, new(正文干净文本), conf, warn(批注提示或 None)}
    """
    out = []
    for r in refs:
        fields, conf = parse_ref(r["raw"])
        joiner = (fields.get("joiner") or "").strip()
        num = r["num"]  # 保留原文编号，不连续化、不重排
        warn = None
        if fields["author"] and fields["type"]:
            src = fields["source"].strip().lstrip(".。 ").rstrip(".。 ")
            # 会议论文集 [C]// 论文集 连写；其余 [J]. 出处 用句点分隔
            head = f"{fields['author']}. {fields['title']}[{fields['type']}]{joiner}"
            if fields["year"]:
                if src and not re.search(r"(?:19|20)\d{2}", src):
                    src = f"{src}, {fields['year']}"
                new = f"[{num}] {head}. {src}." if not joiner else f"[{num}] {head} {src}."
            else:
                # 可解析但缺年份：结构照排，年份不编造，提示走批注
                new = f"[{num}] {head}. {src}." if not joiner else f"[{num}] {head} {src}."
                warn = "此条缺出版年份，请人工核对补全（引擎不编造年份）"
        else:
            # 低置信：仅做安全格式归一（空白/标点清理，不动内容语义），
            # 正文不带任何提示文字，无法确认处请人工核（走批注）
            new = f"[{num}] {_fallback_normalize(r['raw'])}"
            warn = "此条格式解析不确定，已做基本标点/空格归一，请人工核对"
        out.append({"num": num, "old": r["raw"], "new": new,
                    "conf": conf, "warn": warn})
    return out


def _render(results):
    lines = ["# 参考文献重排报告（GB/T 7714）", ""]
    for r in results:
        tag = "✅" if not r.get("warn") else "⚠"
        lines.append(f"{tag} [{r['num']}] (置信度 {r['conf']})")
        lines.append(f"   原：{r['old'][:60]}")
        lines.append(f"   新：{r['new'][:70]}")
        if r.get("warn"):
            lines.append(f"   提示：{r['warn']}（已挂 Word 批注，正文未写入该文字）")
        lines.append("")
    return "\n".join(lines)


def _replace_in_docx(src, dst, results):
    """把参考文献区 [n] 段落文字写回为归一结果；**凡被改写的条目都挂 Word 批注**。

    - 编号保留原文（不重排），按文档顺序一一对应消费，不会错位
    - 正文写入**干净文本**，绝不含 ⚠/提示文字（符合"只改格式、绝不动内容"）
    - 批注策略（每条至多一条，作者"论文格式医生·提示"，Word 可见可删）：
        * 带 warn 的条目（缺年份/解析不确定）→ 挂对应人工核对提示（无条件挂，
          即使文本恰好没变也提示，因为该条需要人去处理）
        * 无 warn 但条目文本被引擎改写（`原段落 != 新正文`）→ 挂"已按 GB/T 7714
          规范化"告知文案，让改动逐条可查
        * 无 warn 且文本完全未变（条目本已规范）→ 不挂批注，避免无意义噪音
    - 复用 docxutils 的 comments.xml 基础设施（ID 自动接续文档现有批注）
    """
    z, root = load(src)
    try:
        replacements = {}
        in_ref = False
        pend = []  # (段落 Element, 批注文本)
        idx = 0
        for p in root.iter(WR + "p"):
            txt = "".join(t.text or "" for t in p.iter(WR + "t"))
            if docxutils.is_reference_heading(txt):
                in_ref = True
                continue
            if not in_ref:
                continue
            m = re.match(r"^\s*\[(\d+)\]", txt)
            if m and idx < len(results):
                res = results[idx]
                # 判定"是否真改写"：忽略编号前缀（编号格式由引擎统一重写为 [n] ，
                # 不算条目内容改动），只比较条目主体文本
                old_body = re.sub(r"^\s*\[\d+\]\s*", "", txt, count=1).strip()
                new_body = re.sub(r"^\s*\[\d+\]\s*", "", res["new"], count=1).strip()
                changed = old_body != new_body
                if res.get("warn") or changed:
                    _set_para_text(p, res["new"])
                    note = res.get("warn") or _CLEAN_NOTE
                    pend.append((p, note))
                idx += 1
        if pend:
            comments_root = docxutils.ensure_comments_part(z, replacements)
            next_id = docxutils.max_comment_id(z) + 1
            for p, warn in pend:
                docxutils.add_comment_marker(p, next_id, warn, comments_root,
                                             author="论文格式医生·提示", initials="TS")
                next_id += 1
            replacements["word/comments.xml"] = to_doc_xml(comments_root)
        replacements["word/document.xml"] = to_doc_xml(root)
        write_docx_files(src, dst, replacements)
    finally:
        z.close()


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
