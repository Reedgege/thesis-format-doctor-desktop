# -*- coding: utf-8 -*-
# Copyright (c) 2026 芦苇（山东大学 MBA）. All rights reserved.
# 原创作品 | 禁止未经授权转售、二次分发或抄袭 | 授权用户可在扣子/虾评平台内使用
# build: 202608191649 | version: 1.3.5 | file_sha256: 86a4000b6c4189f9
# -*- coding: utf-8 -*-
"""
format-profile.py — 学校论文格式模板 -> 格式画像 profile.json

用法：
  python format-profile.py 模板.docx [-o profile.json]

输出 JSON 结构（供 format-checker.py / headings-fix.py / ref-reformat.py 消费）：
{
  "source": "模板.docx",
  "headings": {            # 各级标题样式规范
     "1": {"styleId":"Heading1","name":"标题 1","font":"","size":"","firstLineChars":"","line":""},
     "2": {...}, "3": {...}
  },
  "headingStyles": {       # 按模板真实 styleId 套用的锚点（含批注覆盖后的字号/字体/大纲级别）
     "1": {"styleId":"...","name":"...","font":"黑体","size":"小三","outline_level":1,
            "from_comment": true, "outline_from_comment": true, "xml":"<w:style>...</w:style>"}
  },
  "body":   {"styleId":"Normal","font":"","size":"","firstLineChars":"","line":""},
  "levels": {                         # 完整 per-level 要素（样式定义+批注合并，批注优先）
     "1": {"zh_font":"黑体","en_font":"Times New Roman","sz":30,"bold":true,
            "align":"center","indent_type":"none","indent_chars":0,
            "line_rule":"single","line_val":240,"before_pt":24,"after_pt":18,"outline_level":1},
     "2": {...}, "3": {...}, "body": {"zh_font":"宋体","sz":24,"align":"both",
            "indent_type":"first","indent_chars":2,"line_rule":"exact","line_val":20}
  },
  "page":   {"top":"","bottom":"","left":"","right":""},   # 页边距(twips)
  "tables": {"threeLine": true/false, "note": "..."},      # 三线表规范判定
  "refExample": [ "...", "..." ],   # 模板里参考文献区的前几条原文，作范式样本
  "comment_count": 31,               # 模板批注条数
  "spec": {                          # 批注驱动权威规范（用户要求"以批注为准"）
     "authoritative_source": "comments",
     "page": {"top_cm":2.8,"bottom_cm":2.5,"left_cm":2.5,"right_cm":2.5},
     "h1": {"zh_font":"黑体","en_font":"Times New Roman","size":"小三","sz":30,
            "bold":true,"outline_level":1,"align":"center","before_pt":24,"after_pt":18,"line_rule":"single"},
     "h2": {...}, "h3": {...}, "body": {...}, "abstract":{...}, "keywords":{...},
     "toc":{...}, "reference":{...}, "figure":{...}, "table":{...}, "footnote":{...}, ...
  }
}

说明：本工具只读"样式定义 + 批注"，不改正文内容；符合"仅调样式、不改正文"的合规边界。
批注（comments.xml）中逐条写明的字号/字体/大纲级别/对齐/缩进/段间距/行距/页边距，
优先于样式定义被采纳（headingStyles 与 spec 均据批注覆盖），确保"以批注为准"。
"""
import sys
import json
import os
import re
import xml.etree.ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from docxutils import (load_styles, load, margins, tables, WR, detect_heading,
                       is_toc_residue, style_xml, comments, extract_comment_specs,
                       is_body_style_name, is_reference_heading)

__author__ = "芦苇"
__copyright__ = "Copyright (c) 2026 芦苇（山东大学 MBA）"
__version__ = "1.3.5"



_INTEGRITY_TAG = "Q29weXJpZ2h0IChjKSAyMDI2IOiKpuiLh++8iOWxseS4nOWkp+WtpiBNQkHvvIkuIEFsbCByaWdodHMgcmVzZXJ2ZWQuIHwgYnVpbGQ6MjAyNjA4MTkxNjQ5"  # noqa
def _style_level(sid, name):
    """根据 styleId / 样式名推断标题级别 1/2/3。"""
    for token in (sid or "", name or ""):
        m = re.search(r"(?:标题|Heading)\s*(\d)", token, re.I)
        if m:
            return int(m.group(1))
    return None


def _style_spec(info):
    """把样式定义（load_styles 的 {rPr, pPr}）转换为与批注 spec 同构的要素字典。

    用于「无批注 / 批注不全」的模板：也能从真实样式定义里挖出
    中文字体 / 西文字体 / 字号 / 加粗 / 对齐 / 首行缩进 / 段前 / 段后 / 行距
    等全部要素（用户要求各级别都注意这些要素）。山大等把格式写在批注里的
    模板，其样式定义多为空，这里返回 {}，再由批注 spec 兜底填充。
    """
    rpr = (info or {}).get("rPr") or {}
    ppr = (info or {}).get("pPr") or {}
    spec = {}
    if rpr.get("eastAsia"):
        spec["zh_font"] = rpr["eastAsia"]
    elif rpr.get("ascii"):
        spec["zh_font"] = rpr["ascii"]
    if rpr.get("hAnsi"):
        spec["en_font"] = rpr["hAnsi"]
    elif rpr.get("ascii"):
        spec["en_font"] = rpr["ascii"]
    if rpr.get("sz"):
        try:
            spec["sz"] = int(rpr["sz"])
        except (TypeError, ValueError):
            pass
    if rpr.get("bold") is not None:
        try:
            spec["bold"] = bool(rpr["bold"])
        except (TypeError, ValueError):
            pass
    if ppr.get("jc"):
        spec["align"] = ppr["jc"]
    # 首行缩进：firstLineChars 单位=百分之一字符；firstLine 单位=twips
    flc = ppr.get("firstLineChars")
    fl = ppr.get("firstLine")
    if flc:
        spec["indent_type"] = "first"
        try:
            spec["indent_chars"] = int(flc) // 100
        except (TypeError, ValueError):
            pass
    elif fl:
        spec["indent_type"] = "first"
        try:
            spec["indent_chars"] = int(fl) // 240
        except (TypeError, ValueError):
            pass
    # 行距：auto 以 240 分度为基准（240=单倍，360=1.5倍）；exact/atLeast 以 twips 计
    line = ppr.get("line")
    lr = ppr.get("lineRule")
    if line is not None:
        try:
            line = int(line)
        except (TypeError, ValueError):
            line = None
    if line is not None:
        if lr == "exact":
            spec["line_rule"] = "exact"
            spec["line_val"] = line // 20
        elif lr == "atLeast":
            spec["line_rule"] = "atLeast"
            spec["line_val"] = line // 20
        else:  # auto（单倍 / 1.5 倍 …）
            spec["line_rule"] = "auto"
            spec["line_val"] = line
    if ppr.get("before") is not None:
        try:
            spec["before_pt"] = int(ppr["before"]) // 20
        except (TypeError, ValueError):
            pass
    if ppr.get("after") is not None:
        try:
            spec["after_pt"] = int(ppr["after"]) // 20
        except (TypeError, ValueError):
            pass
    return spec


def _merge_specs(base, override):
    """批注(override)优先、样式定义(base)兜底，合并为完整 per-level 要素字典。"""
    out = dict(base or {})
    for k, v in (override or {}).items():
        if v is not None:
            out[k] = v
    return out


def extract(path):
    styles = load_styles(path)
    headings = {}
    body = None
    for sid, info in styles.items():
        nm = info.get("name") or ""
        # 标题层级
        if sid in ("Heading1", "Heading2", "Heading3") or re.search(r"标题\s*\d", nm):
            lvl = _style_level(sid, nm)
            if lvl and lvl <= 3:
                headings[str(lvl)] = {
                    "styleId": sid,
                    "name": nm,
                    "font": (info.get("rPr") or {}).get("eastAsia") or (info.get("rPr") or {}).get("ascii") or "",
                    "size": (info.get("rPr") or {}).get("sz") or "",
                    "firstLineChars": (info.get("pPr") or {}).get("firstLineChars") or "",
                    "line": (info.get("pPr") or {}).get("line") or "",
                }
        # 正文（跨校：用 is_body_style_name 识别中英文/异校正文样式名，不写死"正文"）
        if is_body_style_name(sid) or is_body_style_name(nm):
            body = {
                "styleId": sid,
                "name": nm,
                "font": (info.get("rPr") or {}).get("eastAsia") or (info.get("rPr") or {}).get("ascii") or "",
                "size": (info.get("rPr") or {}).get("sz") or "",
                "firstLineChars": (info.get("pPr") or {}).get("firstLineChars") or "",
                "line": (info.get("pPr") or {}).get("line") or "",
            }
    # 页边距
    _, root = load(path)
    pg = margins(root)
    # 三线表规范：取模板中第一个表格的边框判定
    tbls = tables(root)
    three_line = None
    note = "模板未检测到表格样例"
    if tbls:
        b = tbls[0]
        has_v = b.get("insideV") not in (None, "none") or b.get("left") not in (None, "none") or b.get("right") not in (None, "none")
        three_line = (not has_v) and b.get("top") not in (None, "none") and b.get("bottom") not in (None, "none")
        note = "模板首表判定为三线表" if three_line else "模板首表含竖线/左右边线（非标准三线表）"
    # 参考文献范式样本（跨校：用 is_reference_heading 定位，不写死中文"参考文献"）
    ref_example = []
    paras = []
    for p in root.iter(WR + "p"):
        txt = "".join(t.text or "" for t in p.iter(WR + "t"))
        paras.append(txt)
    ref_start_idx = next((i for i, tx in enumerate(paras) if is_reference_heading(tx)), None)
    if ref_start_idx is not None:
        seg = "\n".join(paras[ref_start_idx:ref_start_idx + 60])
        for line in seg.splitlines():
            line = line.strip()
            if re.match(r"\[\d+\]", line):
                ref_example.append(line)
    # ------------------------------------------------------------------
    # 批注驱动规范（用户明确要求：以批注写明的要求为准，优先于样式定义）
    # ------------------------------------------------------------------
    comment_specs = extract_comment_specs(comments(path))
    spec = _build_spec(comment_specs)

    # 用批注中的权威字号/字体/大纲级别覆盖 headingStyles（批注优先）
    hs = _heading_style_map(path, styles, load(path))
    for lvl in (1, 2, 3):
        hk = str(lvl)
        if hk not in hs:
            continue
        cspec = (spec.get("cats") or {}).get("h%d" % lvl)
        if not cspec:
            continue
        if cspec.get("size"):
            hs[hk]["size"] = cspec["size"]
            hs[hk]["from_comment"] = True
        if cspec.get("zh_font"):
            hs[hk]["font"] = cspec["zh_font"]
            hs[hk]["from_comment"] = True
        if cspec.get("outline_level") is not None:
            hs[hk]["outline_level"] = cspec["outline_level"]
            hs[hk]["outline_from_comment"] = True

    # ------------------------------------------------------------------
    # 完整 per-level 要素画像 levels：先「真实样式定义」挖全要素，再用批注
    # spec 覆盖（批注优先）。即使模板无批注 / 批注不全，也能从样式定义取到
    # 字体 / 字号 / 加粗 / 对齐 / 缩进 / 段前 / 段后 / 行距（用户要求各级别
    # 都注意这些要素）。headings-fix / format-checker 均优先消费 levels。
    # ------------------------------------------------------------------
    levels = {}
    for lvl in (1, 2, 3):
        hk = str(lvl)
        base = {}
        if hk in hs:
            base = _style_spec(styles.get(hs[hk].get("styleId"), {}))
        cs = (spec.get("cats") or {}).get("h%d" % lvl)
        merged = _merge_specs(base, cs)
        if cs and cs.get("outline_level") is not None:
            merged["outline_level"] = cs["outline_level"]
        levels[hk] = merged
    # 正文：从 Normal 样式定义挖全要素，再用批注 body 覆盖
    body_sid = (body or {}).get("styleId") or "Normal"
    levels["body"] = _merge_specs(_style_spec(styles.get(body_sid, {})),
                                  (spec.get("cats") or {}).get("body"))
    # 其余全部批注类别（表题/图题/脚注/摘要/参考文献等）一并纳入 levels，
    # 使 headings-fix 可通过 _lvl_spec(key) 统一读取并套用。
    _core = {"1", "2", "3", "body"}
    for ck, cv in (spec.get("cats") or {}).items():
        if ck in _core:
            continue
        levels[ck] = cv

    # 把 levels 的完整要素回填进 headings / headingStyles / body 字典（保留既有键，仅填空缺），
    # 使画像 JSON 在「各级别要素」上完整、可一览，且不破坏既有消费方（如 TC1）。
    for hk, ls in levels.items():
        if hk == "body":
            targets = [body]
        else:
            targets = [headings.get(hk), hs.get(hk)]
        for target in targets:
            if target is None:
                continue
            for k, v in ls.items():
                if v is not None and k not in target:
                    target[k] = v

    return {
        "source": os.path.basename(path),
        "headings": headings,
        "headingStyles": hs,
        "body": body,
        "levels": levels,
        "page": pg,
        "tables": {"threeLine": three_line, "note": note},
        "refExample": ref_example[:8],
        "spec": spec,
        "comment_count": comment_specs["count"],
    }


def profile_summary_block(profile):
    """把模板画像关键要素转成 Markdown 摘要块（模板驱动诊断/修改报告开头用）。

    直接回答"从学校模板提取到了什么"：模板名、批注条数、各级标题样式 ID 与
    字体字号、正文要求、页边距、参考文献。画像为空（无标准标题样式且批注未写
    明要求）时给出明确警告，避免"静默退化为通用规范"造成客户困惑。
    """
    if not profile:
        return ("⚠️ **未使用学校模板画像**：本次按【通用规范】处理。\n"
                "> 如需按学校模板驱动，请在第①步选择学校模板并重新提取。\n")
    L = []
    src = profile.get("source") or "（未知）"
    n_cmt = profile.get("comment_count", 0)
    hs = profile.get("headingStyles") or {}
    levels = profile.get("levels") or {}
    L.append(f"### 学校模板画像摘要（来源：`{src}`）")
    L.append(f"> 批注 {n_cmt} 条 ｜ 标题样式 {len(hs)} 级 ｜ per-level 要素 {len(levels)} 组")

    def _pt(v):
        try:
            fv = float(v)
        except (TypeError, ValueError):
            return v
        # 半磅（OOXML sz）转磅显示；20 以上视为半磅（小四=24 半磅等）
        return ("%.1f" % (fv / 2.0)).rstrip("0").rstrip(".") if fv > 20 else ("%g" % fv)

    _lvl_names = {"1": "一级标题", "2": "二级标题", "3": "三级标题"}
    for lvl in ("1", "2", "3"):
        h = hs.get(lvl) or {}
        ls = levels.get(lvl) or {}
        sid = h.get("styleId") or ""
        font = ls.get("zh_font") or ls.get("font") or h.get("font") or ""
        sz = ls.get("size") or ls.get("sz") or h.get("size") or ""
        al = ls.get("align") or ""
        _frag = []
        if sid:
            _frag.append("样式 `%s`" % sid)
        else:
            _frag.append("样式（未提取）")
        _frag.append("字体 %s" % (font or "—"))
        _frag.append("字号 %s" % (_pt(sz) if sz else "—"))
        if al:
            _frag.append("对齐 %s" % al)
        L.append("- %s L%s：%s" % (_lvl_names.get(lvl, lvl), lvl, " ｜ ".join(_frag)))
    b = levels.get("body") or {}
    _bf = []
    _bf.append("字体 %s" % (b.get("zh_font") or b.get("font") or "—"))
    _bf.append("字号 %s" % (_pt(b.get("size") or b.get("sz")) if (b.get("size") or b.get("sz")) else "—"))
    _bf.append("首行缩进 %s" % (b.get("indent_chars") or b.get("firstLineChars") or "—"))
    _bf.append("行距 %s" % (b.get("line_val") or b.get("line") or "—"))
    L.append("- 正文：%s" % " ｜ ".join(_bf))
    pg = profile.get("page") or {}
    if pg:
        def _cm(v):
            try:
                return "%.2fcm" % (float(v) / 567.0)
            except (TypeError, ValueError):
                return str(v)
        L.append("- 页边距：上 %s ｜ 下 %s ｜ 左 %s ｜ 右 %s" % (
            _cm(pg.get("top")), _cm(pg.get("bottom")), _cm(pg.get("left")), _cm(pg.get("right"))))
    refs = profile.get("refExample") or []
    _has_ref = bool(refs) or bool(levels.get("reference")) or bool((profile.get("spec") or {}).get("reference"))
    L.append("- 参考文献格式：%s" % ("已提取（%d 条例）" % len(refs) if _has_ref else "未提取"))
    if not hs and not levels:
        L.append("")
        L.append("⚠️ **该模板未提取到格式要求**（无标准标题样式，且批注未写明要求）。")
        L.append("> 本次修正/检查将退化为【通用规范】，可能无法满足学校精确要求。")
        L.append("> 建议：① 确认模板是 Word 文档且内含标题样式或批注；② 换用学校官方模板文件重试。")
    return "\n".join(L) + "\n"


def _build_spec(comment_specs):
    """把 extract_comment_specs 的输出整理为 profile 的 spec 段（批注为准）。"""
    cats = comment_specs.get("cats", {})
    page = comment_specs.get("page")
    spec = {
        "authoritative_source": "comments",
        "comment_count": comment_specs.get("count", 0),
        "note": "以下规范来自模板批注（用户明确要求以批注写明的要求为准），"
                "已覆盖并优先于样式定义中的同名项。",
        "page": page,
        "cats": cats,
        # 便捷别名：诊断/套样式时高频使用
        "h1": cats.get("h1"), "h2": cats.get("h2"), "h3": cats.get("h3"),
        "body": cats.get("body"),
        "abstract": cats.get("abstract"),
        "keywords": cats.get("keywords"),
        "toc": cats.get("toc"),
        "reference": cats.get("reference"),
        "reference_heading": cats.get("reference_heading"),
        "figure": cats.get("figure"),
        "table": cats.get("table"),
        "footnote": cats.get("footnote"),
        "title": cats.get("title"),
        "conclusion": cats.get("conclusion"),
        "ack": cats.get("ack"),
        "appendix": cats.get("appendix"),
        "en_abstract": cats.get("en_abstract"),
        "en_keywords": cats.get("en_keywords"),
    }
    return spec


def _is_heading_style_id(sid, info):
    """判断 styleId / 样式名是否是真正的标题样式（而非 Normal/正文/目录等）。

    判定标准（满足任一即可）：
      - 样式名包含 "heading" 或 "标题"（不区分大小写）
      - 样式 ID 为 "Heading1"/"Heading2"/"Heading3"（大小写敏感的 OOXML 标准 ID）
    """
    name = (info or {}).get("name") or ""
    if re.search(r"heading|标题", name, re.I):
        return True
    if sid in ("Heading1", "Heading2", "Heading3"):
        return True
    return False


def _heading_style_map(path, styles, loaded):
    """扫描模板正文，按层级取最常用的标题 styleId（排除目录域残留/目录样式），并附该样式定义 XML。

    返回 {str(level): {"styleId","name","font","size","xml"}}。这是「按模板应用样式」的锚点：
    客户论文的对应层级标题将被套用模板这一确切 styleId，而非通用 HeadingN。

    v1.3.3 修复：
      - 排除目录(TOC/Contents)样式段落，避免目录条目数量压过真实标题导致 most_common
        选成目录样式（山大模板实测踩坑）。
      - 当模板正文一级/二级标题使用 Normal 等非标题样式时，不将其作为标题 styleId 输出，
        交由 headings-fix 回退到内置 Heading1/2/3（Normal 不是标题样式，套上去会污染正文）。
    """
    from collections import Counter
    _, root = loaded
    counter = {1: Counter(), 2: Counter(), 3: Counter()}
    for p in root.iter(WR + "p"):
        txt = "".join(t.text or "" for t in p.iter(WR + "t")).strip()
        if is_toc_residue(txt):
            continue
        lvl, _ = detect_heading(txt)
        if lvl is None or lvl > 3:
            continue
        ppr = p.find(WR + "pPr")
        sid = ppr.find(WR + "pStyle").get(WR + "val") if (ppr is not None and ppr.find(WR + "pStyle") is not None) else None
        if sid:
            # 排除目录样式：样式名含 TOC/目录 的不算真实标题
            sinfo = styles.get(sid, {})
            sname = (sinfo.get("name") or "")
            sname_lower = sname.lower()
            sid_lower = sid.lower()
            if ("toc" in sname_lower or "目录" in sname
                    or sid_lower.startswith("toc")
                    or sid_lower.startswith("contents")):
                continue
            counter[lvl][sid] += 1
    out = {}
    for lvl in (1, 2, 3):
        if not counter[lvl]:
            continue
        sid = counter[lvl].most_common(1)[0][0]
        info = styles.get(sid, {})
        # 非标题样式（如 Normal/正文）不能作为标题 styleId——否则客户论文所有
        # 该级标题被套成 Normal，污染正文；跳过该级别，由 headings-fix 回退内置 Heading。
        if not _is_heading_style_id(sid, info):
            continue
        out[str(lvl)] = {
            "level": lvl,
            "styleId": sid,
            "name": info.get("name") or sid,
            "font": (info.get("rPr") or {}).get("eastAsia") or (info.get("rPr") or {}).get("ascii") or "",
            "size": (info.get("rPr") or {}).get("sz") or "",
            "xml": style_xml(path, sid),
        }
    return out


def main():
    if len(sys.argv) < 2:
        print("用法: python format-profile.py 模板.docx [-o profile.json]")
        sys.exit(1)
    path = sys.argv[1]
    out = None
    if "-o" in sys.argv:
        out = sys.argv[sys.argv.index("-o") + 1]
    prof = extract(path)
    text = json.dumps(prof, ensure_ascii=False, indent=2)
    if out:
        with open(out, "w", encoding="utf-8") as f:
            f.write(text)
        print(f"已写出格式画像 -> {out}")
    else:
        print(text)


if __name__ == "__main__":
    main()
