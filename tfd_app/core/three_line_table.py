# -*- coding: utf-8 -*-
# 文档格式先生 · 学术三线表自动转换（纯标准库，零三方依赖）
#
# 移植自海外版 thesis-format-doctor-global 的
#   src/engine/fixer.py   : _apply_three_line_tables / _tbl_is_three_line
#   src/engine/docx_reader.py: sample_table_lines / _effective_border 等
# 用 xml.etree.ElementTree 重写，以保国内版「零第三方依赖」底线。
#
# 三线表定义（学术规范，与海外版同一口径）：
#   ① 有顶线  ② 有底线  ③ 表头下方一条栏目线（首行各格 bottom）
#   ④ 无竖线（left / right / insideV）  ⑤ 无内部横线（insideH）
# 单行 / 1×1 布局表（rows<2）无法形成表头线，不判也不修（与 checker 同一口径）。
#
# 铁律（移植时一处都不能省）：
#   · 单元格级 tcBorders 优先级高于表格级 tblBorders —— 不清就是白改，必须逐格删残留。
#   · OOXML 子元素顺序：tblBorders 排 shd/tblLayout… 前；tcBorders 排 shd/noWrap…
#     前，否则 Word 报「文件已损坏」。_ordered_child 专门管这个顺序。
#   · 幂等：已是三线表的表跳过，保证第二遍零动作（否则报告红杠永远点不掉）。
#   · 三层边框解析：单元格 > 表格 > 表格样式（tblStyle）—— 漏看任一层都会误判。
import zipfile
import xml.etree.ElementTree as ET

from docxutils import WR, reg_ns

_THICK_SZ = 12   # w:sz 单位是 1/8 磅 → 12 = 1.5 磅（顶线 / 底线）
_THIN_SZ = 6     # 6 = 0.75 磅（表头栏目线）

# CT_TblPrBase 子元素顺序：tblBorders 必须排在 shd/tblLayout/… 之前
_TBLPR_AFTER_BORDERS = ("shd", "tblLayout", "tblCellMar", "tblLook",
                        "tblCaption", "tblDescription")
# CT_TcPr 子元素顺序：tcBorders 必须排在 shd/noWrap/tcMar/… 之前
_TCPR_AFTER_BORDERS = ("shd", "noWrap", "tcMar", "textDirection",
                       "tcFitText", "vAlign", "hideMark")

# CT_TcBorders / CT_TblBorders 内部子元素的 schema 规定顺序
_TBL_BORDER_ORDER = ("top", "start", "left", "bottom", "end", "right",
                     "insideH", "insideV")
_TCPR_BORDER_ORDER = _TBL_BORDER_ORDER + ("tl2br", "tr2bl")

# 🔴 斜线表头的对角线：这是**表头语义**（"内部/外部环境 × 优势/劣势"），
# 不是表格框线。三线表只该去掉竖线与内部横线，绝不能把对角线一起删掉 ——
# 否则同一批文档里「画成形状的斜线留着、设成对角边框的斜线没了」，客户看到的就是"有的在有的不在"。
_DIAGONAL_SIDES = ("tl2br", "tr2bl")


def _q(name):
    return WR + name


def _load_styles_root(src_path):
    """读取 word/styles.xml 的 ET root（供三层边框解析）；读不到返回 None。"""
    try:
        with zipfile.ZipFile(src_path) as z:
            try:
                data = z.read("word/styles.xml")
            except KeyError:
                return None
        return ET.fromstring(data.decode("utf-8"))
    except Exception:
        return None


def _ordered_child(parent, tag, after_tags):
    """按 OOXML 规定顺序取（必要时新建）子元素，避免 Word 报「文件已损坏」。"""
    el = parent.find(_q(tag))
    if el is not None:
        return el
    el = ET.Element(_q(tag))
    after_set = {_q(t) for t in after_tags}
    idx = len(parent)
    for i, child in enumerate(parent):
        if child.tag in after_set:
            idx = i
            break
    parent.insert(idx, el)
    return el


def _sort_border_box(box):
    """按 schema 顺序重排边框容器子元素。

    保留对角线（tl2br/tr2bl）后，新写进去的 bottom 会被追加到它们后面，
    而 CT_TcBorders 规定 tl2br/tr2bl 必须排在最后 —— 顺序一错 Word 就报「文件已损坏」。
    """
    order = _TCPR_BORDER_ORDER if box.tag == _q("tcBorders") else _TBL_BORDER_ORDER
    rank = {_q(t): i for i, t in enumerate(order)}
    children = list(box)
    children.sort(key=lambda e: rank.get(e.tag, 999))
    for ch in children:
        box.remove(ch)
    for ch in children:
        box.append(ch)


def _write_border(box, side, val, sz):
    """在 tblBorders / tcBorders 容器写某一边（先删后建，保证属性干净）。"""
    old = box.find(_q(side))
    if old is not None:
        box.remove(old)
    el = ET.SubElement(box, _q(side))
    el.set(_q("val"), val)
    if val not in ("nil", "none"):
        el.set(_q("sz"), str(sz))
        el.set(_q("space"), "0")
        el.set(_q("color"), "auto")
    _sort_border_box(box)


def _read_border(box, side):
    """从 tblBorders / tcBorders 容器读某边：返回 (val, sz)；无该边返回 None。"""
    if box is None:
        return None
    el = box.find(_q(side))
    if el is None:
        return None
    try:
        sz = int(el.get(_q("sz")) or 0)
    except Exception:
        sz = 0
    return (el.get(_q("val")) or "single", sz)


def _border_visible(border):
    """边框是否真的画得出来（nil / none = 不画）。"""
    return bool(border) and str(border[0]).lower() not in ("nil", "none")


def _table_style_borders(styles_root, tblPr):
    """表格引用样式（tblStyle）里定义的 tblBorders；没有返回 None。"""
    if styles_root is None or tblPr is None:
        return None
    st = tblPr.find(_q("tblStyle"))
    if st is None:
        return None
    sid = st.get(_q("val"))
    if not sid:
        return None
    for style in styles_root.findall(_q("style")):
        if style.get(_q("styleId")) == sid:
            spPr = style.find(_q("tblPr"))
            if spPr is not None:
                return spPr.find(_q("tblBorders"))
    return None


def _row_ex_box(row_el):
    """行级 w:tblPrEx 里的 tblBorders（Table Property Exceptions）。

    🔴 2026-11-01 血泪坑：Word 的三层优先级其实是
        单元格 tcPr/tcBorders  >  行 tblPrEx/tblBorders  >  表格 tblPr/tblBorders  >  样式
    从 Word 转存/老模板来的文档，每一行 w:tr 里都带一个 tblPrEx 把整张表的框线又写了一遍
    全边框（single sz=4）。只改表格级 tblBorders=nil，会被行级这份直接盖掉 ——
    XML 看着是三线表，Word 里打开还是满格框线（客户原话："表格并没有改成三线表"）。
    """
    ex = row_el.find(_q("tblPrEx"))
    if ex is None:
        return None
    return ex.find(_q("tblBorders"))


def _effective_border(tblPr, style_box, tcPr, side, row_box=None):
    """有效边框：单元格 → 行级 tblPrEx → 表格 → 表格样式，第一个有定义的就赢。"""
    b = _read_border(tcPr.find(_q("tcBorders")) if tcPr is not None else None, side)
    if b is not None:
        return b
    b = _read_border(row_box, side)
    if b is not None:
        return b
    b = _read_border(tblPr.find(_q("tblBorders")) if tblPr is not None else None, side)
    if b is not None:
        return b
    return _read_border(style_box, side)


def _row_tc_prs(row_el):
    """一行里每格的 tcPr（合并单元格会重复返回同一 tc，这里去重）。"""
    out, seen = [], set()
    for tc in row_el.findall(_q("tc")):
        if tc in seen:
            continue
        seen.add(tc)
        out.append(tc.find(_q("tcPr")))
    return out


def _analyze_table(tbl, styles_root):
    """采样单表边框现状。rows<2（单行/1×1 布局表）返回 None。

    ⚠️ 检查器的 sample_table_lines 与修正器的幂等判定 _table_is_three_line
    必须走**同一份**实现：此前是两段复制粘贴的逻辑，改一处忘一处就会出现
    「检查说已经不是问题、Word 里却还是满格框线」这种自伤。
    """
    tblPr = tbl.find(_q("tblPr"))
    rows = tbl.findall(_q("tr"))
    if len(rows) < 2:
        return None
    style_box = _table_style_borders(styles_root, tblPr)
    first = _row_tc_prs(rows[0])
    last = _row_tc_prs(rows[-1])
    ex_first = _row_ex_box(rows[0])
    ex_last = _row_ex_box(rows[-1])
    has_top = any(_border_visible(_effective_border(tblPr, style_box, p, "top", ex_first))
                  for p in first)
    has_bottom = any(_border_visible(_effective_border(tblPr, style_box, p, "bottom", ex_last))
                     for p in last)
    has_vertical = False
    has_inside_h = False
    for row in rows:
        exb = _row_ex_box(row)
        for p in _row_tc_prs(row):
            if (_border_visible(_effective_border(tblPr, style_box, p, "left", exb))
                    or _border_visible(_effective_border(tblPr, style_box, p, "right", exb))
                    or _border_visible(_effective_border(tblPr, style_box, p, "insideV", exb))):
                has_vertical = True
            if _border_visible(_effective_border(tblPr, style_box, p, "insideH", exb)):
                has_inside_h = True
    # 表头行**任一**单元格有 bottom 即算有表头线（合并单元格常只给部分格写）。
    has_header_line = bool(first) and any(
        _border_visible(_effective_border(tblPr, style_box, p, "bottom", ex_first))
        for p in first)
    return {
        "rows": len(rows),
        "cols": len(first),
        "has_top": has_top,
        "has_bottom": has_bottom,
        "has_header_line": has_header_line,
        "has_vertical": has_vertical,
        "has_inside_h": has_inside_h,
        "three_line": bool(has_top and has_bottom and not has_vertical
                           and not has_inside_h and has_header_line),
    }


def _table_is_three_line(tbl, styles_root):
    """单表是否已是学术三线表（与 sample_table_lines 同一口径）。用于幂等跳过。"""
    s = _analyze_table(tbl, styles_root)
    return bool(s and s["three_line"])


def sample_table_lines(root, styles_root=None):
    """采样每张表的边框现状，供 checker 判「是不是学术三线表」。

    三线表 = ① 有顶线 ② 有底线 ③ 表头下方有栏目线（≥2 行时）
             ④ 无竖线（left / right / insideV）  ⑤ 无内部横线（insideH）。
    每张表返回一个 dict；读不出来就少给一条，绝不臆造。
    """
    reg_ns()
    out = []
    for idx, tbl in enumerate(root.iter(_q("tbl")), 1):
        s = _analyze_table(tbl, styles_root)
        if s is None:
            # 单行（含 1×1 布局 / 摘要框）无法形成「表头线」，不属于数据表。
            continue
        s["index"] = idx
        out.append(s)
    return out


def apply_three_line_tables(root, styles_root=None):
    """把所有数据表改成学术三线表；返回 (改动表数, [每张表的 {index,rows,cols}])。

    ① 表格级：顶线 / 底线 1.5 磅，四边竖线与内部横线一律 nil；
    ② 逐格：先清单元格级 tcBorders 残留（优先级高于表格级，不清就是白改）；
    ③ 表头栏目线：首行各格的下边框 0.75 磅。
    单行 / 1×1 布局表（rows<2）跳过；已是三线表的表跳过（幂等）。
    """
    reg_ns()
    updated = 0
    changed = []
    for tbl in root.iter(_q("tbl")):
        tblPr = tbl.find(_q("tblPr"))
        if tblPr is None:
            continue
        rows = tbl.findall(_q("tr"))
        # 单行 / 布局表（含 1×1 摘要框）不算数据表，跳过（与 checker 同一口径）。
        if len(rows) < 2:
            continue
        # 已经是三线表 → 跳过，保证幂等（第二遍零动作）。
        if _table_is_three_line(tbl, styles_root):
            continue
        # ① 表格级：顶线 / 底线 1.5 磅，四边竖线与内部横线一律 nil
        box = _ordered_child(tblPr, "tblBorders", _TBLPR_AFTER_BORDERS)
        _write_border(box, "top", "single", _THICK_SZ)
        _write_border(box, "bottom", "single", _THICK_SZ)
        for _side in ("left", "right", "insideH", "insideV"):
            _write_border(box, _side, "nil", 0)
        # ② 行级 tblPrEx：清掉每一行里重复定义的全边框
        #    🔴 优先级高于表格级！Word 转存的文档每行都带 tblPrEx/tblBorders=全框线，
        #    不清它，上面表格级写了 nil 也是白改（Word 里照样满格框线）。
        for row in rows:
            ex = row.find(_q("tblPrEx"))
            if ex is None:
                continue
            exb = ex.find(_q("tblBorders"))
            if exb is not None:
                ex.remove(exb)
            if len(ex) == 0:
                row.remove(ex)   # tblPrEx 只剩空壳就没必要留着
        # ②b 逐格：清掉单元格级框线残留（优先级高于表格级，不清就是白改）；
        #     但**对角线 tl2br/tr2bl 必须原样留着** —— 那是斜线表头的表头语义，
        #     不是框线。整块 remove 会把客户的斜线表头删掉，而且删得"看运气"
        #     （形状式斜线删不到、边框式斜线全没），客户就会问"为什么有的在有的不在"。
        for row in rows:
            for tc in row.findall(_q("tc")):
                tcPr = tc.find(_q("tcPr"))
                if tcPr is None:
                    continue
                old = tcPr.find(_q("tcBorders"))
                if old is None:
                    continue
                diag_tags = {_q(s) for s in _DIAGONAL_SIDES}
                if any(e.tag in diag_tags for e in old):
                    for e in list(old):
                        if e.tag not in diag_tags:
                            old.remove(e)
                    _sort_border_box(old)
                else:
                    tcPr.remove(old)
        # ③ 表头栏目线：第一行各格的下边框 0.75 磅（≥2 行才画，单行表谈不上表头）
        if len(rows) >= 2:
            for tc in rows[0].findall(_q("tc")):
                tcPr = tc.find(_q("tcPr"))
                if tcPr is None:
                    # tcPr 必须是 tc 的第一个子元素，否则 OOXML 非法
                    tcPr = ET.Element(_q("tcPr"))
                    tc.insert(0, tcPr)
                cbox = _ordered_child(tcPr, "tcBorders", _TCPR_AFTER_BORDERS)
                _write_border(cbox, "bottom", "single", _THIN_SZ)
        updated += 1
        changed.append({"index": len(changed) + 1, "rows": len(rows),
                        "cols": len(rows[0].findall(_q("tc"))) if rows else 0})
    return updated, changed


__all__ = ["sample_table_lines", "apply_three_line_tables", "_load_styles_root"]
