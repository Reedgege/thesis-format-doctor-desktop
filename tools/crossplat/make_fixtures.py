# -*- coding: utf-8 -*-
"""生成「结构完整、格式全错」的中文毕业论文素材，供三平台功能验收。

为什么不用手写 XML 的极简包
--------------------------
`tests/test_headless_regression.py` 里那份 build_thesis 是手写 XML 拼的，
styles.xml 只声明了 Normal 和空的 Heading1 —— 拿它验收会得出误导性结论：
修正器把段落指向 Heading1/2/3，可目标文档里压根没有这几条样式定义，
Word 打开只会显示成"正文"，看着就像"改了没用"。真实客户的论文是 Word
存的，样式表齐全。素材不自洽，验出来的问题就不是产品的问题。

结构（客户真实投稿的样子）
    摘要 / 关键词 / 1 绪论 / 1.1 研究背景 / 2 文献综述 / 2.1 已有研究 /
    3 研究方法 / 正文多段 / 一张带框线的表 / 参考文献 3 条 / 致谢

故意不合规的点（模板要求见 PROFILE）
    · 正文字号五号(10.5pt)，要求小四(12pt)
    · 行距单倍，要求固定 20 磅
    · 首行不缩进，要求 2 字符
    · 标题全是手打纯文本（Normal、不加粗），没有套任何标题样式
    · 表格用 Table Grid（有框线），要求学术三线表
    · 封面缺题目/作者/专业

用法：python tools/crossplat/make_fixtures.py [输出目录]
"""

from __future__ import annotations

import json
import os
import sys

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

HERE = os.path.dirname(os.path.abspath(__file__))
OUT_DIR = os.path.join(HERE, "fixtures")

PROFILE = {
    "source": "学校模板（画像）",
    "levels": {
        "abstract": {"zh_font": "宋体", "en_font": "Times New Roman", "sz": "22",
                     "bold": True, "align": "center"},
        "keywords": {"zh_font": "宋体", "en_font": "Times New Roman", "sz": "22",
                     "bold": True, "align": "left"},
        "ack": {"zh_font": "黑体", "en_font": "Times New Roman", "sz": "22",
                "bold": True, "align": "center"},
    },
    "headingStyles": {
        "1": {"styleId": "Heading1", "name": "标题 1", "font": "黑体",
              "size": "小三", "outline_level": 1},
        "2": {"styleId": "Heading2", "name": "标题 2", "font": "黑体",
              "size": "四号", "outline_level": 2},
        "3": {"styleId": "Heading3", "name": "标题 3", "font": "黑体",
              "size": "小四", "outline_level": 3},
    },
    "body": {"styleId": "Normal", "font": "宋体", "size": "小四",
             "line_spacing_pt": 20, "first_line_indent_chars": 2},
    # 学校要求表格转学术三线表。字段名来自 headings_fix.fix()：
    # profile["tables"]["threeLine"] 为真才转（three_line=None 时跟随画像）。
    "tables": {"threeLine": True},
    "page": {},
    "spec": {},
}

ABSTRACT = (
    "本文研究毕业论文格式自动修正方法。高校对本科学位论文的版式有明确要求，"
    "但学生在写作过程中往往只关注内容，忽视字体、字号、行距、缩进与参考文献体例，"
    "定稿前集中返工的现象十分普遍。本文在梳理常见格式问题的基础上，提出一种由"
    "学校模板驱动的自动修正方案，把格式要求抽象为可执行的画像，逐项比对并施加修改。"
    "实验表明，该方案对标题层级、正文格式、参考文献体例与表格线型的修正均稳定可用。"
)

KEYWORDS = "神经网络；模板驱动；格式修正；毕业论文"

# (级别, 标题文本) —— 全是手打纯文本，故意不套样式
HEADINGS = [
    (1, "1 绪论"),
    (2, "1.1 研究背景"),
    (2, "1.2 研究意义"),
    (1, "2 文献综述"),
    (2, "2.1 已有研究"),
    (1, "3 研究方法"),
    (2, "3.1 数据来源"),
    (1, "4 结论"),
]

BODY = (
    "毕业论文是本科培养方案中的最后一个环节，其版式规范通常由学校教务处统一发布。"
    "由于规范条目细碎，学生在排版时容易遗漏，常见问题包括正文字号不一致、行距不统一、"
    "首行缩进缺失、标题层级混乱以及参考文献著录格式不符。"
)

REFERENCES = [
    "[1] 张三. 毕业论文格式规范研究[J]. 高等教育研究, 2020, 12(3): 45-50.",
    "[2] 李四. 自动化排版系统的设计与实现[D]. 北京: 某大学, 2021.",
    "[3] 王五, 赵六. 模板驱动的文档格式校验方法[J]. 计算机应用, 2022, 42(7): 101-108.",
]

TABLE = [
    ["变量", "含义", "取值范围"],
    ["font_size", "正文字号", "小四 / 五号"],
    ["line_spacing", "行距", "固定值 20 磅"],
    ["indent", "首行缩进", "2 字符"],
]


def _set_zh_font(run, name: str, size_pt: float | None = None, bold=None):
    run.font.name = name
    if size_pt is not None:
        run.font.size = Pt(size_pt)
    if bold is not None:
        run.bold = bold
    rpr = run._element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = rpr.makeelement(qn("w:rFonts"), {})
        rpr.append(rf)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rf.set(qn(attr), name)


def _para(doc, text, size_pt=10.5, font="宋体", align=None, bold=None):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    _set_zh_font(p.add_run(text), font, size_pt, bold)
    return p


def _plain_para(doc, text):
    """只靠 Normal 样式定格式的正文段落 —— 客户最常见的写法。

    为什么素材里必须放几段：绝大多数论文的正文根本不写 run 级字号，字号由
    Normal 样式给（客户在 Word 里把样式设好，正文一段段敲下来就完了）。
    只在「段落有显式字号」时才比对的检查器，对这种写法会**全盘漏检** ——
    实测就是这样：正文全是五号，报告却说正文没问题。素材得覆盖这条真实路径，
    否则 CI 永远发现不了这个回归。
    """
    p = doc.add_paragraph()
    p.add_run(text)
    return p


def build_thesis(path: str) -> str:
    doc = Document()

    # 页面：A4 + 2.54cm（默认就是，写清楚便于日后核对）
    for sec in doc.sections:
        sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
        sec.top_margin = sec.bottom_margin = Cm(2.54)
        sec.left_margin = sec.right_margin = Cm(2.54)

    # 正文样式调成"不合规"：五号、单倍行距、无缩进
    normal = doc.styles["Normal"]
    normal.font.name = "宋体"
    normal.font.size = Pt(10.5)
    rpr = normal.element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = rpr.makeelement(qn("w:rFonts"), {})
        rpr.append(rf)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rf.set(qn(attr), "宋体")
    pf = normal.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
    pf.first_line_indent = Pt(0)
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)

    # 摘要 / 关键词（手打纯文本，不套样式）
    _para(doc, "摘要", size_pt=10.5, align=1)
    _para(doc, ABSTRACT)
    _para(doc, "关键词：" + KEYWORDS)

    # 正文 + 手打标题
    for level, text in HEADINGS:
        _para(doc, text, size_pt=10.5, bold=None)   # 故意不加粗、不套样式
        _para(doc, BODY)
        if level == 1:
            # 一级章节后面多一段「只靠 Normal 定字号」的正文（只给文字，不设格式）
            _plain_para(doc, BODY)

    # 一张带框线的表（要求三线表）
    # 表题写在表格**上方** —— 这是学术规范，也是引擎 _caption_targets 的定位口径
    # （表题在「紧邻 w:tbl 之前」的那一段）。放到表格下面，引擎会当正文段落去核查，
    # 报「首行缩进缺失 / 字号应为 12pt」，那是素材不自洽，不是产品问题。
    _para(doc, "表1-1 格式要素示例", size_pt=10.5, align=1)
    t = doc.add_table(rows=len(TABLE), cols=len(TABLE[0]))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for r, row in enumerate(TABLE):
        for c, val in enumerate(row):
            cell = t.cell(r, c)
            cell.text = ""
            _set_zh_font(cell.paragraphs[0].add_run(val), "宋体", 10.5)

    # 参考文献（格式标准，验的是编号重排与体例）
    _para(doc, "参考文献", size_pt=10.5)
    for ref in REFERENCES:
        _para(doc, ref, size_pt=10.5)

    # 致谢
    _para(doc, "致谢", size_pt=10.5)
    _para(doc, "感谢导师在写作过程中的悉心指导。")

    os.makedirs(os.path.dirname(path), exist_ok=True)
    doc.save(path)
    return path


def _three_line_table(doc):
    """在模板里放一张标准三线表样例。

    为什么要它：引擎判定「该校要求什么表格线型」靠的是模板里有没有三线表样例
    （format_profile.extract → profile["tables"]["threeLine"]）。模板里没有数据表，
    画像就是 None，客户论文里的框线表不会被转成三线表 —— 这是模板驱动的**正常**
    行为（没依据不乱改），所以验收素材必须把样例给上，否则验不到三线化能力。
    """
    rows = [["变量", "含义", "取值范围"],
            ["font_size", "正文字号", "小四 / 五号"],
            ["indent", "首行缩进", "2 字符"]]
    t = doc.add_table(rows=len(rows), cols=len(rows[0]))
    tbl_pr = t._tbl.find(qn("w:tblPr"))

    # 表格级三线：上下单线、四边竖线与内部横线一律 nil
    bd = tbl_pr.makeelement(qn("w:tblBorders"), {})
    for tag, val, sz in (("top", "single", 12), ("bottom", "single", 12),
                         ("left", "nil", 0), ("right", "nil", 0),
                         ("insideH", "nil", 0), ("insideV", "nil", 0)):
        el = bd.makeelement(qn("w:" + tag), {})
        el.set(qn("w:val"), val)
        el.set(qn("w:sz"), str(sz))
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "000000")
        bd.append(el)
    # tblPr 子元素有 schema 顺序，必须插在 shd 之前（直接 append 会被 Word 判坏文档）
    tbl_pr.insert_element_before(bd, "w:shd", "w:tblLayout", "w:tblCellMar",
                                 "w:tblLook", "w:tblCaption", "w:tblDescription",
                                 "w:tblPrChange")

    for c, val in enumerate(rows[0]):
        cell = t.cell(0, c)
        cell.text = ""
        _set_zh_font(cell.paragraphs[0].add_run(val), "黑体", 12.0, True)
        tcpr = cell._tc.get_or_add_tcPr()
        tb = tcpr.makeelement(qn("w:tcBorders"), {})
        b = tb.makeelement(qn("w:bottom"), {})
        b.set(qn("w:val"), "single")
        b.set(qn("w:sz"), "6")
        b.set(qn("w:space"), "0")
        b.set(qn("w:color"), "000000")
        tb.append(b)
        tcpr.insert_element_before(tb, "w:shd", "w:noWrap", "w:tcMar",
                                   "w:textDirection", "w:tcFitText", "w:vAlign",
                                   "w:hideMark")
    for r in range(1, len(rows)):
        for c, val in enumerate(rows[r]):
            cell = t.cell(r, c)
            cell.text = ""
            _set_zh_font(cell.paragraphs[0].add_run(val), "宋体", 12.0)
    return t


def build_template(path: str) -> str:
    """造一份「学校模板」docx：把格式要求写进样式定义里。

    为什么要它：真实客户的流程是「上传学校模板 → 引擎提取画像 → 按画像检查/修正」。
    手写 profile.json 是我替引擎做的决定，字段名一旦写错，检查器就读不到要求、
    报告"已达标"，而修正器却按另一套字段改了 —— 这正是本轮实测踩到的坑
    （报告说 ✅ 无需修正，实际改了 24 处）。所以验收必须让引擎自己从模板提画像。

    模板内容 = 各级标题样式（黑体 小三/四号/小四）+ 正文（宋体 小四、固定 20 磅、
    首行缩进 2 字符）+ 表格三线要求，外加几段示例正文供样式取样。
    """
    from docx.enum.text import WD_LINE_SPACING

    doc = Document()
    for sec in doc.sections:
        sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
        sec.top_margin = sec.bottom_margin = Cm(2.54)
        sec.left_margin = sec.right_margin = Cm(2.54)

    # 各级标题样式：黑体 + 小三(15pt)/四号(14pt)/小四(12pt)
    for level, size_pt in ((1, 15.0), (2, 14.0), (3, 12.0)):
        st = doc.styles["Heading %d" % level]
        st.font.name = "黑体"
        st.font.size = Pt(size_pt)
        st.font.bold = True
        rpr = st.element.get_or_add_rPr()
        rf = rpr.find(qn("w:rFonts"))
        if rf is None:
            rf = rpr.makeelement(qn("w:rFonts"), {})
            rpr.append(rf)
        for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
            rf.set(qn(attr), "黑体")

    # 正文：宋体 小四(12pt) + 固定 20 磅行距 + 首行缩进 2 字符
    normal = doc.styles["Normal"]
    normal.font.name = "宋体"
    normal.font.size = Pt(12.0)
    rpr = normal.element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = rpr.makeelement(qn("w:rFonts"), {})
        rpr.append(rf)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rf.set(qn(attr), "宋体")
    pf = normal.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    pf.line_spacing = Pt(20)
    pf.first_line_indent = Pt(24)      # 2 字符 × 12pt = 24pt
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)

    # 放几段示例，让画像能从"真实使用样"里取到每级要素
    doc.add_heading("第一章 绪论", level=1)
    _para(doc, "这里是学校模板的正文示例段落，用于向引擎展示正文的字体字号与缩进。" * 2,
          size_pt=12.0, font="宋体")
    doc.add_heading("1.1 研究背景", level=2)
    _para(doc, "二级标题下的正文示例。" * 6, size_pt=12.0, font="宋体")
    doc.add_heading("1.1.1 问题提出", level=3)
    _para(doc, "三级标题下的正文示例。" * 6, size_pt=12.0, font="宋体")
    doc.add_heading("参考文献", level=1)
    _para(doc, "[1] 张三. 示例文献[J]. 示例期刊, 2020, 1(1): 1-10.", size_pt=12.0)

    # 一张三线表样例 —— 告诉引擎"本校表格用三线表"
    _para(doc, "表1 示例三线表", size_pt=12.0, align=1)
    _three_line_table(doc)

    os.makedirs(os.path.dirname(path), exist_ok=True)
    doc.save(path)
    return path


def main(argv=None) -> int:
    argv = argv if argv is not None else sys.argv[1:]
    out = argv[0] if argv else OUT_DIR
    os.makedirs(out, exist_ok=True)
    thesis = build_thesis(os.path.join(out, "thesis.docx"))
    template = build_template(os.path.join(out, "template.docx"))
    profile = os.path.join(out, "profile.json")
    with open(profile, "w", encoding="utf-8") as f:
        json.dump(PROFILE, f, ensure_ascii=False, indent=2)
    print("thesis  : %s (%d bytes)" % (thesis, os.path.getsize(thesis)))
    print("template: %s (%d bytes)" % (template, os.path.getsize(template)))
    print("profile : %s (%d bytes, 手写兜底；验收请用引擎从 template 现提)"
          % (profile, os.path.getsize(profile)))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
