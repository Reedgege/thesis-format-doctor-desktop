# -*- coding: utf-8 -*-
"""回归：模板画像必须认对正文样式，题注不能按正文核查。

两个都是「报告看起来对、客户拿到的文档却不对」的坑：

1. **正文样式被内置引用块样式挤掉**。Word 内置样式表里 `Body Text`、`Body Text 2`、
   `Body Text 3`、`Body Text 3 Char` 这些名字都以 `Body Text` 开头，在 styles.xml
   里的顺序又排在 `Normal` 之后。`is_body_style_name` 用前缀匹配，遍历下来最后一个
   命中的就会覆盖前面 —— 实测结果是画像正文取到 `Body Text 3 Char`（sz=16，即 8pt），
   而模板里的 `Normal` 明明是宋体 12pt / 固定 20 磅 / 首行缩进 2 字符。
   画像错了，检查器就按 8pt 去判客户的论文，一键修正也按 8pt 改。

2. **表题被当成正文核查**。`表1-1 格式要素示例` 是题注（学术规范里题注常用五号、
   居中、无首行缩进），拿正文标准判它，客户会看到「未设置首行缩进，模板要求首行
   缩进 2 字符」这种既改不掉也不该改的报错 —— 修正器只在模板给出题注规范时才动题注，
   于是这条错误永远消不掉，一键修正做完还是报 3 个问题。
"""
import json
import os
import sys

import pytest

_PY = sys.executable
_CORE = os.path.join(os.path.dirname(__file__), "..", "tfd_app", "core")
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
sys.path.insert(0, _CORE)

docx = pytest.importorskip("docx")
from docx.enum.text import WD_LINE_SPACING          # noqa: E402
from docx.oxml.ns import qn                          # noqa: E402
from docx.shared import Pt                           # noqa: E402

import format_profile                                # noqa: E402
import format_checker                                # noqa: E402


def _set_zh(rpr_owner, font):
    rpr = rpr_owner.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = rpr.makeelement(qn("w:rFonts"), {})
        rpr.append(rf)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rf.set(qn(attr), font)


@pytest.fixture
def template_with_bodytext3(tmp_path):
    """模板：Normal 是正文（宋体 12pt / 固定 20 磅 / 首行缩进 2 字符），
    外加 Python-docx 自带模板里的 Body Text 系列样式（含 sz=16 的 Body Text 3）。"""
    d = docx.Document()

    normal = d.styles["Normal"]
    normal.font.size = Pt(12.0)
    _set_zh(normal.element, "宋体")
    pf = normal.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    pf.line_spacing = Pt(20)
    pf.first_line_indent = Pt(24)

    # 真正的正文段落：都标 Normal
    d.add_paragraph("本文研究毕业论文格式自动修正方法，正文段落一律使用 Normal 样式。" * 2)
    d.add_paragraph("第二段正文，同样套 Normal，用于让引擎统计样式的实际使用次数。")

    # 一个引用块段落套 Body Text 3（模拟真实论文里存在的引用样式）。
    # 它 sz=16 即 8pt，若被当成正文，画像里的正文字号就会变成 8pt。
    try:
        p = d.add_paragraph("这是一段引用块文字。")
        p.style = d.styles["Body Text 3"]
    except KeyError:      # 某些 python-docx 版本自带模板里没有这条，跳过不影响断言
        pass

    path = tmp_path / "template.docx"
    d.save(str(path))
    return str(path)


def test_body_style_is_normal_not_bodytext3char(template_with_bodytext3):
    prof = format_profile.extract(template_with_bodytext3)

    body = prof.get("body") or {}
    assert body.get("styleId") == "Normal", (
        "正文样式认成了 %r —— 内置 Body Text 系列把 Normal 挤掉了"
        % (body.get("styleId"),))

    lv = (prof.get("levels") or {}).get("body") or {}
    assert lv.get("zh_font") == "宋体", lv
    assert str(lv.get("sz")) == "24", "正文12pt 应记作半磅 24，实际 %r" % (lv.get("sz"),)
    assert lv.get("indent_chars") == 2, lv
    assert (lv.get("line_rule"), lv.get("line_val")) == ("exact", 20), lv


def test_body_style_prefers_the_style_document_actually_uses(tmp_path):
    """模板把正文写成「正文」样式（中文名），且 Normal 空着 —— 应认「正文」。

    学校自制的中文模板常这么干：把 Normal 留空，另建一个叫「正文」的样式。
    此时使用次数是唯一的判断依据，不能退回 Normal。
    """
    d = docx.Document()
    st = d.styles.add_style("正文", 1)      # 1 = WD_STYLE_TYPE.PARAGRAPH
    st.font.size = Pt(14.0)
    _set_zh(st.element, "仿宋")
    st.paragraph_format.first_line_indent = Pt(28)
    d.add_paragraph("正文段落一，使用自定义的「正文」样式。" * 2).style = st
    d.add_paragraph("正文段落二，同样使用「正文」样式。").style = st

    path = tmp_path / "t.docx"
    d.save(str(path))
    prof = format_profile.extract(str(path))

    lv = (prof.get("levels") or {}).get("body") or {}
    assert lv.get("zh_font") == "仿宋", (
        "正文样式应认到「正文」（仿宋），实际拿到 %r" % (lv.get("zh_font"),))
    assert str(lv.get("sz")) == "28", lv


@pytest.mark.parametrize("text", [
    "表1-1 格式要素示例",
    "表 3－1 变量说明",
    "续表3-1 变量说明（续）",
    "Table 1-1 Format elements",
    "Tab.1-1 Format elements",
    "图1-1 系统架构",
    "Figure 2-1 Architecture",
    "资料来源：作者整理",
    "注：数据截至 2025 年。",
])
def test_caption_paragraphs_are_recognized(text):
    assert format_checker._is_caption_para(text) is True, text


@pytest.mark.parametrize("text", [
    "表1-1 显示了三个变量之间的关系，其中 font_size 表示正文字号，indent 表示首行缩进，"
    "这两个变量共同决定了排版是否合规，因此需要在修正时一并处理。",
    "本文研究毕业论文格式自动修正方法。",
    "",
    None,
])
def test_normal_prose_is_not_a_caption(text):
    assert format_checker._is_caption_para(text) is False, text


@pytest.mark.parametrize("sz,expect", [
    (24, "12pt"), (30, "15pt"), (28, "14pt"), (21, "10.5pt"),
    ("24", "12pt"), (None, "未设置"), ("", "未设置"),
])
def test_sz_display_never_renders_none(sz, expect):
    """报告里曾出现「字号应为 None（sz=30，模板）」—— 半磅键叫 sz，没有 size。"""
    got = format_checker._sz_disp(sz)
    assert got == expect, (sz, got)
    assert "None" not in got


# ---------------------------------------------------------------------------
# 正文必须真的被核查
# ---------------------------------------------------------------------------
_BODY_PROFILE = {
    "source": "unit-test",
    "headingStyles": {"1": {"styleId": "Heading1", "name": "heading 1"}},
    "levels": {
        "1": {"zh_font": "黑体", "sz": 30},
        "body": {"zh_font": "宋体", "sz": 24, "indent_chars": 2,
                 "line_rule": "exact", "line_val": 20},
    },
}


@pytest.fixture
def thesis_with_bad_body(tmp_path):
    """正文全部不合规：五号(10.5pt)、单倍行距、没有首行缩进；标题也没套样式。"""
    d = docx.Document()
    normal = d.styles["Normal"]
    normal.font.size = Pt(10.5)
    _set_zh(normal.element, "宋体")

    d.add_paragraph("摘要：本文研究毕业论文格式自动修正方法。")
    d.add_paragraph("1 绪论")
    for _ in range(3):
        d.add_paragraph(
            "毕业论文是本科培养方案中的最后一个环节，其版式规范通常由学校教务处统一发布，"
            "由于规范条目细碎，学生在排版时容易遗漏，常见问题包括正文字号不一致、"
            "行距不统一、首行缩进缺失。")
    path = tmp_path / "thesis.docx"
    d.save(str(path))
    return str(path)


def test_body_paragraphs_are_actually_checked(thesis_with_bad_body):
    """正文段落必须进核查 —— 曾因一道 `continue` 把全部正文整段跳过。

    症状（客户视角）：论文正文全都字号不对、行距不对、没有首行缩进，报告却只列
    标题的问题，正文一个字都没提。客户买这个软件就是为了知道正文对不对，
    漏掉正文等于核心功能失效。
    """
    issues, _struct, _st = format_checker._diagnose(
        thesis_with_bad_body, json.loads(json.dumps(_BODY_PROFILE)))
    msgs = [m for _sev, m in issues]

    body_msgs = [m for m in msgs if "正文" in m]
    assert body_msgs, (
        "正文问题一条都没报出来，说明正文核查分支又被跳过了。全部问题：%r" % (msgs,))

    joined = "\n".join(body_msgs)
    assert "12pt" in joined, "正文实际 10.5pt、模板要求 12pt，应报字号不符：%s" % joined
    assert "缩进" in joined, "正文没有首行缩进，应报缩进缺失：%s" % joined


def test_similar_issues_are_merged_not_repeated(thesis_with_bad_body):
    """同类问题要合并成一条并计数，不能让 42 段正文刷出一屏一模一样的话。"""
    issues, _struct, _st = format_checker._diagnose(
        thesis_with_bad_body, json.loads(json.dumps(_BODY_PROFILE)))
    body_msgs = [m for _sev, m in issues if "正文" in m]
    assert len(body_msgs) <= 5, "正文问题应合并，实际 %d 条：%r" % (len(body_msgs), body_msgs)
    assert any("共" in m for m in body_msgs), "多段同类问题应带「共 N 处」计数：%r" % (body_msgs,)
