# -*- coding: utf-8 -*-
"""学生版 · 批注幂等回归（Codex 审查 P2 修复验证）

核心验证：用同一画像对一篇已修正的论文「再跑一遍」一键修正，
提示类批注（关键词 / 附录 / 疑似题注）不得重复追加气泡。

这正是用户反复踩的坑——"改完一遍后，拿同一个模板再跑一遍又改出一堆问题"
（此处特指批注气泡翻倍）。修复位于 headings_fix._apply_comments 的
幂等去重分支（v1.3.121）。

场景：
  1) 构造含「关键词：...」同行行 + 附录区的论文 docx；
     配含 keywords / appendix 级别的模板画像；
  2) 第一次一键修正（add_comments=True）→ 统计批注数 C1（应 >= 1）；
  3) 对修正产物再跑一次 → 统计批注数 C2；
  4) 断言 C2 == C1（幂等，不再翻倍）。

用法：python tests/test_comment_idempotency.py
退出码 0=通过，1=失败。
"""
import os
import sys
import json
import zipfile
import tempfile
import shutil
import traceback

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, ".."))   # tfd-desktop/
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from tfd_app import engine   # noqa: E402

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


# ---------------------------------------------------------------- 构造测试 docx
def _para(text, style=None):
    ppr = '<w:pPr><w:pStyle w:val="%s"/></w:pPr>' % style if style else ""
    return '<w:p>%s<w:r><w:t xml:space="preserve">%s</w:t></w:r></w:p>' % (ppr, text)


def _table():
    return (
        '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/>'
        '<w:tblW w:w="0" w:type="auto"/></w:tblPr>'
        '<w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>'
        '<w:p><w:r><w:t>变量</w:t></w:r></w:p></w:tc>'
        '<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>'
        '<w:p><w:r><w:t>说明</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
    )


def build_thesis(path):
    # 关键词用「关键词：kw1；kw2」同行结构（触发 keywords_note 批注）。
    # 末尾加「附录」区（触发 appendix_note 批注，且其必须在参考文献之后）。
    body = "".join([
        _para("摘要"),
        _para("本文研究毕业论文格式自动修正方法，验证批注幂等。"),
        _para("关键词：神经网络；模板驱动；格式修正"),
        _para("1 绪论"),
        _para("这是正文内容，用于验证正文区的格式套用与隔离。"),
        _para("2 文献综述"),
        _para("已有研究从多个角度探讨了格式自动化。"),
        _para("参考文献"),
        _para("[1] 张三. 毕业论文格式规范研究[J]. 高等教育, 2020, 12(3): 45-50."),
        _para("附录"),
        _para("附录A 调查问卷样本与数据表说明。"),
    ])
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<w:document xmlns:w="%s"><w:body>%s'
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
           '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>'
           '</w:body></w:document>' % (W, body))
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              '<w:styles xmlns:w="%s">'
              '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
              '<w:name w:val="Normal"/></w:style>'
              '<w:style w:type="paragraph" w:styleId="Heading1">'
              '<w:name w:val="heading 1"/><w:basedOn w:val="Normal"/></w:style>'
              '</w:styles>' % W)
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
          '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
          '</Types>')
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="word/styles.xml"/>'
            '</Relationships>')
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", rels)
        z.writestr("word/document.xml", doc)
        z.writestr("word/styles.xml", styles)


def build_profile_json(path):
    """含 keywords / appendix 级别的合法画像（驱动关键词、附录提示批注）。"""
    prof = {
        "source": "test_template.docx",
        "levels": {
            "abstract": {"zh_font": "宋体", "en_font": "Times New Roman", "sz": "22",
                         "bold": True, "align": "center"},
            "keywords": {"zh_font": "宋体", "en_font": "Times New Roman", "sz": "22",
                         "bold": True, "align": "left"},
            "ack": {"zh_font": "黑体", "en_font": "Times New Roman", "sz": "22",
                    "bold": True, "align": "center"},
            "appendix": {"zh_font": "黑体", "en_font": "Times New Roman", "sz": "22",
                         "bold": True, "align": "center"},
        },
        "headingStyles": {
            "1": {"styleId": "Heading1", "name": "标题 1", "font": "黑体",
                  "size": "小三", "outline_level": 1},
        },
        "body": {"styleId": "Normal", "font": "宋体", "size": "小四"},
        "page": {},
        "spec": {},
    }
    with open(path, "w", encoding="utf-8") as f:
        json.dump(prof, f, ensure_ascii=False, indent=2)


def _is_valid_docx(p):
    if not os.path.isfile(p):
        return False
    try:
        with zipfile.ZipFile(p) as z:
            return "word/document.xml" in z.namelist()
    except Exception:
        return False


def _count_comments(path):
    """统计 word/comments.xml 中的批注气泡数（<w:comment ...> 元素个数）。"""
    with zipfile.ZipFile(path) as z:
        if "word/comments.xml" not in z.namelist():
            return 0
        data = z.read("word/comments.xml").decode("utf-8")
    return data.count("<w:comment ")


def run_all():
    tmp = tempfile.mkdtemp(prefix="tfd_cid_")
    try:
        src = os.path.join(tmp, "thesis.docx")
        profile = os.path.join(tmp, "profile.json")
        build_thesis(src)
        build_profile_json(profile)
        if not _is_valid_docx(src):
            print("  XX  构造测试 docx 非法")
            return False
        print("  OK  构造测试 docx 合法（含关键词同行行 + 附录区）")

        # 第一次修正（带批注）
        dst1 = os.path.join(tmp, "fixed1.docx")
        try:
            engine.run_fix_headings(src, dst1, profile_path=profile, add_comments=True)
            c1 = _count_comments(dst1)
            print("  OK  第一次一键修正产出合法 docx，批注数=%d" % c1)
        except Exception:
            print("  XX  第一次一键修正异常：" + traceback.format_exc().splitlines()[-1])
            return False

        # 第二次修正（对修正产物再跑一遍——用户反复踩坑的场景）
        dst2 = os.path.join(tmp, "fixed2.docx")
        try:
            engine.run_fix_headings(dst1, dst2, profile_path=profile, add_comments=True)
            c2 = _count_comments(dst2)
            print("  OK  第二次一键修正产出合法 docx，批注数=%d" % c2)
        except Exception:
            print("  XX  第二次一键修正异常：" + traceback.format_exc().splitlines()[-1])
            return False

        if c1 < 1:
            print("  XX  第一次修正应至少产生 1 条提示批注（关键词/附录），实际=%d" % c1)
            return False
        if c2 != c1:
            print("  XX  批注幂等失败：二次修正批注翻倍 %d -> %d" % (c1, c2))
            return False
        print("  OK  批注幂等：二次修正未重复追加气泡（%d == %d）" % (c1, c2))
        return True
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    ok = run_all()
    print("\n==== 学生版批注幂等回归: %s ====" % ("通过" if ok else "失败"))
    sys.exit(0 if ok else 1)
