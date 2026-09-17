# -*- coding: utf-8 -*-
"""回归：结构页套用 pass 不得误伤目录条目（如「摘要I」「ABSTRACTII」）。"""
import io, sys, zipfile, xml.etree.ElementTree as ET, os, json
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "tfd_app", "core"))

from docxutils import WR
import headings_fix as hf

NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = "{%s}" % NS

def _build_docx(styles_xml, document_xml, comments_xml=""):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>')
        z.writestr("word/document.xml", document_xml)
        z.writestr("word/styles.xml", styles_xml)
        z.writestr("word/_rels/document.xml.rels",
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>')
        if comments_xml:
            z.writestr("word/comments.xml", comments_xml)
    buf.seek(0)
    return zipfile.ZipFile(buf)


def test_structural_pass_skips_toc_entries():
    """目录条目即使文本像「摘要I/ABSTRACTII」，也不应被结构页 pass 批注/改格式。"""
    styles_xml = (
        '<w:styles xmlns:w="%s">'
        '<w:style w:type="paragraph" w:styleId="22"><w:name w:val="toc 1"/></w:style>'
        '</w:styles>' % NS
    )
    # 构造一个极简文档：
    #   p0: 摘要标题（真正结构页标题）
    #   p1: 摘要I（目录条目，样式 toc 1，含 TOC 域）
    #   p2: ABSTRACTII（目录条目，样式 toc 1）
    document_xml = (
        '<w:document xmlns:w="%s"><w:body>'
        '<w:p><w:pPr><w:pStyle w:val="3"/></w:pPr><w:r><w:t>摘要</w:t></w:r></w:p>'
        '<w:p><w:pPr><w:pStyle w:val="22"/></w:pPr>'
        '<w:r><w:instrText xml:space="preserve"> HYPERLINK \\l "_Toc1" </w:instrText></w:r>'
        '<w:r><w:t>摘要I</w:t></w:r>'
        '<w:r><w:instrText xml:space="preserve"> PAGEREF _Toc1 \\h </w:instrText></w:r>'
        '</w:p>'
        '<w:p><w:pPr><w:pStyle w:val="22"/></w:pPr>'
        '<w:r><w:instrText xml:space="preserve"> HYPERLINK \\l "_Toc2" </w:instrText></w:r>'
        '<w:r><w:t>ABSTRACTII</w:t></w:r>'
        '<w:r><w:instrText xml:space="preserve"> PAGEREF _Toc2 \\h </w:instrText></w:r>'
        '</w:p>'
        '</w:body></w:document>' % NS
    )
    profile = {
        "source": "fake",
        "headingStyles": {"1": {"styleId": "3", "name": "heading 1"}},
        "levels": {
            "abstract": {"zh_font": "黑体", "sz": 30, "bold": True, "align": "center"},
            "abstract_title": {"zh_font": "黑体", "sz": 30, "bold": True, "align": "center"},
        },
    }
    z = _build_docx(styles_xml, document_xml)
    # 临时写出 docx
    tmp_src = os.path.join(os.path.dirname(__file__), "_tmp_structural_toc.docx")
    tmp_dst = os.path.join(os.path.dirname(__file__), "_tmp_structural_toc_fixed.docx")
    with open(tmp_src, "wb") as f:
        f.write(z.read("[Content_Types].xml"))
    # 重新打包成合法 docx
    with zipfile.ZipFile(tmp_src, "w", zipfile.ZIP_DEFLATED) as zout:
        for name in z.namelist():
            zout.writestr(name, z.read(name))
    changes = hf.fix(tmp_src, tmp_dst, profile=profile, add_comments=True)
    if isinstance(changes, tuple):
        changes = changes[0]

    toc_texts = set()
    with zipfile.ZipFile(tmp_dst, 'r') as z2:
        root = ET.fromstring(z2.read("word/document.xml"))
        for p in root.iter(W + "p"):
            text = "".join(t.text or "" for t in p.iter(W + "t"))
            refs = list(p.iter(W + "commentReference"))
            if text in ("摘要I", "ABSTRACTII"):
                assert not refs, f"目录条目 {text!r} 不应被加批注"
                toc_texts.add(text)
    assert "摘要I" in toc_texts and "ABSTRACTII" in toc_texts

    # 清理
    for f in (tmp_src, tmp_dst):
        if os.path.exists(f):
            os.remove(f)
    print("✓ test_structural_pass_skips_toc_entries")


if __name__ == "__main__":
    test_structural_pass_skips_toc_entries()
    print("全部通过")
