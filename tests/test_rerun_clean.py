# -*- coding: utf-8 -*-
"""回归：fix 重跑时清掉本工具上次批注，且只清本工具的、不破坏用户批注、无悬空引用。

同时验证：目录(TOC 域)段落不被批注（bug2）。
"""
import io, sys, zipfile, xml.etree.ElementTree as ET, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "tfd_app", "core"))

from docxutils import WR
import headings_fix as hf

NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = "{%s}" % NS

def _build_docx(comments_xml, document_xml):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>')
        z.writestr("word/document.xml", document_xml)
        z.writestr("word/comments.xml", comments_xml)
        z.writestr("word/_rels/document.xml.rels",
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>')
    buf.seek(0)
    return zipfile.ZipFile(buf)

def test_strip_our_comments_only():
    comments_xml = (
        '<w:comments xmlns:w="%s">'
        '<w:comment w:id="1" w:author="论文格式医生"><w:p><w:r><w:t>本工具批注</w:t></w:r></w:p></w:comment>'
        '<w:comment w:id="2" w:author="张老师"><w:p><w:r><w:t>用户批注</w:t></w:r></w:p></w:comment>'
        '</w:comments>' % NS
    )
    document_xml = (
        '<w:document xmlns:w="%s"><w:body>'
        '<w:p><w:commentRangeStart w:id="1"/><w:r><w:t>段落A</w:t></w:r>'
        '<w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r></w:p>'
        '<w:p><w:commentRangeStart w:id="2"/><w:r><w:t>段落B</w:t></w:r>'
        '<w:commentRangeEnd w:id="2"/><w:r><w:commentReference w:id="2"/></w:r></w:p>'
        '</w:body></w:document>' % NS
    )
    z = _build_docx(comments_xml, document_xml)
    root = ET.fromstring(z.read("word/document.xml"))
    replacements = {}
    n = hf._strip_our_comments(z, root, replacements, "论文格式医生")
    assert n == 1, "应只清掉本工具的1条，实际 %d" % n
    # 用户批注(id=2)保留
    croot = ET.fromstring(replacements["word/comments.xml"])
    ids = [c.get(W + "id") for c in croot.findall(W + "comment")]
    assert "2" in ids, "用户批注不应被清"
    assert "1" not in ids, "本工具批注应被清"
    # document.xml（内存 root）中本工具标记已移除，用户标记保留
    refs = [r.get(W + "id") for r in root.iter(W + "commentReference")]
    assert "1" not in refs, "本工具 commentReference 应被移除"
    assert "2" in refs, "用户 commentReference 应保留"
    starts = [e.get(W + "id") for e in root.iter(W + "commentRangeStart")]
    assert "1" not in starts, "本工具 commentRangeStart 应被移除"
    print("✓ test_strip_our_comments_only")

def test_para_has_toc_field_detects_real_toc():
    # 真实目录条目（HYPERLINK _Toc）应被判为 TOC 段
    doc = ET.fromstring(
        ('<w:document xmlns:w="%s"><w:body><w:p>'
         '<w:r><w:instrText xml:space="preserve"> HYPERLINK \\l "_Toc149316349" </w:instrText></w:r>'
         '<w:r><w:t>1 绪论</w:t></w:r>'
         '<w:r><w:instrText xml:space="preserve"> PAGEREF _Toc149316349 \\h </w:instrText></w:r>'
         '</w:p></w:body></w:document>' % NS).encode("utf-8")
    )
    p = doc.find(W + "body").find(W + "p")
    assert hf._para_has_toc_field(p) is True
    print("✓ test_para_has_toc_field_detects_real_toc")

if __name__ == "__main__":
    test_strip_our_comments_only()
    test_para_has_toc_field_detects_real_toc()
    print("\n全部通过")
