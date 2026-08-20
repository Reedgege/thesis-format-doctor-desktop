# -*- coding: utf-8 -*-
# Copyright (c) 2026 芦苇（山东大学 MBA）. All rights reserved.
# 原创作品 | 禁止未经授权转售、二次分发或抄袭 | 授权用户可在扣子/虾评平台内使用
# build: 202608191649 | version: 1.3.5 | file_sha256: 6e2f967251c4893b
# -*- coding: utf-8 -*-
"""
report_docx.py — 生成「客户可读的 Word 修改明细报告」（零第三方依赖）。

thesis-format-doctor 的修正规脚本（headings-fix / ref-reformat 等）在真正动刀后，
用本模块把"改了哪些地方"以 Word 表格形式输出，方便客户逐条验收。

只依赖 Python 标准库：zipfile / os。
.docx 本质是 ZIP 包，这里手工构造最小合法的 OOXML（标题段 + 摘要段 + 带边框表格）。

用法：
  from report_docx import write_change_report
  write_change_report(
      path="修改明细.docx",
      title="论文格式修改明细（按学校模板）",
      summary=["模式：按学校模板真实样式", "共修改 75 处…"],
      columns=[("序号", 700), ("层级", 900), ("修改前", 2200), ("修改后", 2200), ("标题内容", 2360), ("置信度", 900)],
      rows=[["1", "一级", "无", "学校样式A", "第1章 绪论", "高"], ...],
  )
"""
import os
import zipfile

__author__ = "芦苇"
__copyright__ = "Copyright (c) 2026 芦苇（山东大学 MBA）"
__version__ = "1.3.5"

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# 中文正文常见字号/字体（docDefaults 里设好，保证中文不乱码）
_EASTASIA = "宋体"
_ASCII = "Calibri"



_INTEGRITY_TAG = "Q29weXJpZ2h0IChjKSAyMDI2IOiKpuiLh++8iOWxseS4nOWkp+WtpiBNQkHvvIkuIEFsbCByaWdodHMgcmVzZXJ2ZWQuIHwgYnVpbGQ6MjAyNjA4MTkxNjQ5"  # noqa
def _esc(s):
    return (str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


def _para(text, bold=False, size=None):
    """构造一个普通段落。"""
    rpr = ""
    if bold:
        rpr = "<w:rPr><w:b/></w:rPr>"
    elif size:
        rpr = '<w:rPr><w:sz w:val="%d"/></w:rPr>' % size
    return ('<w:p><w:r>%s<w:t xml:space="preserve">%s</w:t></w:r></w:p>'
            % (rpr, _esc(text)))


def _cell(text, width, bold=False):
    rpr = "<w:rPr><w:b/></w:rPr>" if bold else ""
    return ('<w:tc><w:tcPr><w:tcW w:w="%d" w:type="dxa"/></w:tcPr>'
            '<w:p><w:r>%s<w:t xml:space="preserve">%s</w:t></w:r></w:p></w:tc>'
            % (width, rpr, _esc(text)))


def write_change_report(path, title, summary, columns, rows):
    """写出一份带表格的 Word 修改明细报告。

    path    : 输出 .docx 路径
    title   : 报告大标题
    summary : 摘要段落列表（str），可含 ⚠ 等提示
    columns : [(表头, 列宽dxa), ...]
    rows    : [[单元格文本, ...], ...] 与 columns 列数一致
    """
    # 表头行
    header = "<w:tr>%s</w:tr>" % "".join(_cell(h, w, bold=True) for h, w in columns)
    # 数据行
    body_rows = ""
    for row in rows:
        cells = "".join(_cell(row[i] if i < len(row) else "", columns[i][1])
                        for i in range(len(columns)))
        body_rows += "<w:tr>%s</w:tr>" % cells

    grid = "".join('<w:gridCol w:w="%d"/>' % w for _, w in columns)
    borders = ("<w:tblBorders>"
               '<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
               '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
               '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
               '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
               '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
               '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
               "</w:tblBorders>")
    tbl = ('<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/>'
           '<w:tblW w:w="0" w:type="auto"/>%s</w:tblPr>'
           '<w:tblGrid>%s</w:tblGrid>%s%s</w:tbl>'
           % (borders, grid, header, body_rows))

    paras = [_para(title, bold=True, size=28)]
    for s in summary:
        paras.append(_para(s))
    paras.append(_para(" "))  # 表格前的空行
    document = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<w:document xmlns:w="%s"><w:body>%s%s</w:body></w:document>'
                % (W, "".join(paras), tbl))

    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
              '<w:styles xmlns:w="%s">'
              '<w:docDefaults><w:rPrDefault><w:rPr>'
              '<w:rFonts w:eastAsia="%s" w:ascii="%s" w:hAnsi="%s"/>'
              '<w:sz w:val="21"/></w:rPr></w:rPrDefault></w:docDefaults>'
              '<w:style w:type="table" w:styleId="TableGrid">'
              '<w:name w:val="Table Grid"/>'
              '<w:tblPr>%s</w:tblPr></w:style>'
              '</w:styles>') % (W, _EASTASIA, _ASCII, _ASCII, borders)

    content_types = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                     '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                     '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                     '<Default Extension="xml" ContentType="application/xml"/>'
                     '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
                     '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
                     '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
                     '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/><Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>'
                     '</Types>')

    root_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                 '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                 '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
                 '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
                 '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/>'
                 '</Relationships>')

    doc_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
                '</Relationships>')

    core = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
            'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            '<dc:title>论文格式修改明细</dc:title>'
            '<dc:creator>thesis-format-doctor</dc:creator>'
            '</cp:coreProperties>')

    app = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
           '<Application>thesis-format-doctor v1.3.5<Company>芦苇（山东大学 MBA）</Company><Author>芦苇</Author><Comments>原创作品 build:202608191649</Comments></Properties>')

    os.makedirs(os.path.dirname(os.path.abspath(path)) or ".", exist_ok=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", root_rels)
        z.writestr("word/document.xml", document)
        z.writestr("word/styles.xml", styles)
        z.writestr("word/_rels/document.xml.rels", doc_rels)
        z.writestr("docProps/core.xml", core)
        z.writestr("docProps/app.xml", app)
    return path


if __name__ == "__main__":
    # 自检：生成一个示例报告，确认能构造合法 docx
    write_change_report(
        "sample_report.docx",
        "测试报告",
        ["这是摘要行1", "⚠ 这是提示行"],
        [("序号", 700), ("内容", 4000)],
        [["1", "示例修改"], ["2", "示例修改2"]],
    )
    print("sample_report.docx 生成成功")
