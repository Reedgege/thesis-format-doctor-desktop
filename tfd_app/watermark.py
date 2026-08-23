# -*- coding: utf-8 -*-
"""
试用水印模块 · 论文格式医生
==================================================================
对"修正后的 docx"做后处理（纯标准库 zipfile + ElementTree），注入：
  1. 页眉 + 页脚：红色文字"试用版 · 论文格式医生（正式版无水印）"；
  2. 页眉 VML 背景大水印（Word 显示；WPS 兼容性差，仅作增强层）；
  3. 正文穿插：红色加粗文字行（每 8 段一处）+ 红色水印图片（每 14 段一张，
     图片任何软件必显示，且要逐张删除——最可靠的防白嫖层）。

正式版（有激活码）不调用本模块，输出无水印文档。
"""
import os
import zipfile
import xml.etree.ElementTree as ET

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
PR = "http://schemas.openxmlformats.org/package/2006/relationships"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"

WR = "{%s}" % W   # w: 前缀包裹器（ElementTree tag 需花括号形式）
RR = "{%s}" % R
CTR = "{%s}" % CT
PRR = "{%s}" % PR
WPR = "{%s}" % WP
AR = "{%s}" % A
PICR = "{%s}" % PIC

ET.register_namespace("w", W)
ET.register_namespace("r", R)
ET.register_namespace("ct", CT)
ET.register_namespace("pr", PR)
ET.register_namespace("wp", WP)
ET.register_namespace("a", A)
ET.register_namespace("pic", PIC)

_WM_HEADER = "试用版 · 论文格式医生（正式版无水印）"
_WM_BODY = "【试用版】论文格式医生 · 一键按学校模板修正格式（正式版无水印）"

_HDR_REL_ID = "rIdTfdHdr"
_FTR_REL_ID = "rIdTfdFtr"
_IMG_REL_ID = "rIdTfdWmImg"
_IMG_NAME = "watermark.png"
# 水印图显示尺寸（px→EMU：1px@96dpi=9525EMU；520x150 → 保持宽约 360pt）
_IMG_EMU_W = int(360 * 12700)      # 360pt = 4572000 EMU
_IMG_EMU_H = int(_IMG_EMU_W * 150 / 520)

# 水印图片资源（与 gui 同包；打包后经 assets 数据目录携带）
_IMG_SRC = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                        "assets", _IMG_NAME)


def _wm_run(text, sz="20", color="C00000", bold=True):
    """页眉/页脚里的红字 run（v1.3.60：加大、加粗、红色）。"""
    r = ET.Element(WR + "r")
    rpr = ET.SubElement(r, WR + "rPr")
    c = ET.SubElement(rpr, WR + "color")
    c.set(WR + "val", color)
    if bold:
        ET.SubElement(rpr, WR + "b")
    s = ET.SubElement(rpr, WR + "sz")
    s.set(WR + "val", sz)            # 10pt（页眉页脚用）
    s2 = ET.SubElement(rpr, WR + "szCs")
    s2.set(WR + "val", sz)
    t = ET.SubElement(r, WR + "t")
    t.text = text
    return r


def _wm_body_para():
    """正文穿插的试用水印段落（v1.3.60：红色、加粗、加大、居中）。"""
    p = ET.Element(WR + "p")
    ppr = ET.SubElement(p, WR + "pPr")
    jc = ET.SubElement(ppr, WR + "jc")
    jc.set(WR + "val", "center")
    r = ET.SubElement(p, WR + "r")
    rpr = ET.SubElement(r, WR + "rPr")
    c = ET.SubElement(rpr, WR + "color")
    c.set(WR + "val", "C00000")
    ET.SubElement(rpr, WR + "b")
    s = ET.SubElement(rpr, WR + "sz")
    s.set(WR + "val", "26")          # 13pt
    s2 = ET.SubElement(rpr, WR + "szCs")
    s2.set(WR + "val", "26")
    t = ET.SubElement(r, WR + "t")
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = _WM_BODY
    return p


def _wm_shape(xml_id, top, width, height, rot, opacity=".4"):
    """单个 VML 背景水印 shape（Word 标准写法：font-size:1pt + mso-fit-shape-to-text 自动放大）。

    每个 shape 独立 id，可在一个页眉里放多个 → 每页多个大背景水印。
    """
    return (
        '<v:shape id="PowerPlusWaterMarkObject%s" o:spid="_x0000_s%s" type="#_x0000_t136" '
        'style="position:absolute;margin-left:0;margin-top:%s;width:%spt;height:%spt;'
        'z-index:-251654144;rotation:%s;'
        'mso-position-horizontal:center;mso-position-horizontal-relative:margin;'
        'mso-position-vertical:center;mso-position-vertical-relative:margin" '
        'o:allowincell="f" filled="f" stroked="f">'
        '<v:fill color="#C00000" opacity="%s"/>'
        '<v:textpath style="font-family:&quot;微软雅黑&quot;;font-size:1pt;'
        'mso-fit-shape-to-text:t;color:#C00000" string="%%s" id="PowerPlusWaterMarkObjectPath%s"/>'
        '</v:shape>' % (xml_id, 1025 + xml_id, top, width, height, rot, opacity, xml_id))


def _header_xml(text):
    """页眉 XML：每页 3 个红色斜向/水平大背景水印（删起来麻烦）+ 页眉红字。

    v1.3.61：修正 VML 标准写法（font-size:1pt + mso-fit-shape-to-text 自动放大），
    并放 3 个 shape（上/中/下三个位置，斜向 315° ×2 + 水平 ×1），
    使每页正文背景出现多个大红"试用版"水印。
    """
    esc = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")
    shapes = "".join([
        _wm_shape(1, 0, 415, 160, 315),     # 上部：斜向
        _wm_shape(2, 180, 415, 160, 315),   # 中部：斜向
        _wm_shape(3, 340, 415, 160, 0),     # 下部：水平
    ]) % (esc, esc, esc)                    # 3 个 shape 各有一个 string 占位
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<w:hdr xmlns:w="%s" xmlns:v="urn:schemas-microsoft-com:vml" '
            'xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="%s">'
            '<w:p><w:pPr><w:pStyle w:val="Header"/></w:pPr><w:r><w:rPr><w:noProof/></w:rPr>'
            '<w:pict>%s</w:pict></w:r></w:p>'
            '<w:p><w:pPr><w:pStyle w:val="Header"/></w:pPr><w:r><w:rPr>'
            '<w:color w:val="C00000"/><w:sz w:val="20"/><w:b/></w:rPr>'
            '<w:t>%s</w:t></w:r></w:p></w:hdr>' % (W, R, shapes, esc))


def _xml_str(root):
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + \
           ET.tostring(root, encoding="unicode")


def _insert_sect_refs(doc_root):
    """给 document.xml 里每个 sectPr 插入 headerReference / footerReference。"""
    added = 0
    for sect in doc_root.iter(WR + "sectPr"):
        # 元素顺序要求：headerReference/footerReference 需在 pgSz 之前
        hr = ET.Element(WR + "headerReference")
        hr.set(WR + "type", "default")
        hr.set(RR + "id", _HDR_REL_ID)
        sect.insert(0, hr)
        fr = ET.Element(WR + "footerReference")
        fr.set(WR + "type", "default")
        fr.set(RR + "id", _FTR_REL_ID)
        sect.insert(1, fr)
        added += 1
    return added


def _insert_body_paras(doc_root):
    """正文穿插：按密度每 8 段插一处红色水印行（开头必插），最多 30 处。

    v1.3.61：水印"多而难删"——正文几十处红色水印行，删起来很麻烦。
    """
    body = doc_root.find(WR + "body")
    if body is None:
        return 0
    paras = list(body.findall(WR + "p"))
    if not paras:
        return 0
    step = 8
    positions = list(range(0, len(paras), step))[:30]
    all_p = list(body)
    for pos in reversed(positions):
        idx = pos
        if idx >= len(all_p):
            idx = len(all_p)
            for j, el in enumerate(all_p):
                if el.tag == WR + "sectPr":
                    idx = j
                    break
        body.insert(idx, _wm_body_para())
        all_p = list(body)
    return len(positions)


def _image_para():
    """含水印图片的段落（居中，带红色边框的水印图）。"""
    p = ET.Element(WR + "p")
    ppr = ET.SubElement(p, WR + "pPr")
    jc = ET.SubElement(ppr, WR + "jc")
    jc.set(WR + "val", "center")
    r = ET.SubElement(p, WR + "r")
    dr = ET.SubElement(r, WR + "drawing")
    inline = ET.SubElement(dr, WPR + "inline")
    for attr in ("distT", "distB", "distL", "distR"):
        inline.set(attr, "0")
    extent = ET.SubElement(inline, WPR + "extent")
    extent.set("cx", str(_IMG_EMU_W))
    extent.set("cy", str(_IMG_EMU_H))
    docpr = ET.SubElement(inline, WPR + "docPr")
    docpr.set("id", "1")
    docpr.set("name", "TFDWatermark")
    graphic = ET.SubElement(inline, AR + "graphic")
    gdata = ET.SubElement(graphic, AR + "graphicData")
    gdata.set("uri", PIC)
    pic = ET.SubElement(gdata, PICR + "pic")
    nv = ET.SubElement(pic, PICR + "nvPicPr")
    cnvpr = ET.SubElement(nv, PICR + "cNvPr")
    cnvpr.set("id", "1")
    cnvpr.set("name", "tfd_wm")
    ET.SubElement(nv, PICR + "cNvPicPr")
    blipfill = ET.SubElement(pic, PICR + "blipFill")
    blip = ET.SubElement(blipfill, AR + "blip")
    blip.set(RR + "embed", _IMG_REL_ID)
    stretch = ET.SubElement(blipfill, AR + "stretch")
    ET.SubElement(stretch, AR + "fillRect")
    sppr = ET.SubElement(pic, PICR + "spPr")
    xfrm = ET.SubElement(sppr, AR + "xfrm")
    off = ET.SubElement(xfrm, AR + "off")
    off.set("x", "0")
    off.set("y", "0")
    ext = ET.SubElement(xfrm, AR + "ext")
    ext.set("cx", str(_IMG_EMU_W))
    ext.set("cy", str(_IMG_EMU_H))
    geom = ET.SubElement(sppr, AR + "prstGeom")
    geom.set("prst", "rect")
    ET.SubElement(geom, AR + "avLst")
    return p


def _insert_image_paras(doc_root):
    """正文按密度（每 14 段）插入水印图片段落（WPS/Word 必显示，逐张难删）。"""
    body = doc_root.find(WR + "body")
    if body is None:
        return 0
    paras = list(body.findall(WR + "p"))
    if not paras:
        return 0
    positions = list(range(0, len(paras), 14))[:15]   # 最多 15 张
    all_p = list(body)
    for pos in reversed(positions):
        idx = pos
        if idx >= len(all_p):
            idx = len(all_p)
            for j, el in enumerate(all_p):
                if el.tag == WR + "sectPr":
                    idx = j
                    break
        body.insert(idx, _image_para())
        all_p = list(body)
    return len(positions)


def apply_watermark(path):
    """给修正后的 docx 就地加试用水印（重写 zip）。返回 True 成功；异常返回 False。"""
    tmp = path + ".wm.tmp"
    try:
        with zipfile.ZipFile(path, "r") as zin:
            names = zin.namelist()
            parts = {n: zin.read(n) for n in names}

        doc_xml = parts.get("word/document.xml")
        if doc_xml is None:
            return False
        doc_root = ET.fromstring(doc_xml)
        _insert_body_paras(doc_root)
        _insert_image_paras(doc_root)   # v1.3.62：正文插入水印图片（WPS 必显示）
        _insert_sect_refs(doc_root)
        parts["word/document.xml"] = _xml_str(doc_root).encode("utf-8")

        # 页眉（VML 红色斜向大水印 + 红字）/ 页脚（红字）
        parts["word/header1.xml"] = _header_xml(_WM_HEADER).encode("utf-8")
        parts["word/footer1.xml"] = _xml_str(
            _header_footer("ftr", _WM_HEADER)).encode("utf-8")

        # Content_Types 注册（header/footer/png）
        ct_xml = parts.get("[Content_Types].xml", b"")
        if b"header+xml" not in ct_xml:
            ct_root = ET.fromstring(ct_xml)
            for part, ctype in (("word/header1.xml",
                                 "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"),
                                ("word/footer1.xml",
                                 "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml")):
                ov = ET.SubElement(ct_root, CTR + "Override")
                ov.set("PartName", "/" + part)
                ov.set("ContentType", ctype)
            parts["[Content_Types].xml"] = _xml_str(ct_root).encode("utf-8")
        if b'Extension="png"' not in ct_xml:
            ct_root = ET.fromstring(parts.get("[Content_Types].xml", b""))
            d = ET.SubElement(ct_root, CTR + "Default")
            d.set("Extension", "png")
            d.set("ContentType", "image/png")
            parts["[Content_Types].xml"] = _xml_str(ct_root).encode("utf-8")

        # document.xml.rels 注册（header/footer/图片）
        rels_path = "word/_rels/document.xml.rels"
        rels_xml = parts.get(rels_path, b"")
        if _HDR_REL_ID.encode("utf-8") not in rels_xml:
            rels_root = ET.fromstring(rels_xml) if rels_xml else ET.Element(PRR + "Relationships")
            for rid, target, typ in ((_HDR_REL_ID, "header1.xml",
                                      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"),
                                     (_FTR_REL_ID, "footer1.xml",
                                      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"),
                                     (_IMG_REL_ID, "media/" + _IMG_NAME,
                                      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")):
                rel = ET.SubElement(rels_root, PRR + "Relationship")
                rel.set("Id", rid)
                rel.set("Type", typ)
                rel.set("Target", target)
            parts[rels_path] = _xml_str(rels_root).encode("utf-8")

        # 水印图片媒体文件
        if os.path.isfile(_IMG_SRC):
            with open(_IMG_SRC, "rb") as f:
                parts["word/media/" + _IMG_NAME] = f.read()

        # 重写 zip（新增条目：header1/footer1/media 水印图）
        extra = ["word/header1.xml", "word/footer1.xml",
                 "word/media/" + _IMG_NAME]
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for n in names:
                zout.writestr(n, parts[n])
            for n in extra:
                if n in parts:
                    zout.writestr(n, parts[n])
        os.replace(tmp, path)
        return True
    except Exception:
        try:
            if os.path.exists(tmp):
                os.remove(tmp)
        except Exception:
            pass
        return False


def _header_footer(tag, text):
    root = ET.Element(WR + tag)
    p = ET.SubElement(root, WR + "p")
    ppr = ET.SubElement(p, WR + "pPr")
    jc = ET.SubElement(ppr, WR + "jc")
    jc.set(WR + "val", "center")
    p.append(_wm_run(text))
    return root
