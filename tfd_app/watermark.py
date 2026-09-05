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
import base64
import hashlib
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

_WM_HEADER = "试用版 · 只读预览 ｜ 论文格式医生 · 正式版可编辑无水印"
_WM_BODY = "【试用版 · 只读预览】论文格式医生 · 正式版可编辑无水印"

# v1.3.64：只读保护固定密码（用户设定，不对外公布；客户点编辑需密码才能解锁）。
# 注：v1.3.96 起弃用 base64+XOR 混淆存储（杀软 ML 误报元凶，见 license.py 注释）。
# 该密码在 exe 内本就可被提取，明文常量不降低实际防护等级。
_WM_PASSWORD = "reedskilllunwengeshiyisheng"

_HDR_REL_ID = "rIdTfdHdr"
_FTR_REL_ID = "rIdTfdFtr"
_IMG_REL_ID = "rIdTfdWmImg"
_IMG_NAME = "watermark.png"
_BG_IMG_REL_ID = "rIdTfdBgImg"
_BG_IMG_NAME = "watermark_bg.png"

# 正文穿插水印图显示尺寸（260x75 → 宽约 200pt）
_IMG_EMU_W = int(200 * 12700)      # 200pt = 2540000 EMU
_IMG_EMU_H = int(_IMG_EMU_W * 75 / 260)

# 页眉背景水印图显示尺寸（800x800 透明底 → 显示约 300pt 大图，页面背景居中）
_BG_EMU = int(300 * 12700)         # 300pt 方形

# 水印图片资源（与 gui 同包；打包后经 assets 数据目录携带）
_HERE = os.path.dirname(os.path.abspath(__file__))
_IMG_SRC = os.path.join(_HERE, "assets", _IMG_NAME)
_BG_IMG_SRC = os.path.join(_HERE, "assets", _BG_IMG_NAME)


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


def _bg_anchor_xml():
    """浮动图片背景水印：wp:anchor 相对页面水平+垂直居中，behindDoc 置于文字后方。

    v1.3.63：新增——WPS 不渲染 VML 背景水印，改用标准 drawing 浮动图片
    （透明底红色大字图），Word/WPS/Pages 均显示，每页背景中央出现大红"试用版"。
    """
    cx = cy = str(_BG_EMU)
    return ('<w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" '
            'simplePos="0" relativeHeight="251658240" behindDoc="1" locked="0" '
            'layoutInCell="1" allowOverlap="1">'
            '<wp:simplePos x="0" y="0"/>'
            '<wp:positionH relativeFrom="page"><wp:align>center</wp:align></wp:positionH>'
            '<wp:positionV relativeFrom="page"><wp:align>center</wp:align></wp:positionV>'
            '<wp:extent cx="%s" cy="%s"/>'
            '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
            '<wp:wrapNone/>'
            '<wp:docPr id="1" name="TfdBgWm"/>'
            '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>'
            '<a:graphic><a:graphicData uri="%s">'
            '<pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="tfd_bg_wm"/><pic:cNvPicPr/></pic:nvPicPr>'
            '<pic:blipFill><a:blip r:embed="%s"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="%s" cy="%s"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'
            '</a:graphicData></a:graphic></wp:anchor></w:drawing>'
            % (cx, cy, PIC, _BG_IMG_REL_ID, cx, cy))


def _header_xml(text):
    """页眉 XML：VML 红色大背景水印(Word 增强层) + 浮动图片背景水印(通用核心层) + 红字。"""
    esc = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")
    shapes = "".join([
        _wm_shape(1, 0, 415, 160, 315),     # 上部：斜向
        _wm_shape(2, 180, 415, 160, 315),   # 中部：斜向
        _wm_shape(3, 340, 415, 160, 0),     # 下部：水平
    ]) % (esc, esc, esc)                    # 3 个 shape 各有一个 string 占位
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<w:hdr xmlns:w="%s" xmlns:r="%s" xmlns:wp="%s" xmlns:a="%s" xmlns:pic="%s" '
            'xmlns:v="urn:schemas-microsoft-com:vml" '
            'xmlns:o="urn:schemas-microsoft-com:office:office">'
            # ① VML 背景水印（Word 显示，每页 3 个大字）
            '<w:p><w:pPr><w:pStyle w:val="Header"/></w:pPr><w:r><w:rPr><w:noProof/></w:rPr>'
            '<w:pict>%s</w:pict></w:r></w:p>'
            # ② 浮动图片背景水印（Word/WPS/Pages 均显示，页面背景居中大红字）
            '<w:p><w:pPr><w:pStyle w:val="Header"/></w:pPr><w:r><w:rPr><w:noProof/></w:rPr>'
            '%s</w:r></w:p>'
            # ③ 页眉红字
            '<w:p><w:pPr><w:pStyle w:val="Header"/></w:pPr><w:r><w:rPr>'
            '<w:color w:val="C00000"/><w:sz w:val="20"/><w:b/></w:rPr>'
            '<w:t>%s</w:t></w:r></w:p></w:hdr>' % (W, R, WP, A, PIC, shapes,
                                                   _bg_anchor_xml(), esc))


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


def _is_wm_para(p):
    """判断段落是否已是水印段（文字水印行 或 水印图片段），避免重复计数/重复插入。"""
    if p.find(WR + "drawing") is not None:
        return True
    txt = "".join(t.text or "" for t in p.iter(WR + "t"))
    return _WM_BODY in txt


def _insert_body_paras(doc_root, step=8, cap=150):
    """正文穿插：按全文密度每 step 段插一处红色水印行（开头必插）。

    v1.3.63：修复"只覆盖前几页"根因——原实现 [:30] 硬上限，长论文(如石墨V7
    933段)只覆盖到第 232 段；改为按全文均匀分布，软上限 cap 防极端文档爆量。
    插入位置基于【原始正文段落】计算（过滤已插入的水印段），保证全文覆盖。
    """
    body = doc_root.find(WR + "body")
    if body is None:
        return 0
    paras = [p for p in body.findall(WR + "p") if not _is_wm_para(p)]
    if not paras:
        return 0
    positions = list(range(0, len(paras), step))[:cap]
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


def _image_para(doc_id):
    """含水印图片的段落（居中，带红色边框的水印图）。doc_id：文档级唯一图片 id。"""
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
    docpr.set("id", str(doc_id))
    docpr.set("name", "TFDWatermark")
    graphic = ET.SubElement(inline, AR + "graphic")
    gdata = ET.SubElement(graphic, AR + "graphicData")
    gdata.set("uri", PIC)
    pic = ET.SubElement(gdata, PICR + "pic")
    nv = ET.SubElement(pic, PICR + "nvPicPr")
    cnvpr = ET.SubElement(nv, PICR + "cNvPr")
    cnvpr.set("id", str(doc_id))
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


def _insert_image_paras(doc_root, step=14, cap=80):
    """正文按全文密度（每 step 段）插入水印图片段落（WPS/Word 必显示，逐张难删）。

    v1.3.63：修复"只覆盖前几页"——原实现 [:15] 硬上限，长论文只覆盖到第 196 段；
    改为全文均匀分布。水印图已缩小为 260x75(约7KB)，80 张约 0.6MB，体积可控。
    图片 id 从文档现有最大 docPr id 递增，避免与原文图片冲突。
    """
    body = doc_root.find(WR + "body")
    if body is None:
        return 0
    paras = [p for p in body.findall(WR + "p") if not _is_wm_para(p)]
    if not paras:
        return 0
    # 文档现有最大图片 id（wp:docPr / a:graphic 里的 cNvPr id）
    max_id = 0
    for el in body.iter():
        if el.tag in (WPR + "docPr",):
            try:
                max_id = max(max_id, int(el.get("id") or 0))
            except (TypeError, ValueError):
                pass
    positions = list(range(0, len(paras), step))[:cap]
    all_p = list(body)
    for pos in reversed(positions):
        idx = pos
        if idx >= len(all_p):
            idx = len(all_p)
            for j, el in enumerate(all_p):
                if el.tag == WR + "sectPr":
                    idx = j
                    break
        max_id += 1
        body.insert(idx, _image_para(max_id))
        all_p = list(body)
    return len(positions)


# CT_Settings 中排在 documentProtection 之后的元素（ECMA-376 §17.15.1.78 顺序），
# 用于确定 documentProtection 的合法插入位置（严格顺序，避免 Word 报损坏）。
_AFTER_DP = {
    "autoFormatOverride", "styleLockTheme", "styleLockQFSet", "defaultTabStop",
    "autoHyphenation", "consecutiveHyphenLimit", "hyphenationZone", "doNotHyphenateCaps",
    "showEnvelope", "summaryLength", "clickAndTypeStyle", "defaultTableStyle",
    "evenAndOddHeaders", "bookFoldRevPrinting", "bookFoldPrinting", "bookFoldPrintingSheets",
    "drawingGridHorizontalSpacing", "drawingGridVerticalSpacing",
    "displayHorizontalDrawingGridEvery", "displayVerticalDrawingGridEvery",
    "doNotUseMarginsForDrawingGridOrigin", "drawingGridHorizontalOrigin",
    "drawingGridVerticalOrigin", "doNotShadeFormData", "noPunctuationKerning",
    "characterSpacingControl", "printTwoOnOne", "strictFirstAndLastChars",
    "noLineBreaksAfter", "noLineBreaksBefore", "savePreviewPicture",
    "doNotValidateAgainstSchema", "saveInvalidXml", "ignoreMixedContent",
    "alwaysShowPlaceholderText", "doNotDemarcateInvalidXml", "saveXmlDataOnly",
    "useXSLTWhenSaving", "saveThroughXslt", "showXMLTags", "alwaysMergeEmptyNamespace",
    "updateFields", "hdrShapeDefaults", "footnotePr", "endnotePr", "compat", "rsids",
    "mathPr", "uiCompat97To2003",
}


def _readonly_protection_xml():
    """生成 w:documentProtection（只读+密码哈希）元素对象。

    v1.3.64：试用版文档只读，编辑需密码——防客户删除水印/篡改内容。
    算法：Word 2007+ 标准 legacy hash（SHA-1(salt+UTF16密码) + 100000 次迭代），
    Word/WPS 均支持。密码固定（_WM_PASSWORD），不对外公布。
    注：必须用 ET.Element 构建（自带命名空间），不能用字符串片段 fromstring
    （无 xmlns 声明会报 unbound prefix）。
    """
    salt = os.urandom(4)
    h = hashlib.sha1(salt + _WM_PASSWORD.encode("utf-16-le")).digest()
    for i in range(100000):
        h = hashlib.sha1(("%08x" % i).encode("ascii") + h).digest()
    el = ET.Element(WR + "documentProtection")
    el.set(WR + "edit", "readOnly")
    el.set(WR + "enforcement", "1")
    el.set(WR + "cryptProviderType", "rsaFull")
    el.set(WR + "cryptAlgorithmClass", "hash")
    el.set(WR + "cryptAlgorithmType", "typeAny")
    el.set(WR + "cryptAlgorithmSid", "4")
    el.set(WR + "cryptSpinCount", "100000")
    el.set(WR + "hash", base64.b64encode(h).decode("ascii"))
    el.set(WR + "salt", base64.b64encode(salt).decode("ascii"))
    return el


def _apply_readonly(parts):
    """在 word/settings.xml 注入只读保护（documentProtection）。"""
    settings_path = "word/settings.xml"
    settings_xml = parts.get(settings_path)
    if settings_xml is None:
        return False
    root = ET.fromstring(settings_xml)
    # 移除已有 documentProtection（防重复 / 替换文档自带保护）
    for dp in list(root.iter(WR + "documentProtection")):
        root.remove(dp)
    prot_el = _readonly_protection_xml()
    # 按 CT_Settings 顺序插入：找第一个"排在 documentProtection 之后"的元素，插到其前
    insert_at = None
    for i, el in enumerate(list(root)):
        if el.tag.replace(WR, "") in _AFTER_DP:
            insert_at = i
            break
    if insert_at is None:
        root.append(prot_el)
    else:
        root.insert(insert_at, prot_el)
    parts[settings_path] = _xml_str(root).encode("utf-8")
    return True


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

        # v1.3.64：只读保护（编辑需密码）——防客户删水印/改内容，正式版不调用本模块
        _apply_readonly(parts)

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

        # document.xml.rels 注册（header/footer/正文水印图/背景水印图）
        rels_path = "word/_rels/document.xml.rels"
        rels_xml = parts.get(rels_path, b"")
        if _HDR_REL_ID.encode("utf-8") not in rels_xml:
            rels_root = ET.fromstring(rels_xml) if rels_xml else ET.Element(PRR + "Relationships")
            for rid, target, typ in ((_HDR_REL_ID, "header1.xml",
                                      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"),
                                     (_FTR_REL_ID, "footer1.xml",
                                      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"),
                                     (_IMG_REL_ID, "media/" + _IMG_NAME,
                                      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"),
                                     (_BG_IMG_REL_ID, "media/" + _BG_IMG_NAME,
                                      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")):
                rel = ET.SubElement(rels_root, PRR + "Relationship")
                rel.set("Id", rid)
                rel.set("Type", typ)
                rel.set("Target", target)
            parts[rels_path] = _xml_str(rels_root).encode("utf-8")

        # 水印图片媒体文件（正文穿插图 + 页眉背景图）
        for name, src in ((_IMG_NAME, _IMG_SRC), (_BG_IMG_NAME, _BG_IMG_SRC)):
            if os.path.isfile(src):
                with open(src, "rb") as f:
                    parts["word/media/" + name] = f.read()

        # 重写 zip（新增条目：header1/footer1/media 两张水印图）
        # v1.3.63：用 set 去重——原文档可能已有 header1/footer1（如石墨V7），
        # 若重复写入会生成 Duplicate name zip 条目，Word/WPS 打开可能异常。
        extra = ["word/header1.xml", "word/footer1.xml",
                 "word/media/" + _IMG_NAME, "word/media/" + _BG_IMG_NAME]
        written = set()
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for n in names:
                if n in written:
                    continue
                zout.writestr(n, parts[n])
                written.add(n)
            for n in extra:
                if n in written:
                    continue
                zout.writestr(n, parts[n])
                written.add(n)
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
