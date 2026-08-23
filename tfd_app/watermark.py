# -*- coding: utf-8 -*-
"""
试用水印模块 · 论文格式医生
==================================================================
对"修正后的 docx"做后处理（纯标准库 zipfile + ElementTree），注入：
  1. 页眉 + 页脚：每页顶部/底部灰字"试用版 · 论文格式医生（正式版无水印）"；
  2. 正文穿插：正文开头 + 参考文献前（或末尾）插入显眼的【试用版】文字行。

正式版（有激活码）不调用本模块，输出无水印文档。
"""
import os
import zipfile
import xml.etree.ElementTree as ET

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
PR = "http://schemas.openxmlformats.org/package/2006/relationships"

WR = "{%s}" % W   # w: 前缀包裹器（ElementTree tag 需花括号形式）
RR = "{%s}" % R
CTR = "{%s}" % CT
PRR = "{%s}" % PR

ET.register_namespace("w", W)
ET.register_namespace("r", R)
ET.register_namespace("ct", CT)
ET.register_namespace("pr", PR)

_WM_HEADER = "试用版 · 论文格式医生（正式版无水印）"
_WM_BODY = "【试用版】论文格式医生 · 一键按学校模板修正格式（正式版无水印）"

_HDR_REL_ID = "rIdTfdHdr"
_FTR_REL_ID = "rIdTfdFtr"


def _wm_run(text, sz="16", color="999999"):
    """页眉/页脚里的灰字 run。"""
    r = ET.Element(WR + "r")
    rpr = ET.SubElement(r, WR + "rPr")
    c = ET.SubElement(rpr, WR + "color")
    c.set(WR + "val", color)
    s = ET.SubElement(rpr, WR + "sz")
    s.set(WR + "val", sz)            # 8pt（页眉页脚用）
    s2 = ET.SubElement(rpr, WR + "szCs")
    s2.set(WR + "val", sz)
    t = ET.SubElement(r, WR + "t")
    t.text = text
    return r


def _wm_body_para():
    """正文穿插的显眼试用水印段落（居中、灰字、小号）。"""
    p = ET.Element(WR + "p")
    ppr = ET.SubElement(p, WR + "pPr")
    jc = ET.SubElement(ppr, WR + "jc")
    jc.set(WR + "val", "center")
    r = ET.SubElement(p, WR + "r")
    rpr = ET.SubElement(r, WR + "rPr")
    c = ET.SubElement(rpr, WR + "color")
    c.set(WR + "val", "A6A6A6")
    s = ET.SubElement(rpr, WR + "sz")
    s.set(WR + "val", "18")          # 9pt
    s2 = ET.SubElement(rpr, WR + "szCs")
    s2.set(WR + "val", "18")
    i = ET.SubElement(rpr, WR + "i")
    t = ET.SubElement(r, WR + "t")
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = _WM_BODY
    return p


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
    """正文穿插：开头 1 处 + 参考文献前 1 处（找不到则插到末尾 sectPr 前）。"""
    body = doc_root.find(WR + "body")
    if body is None:
        return 0
    p1 = _wm_body_para()
    p2 = _wm_body_para()
    body.insert(0, p1)
    # 找"参考文献"段落，插到其前
    ref_p = None
    for p in body.findall(WR + "p"):
        txt = "".join(t.text or "" for t in p.iter(WR + "t"))
        if "参考文献" in txt:
            ref_p = p
            break
    if ref_p is not None:
        idx = list(body).index(ref_p)
        body.insert(idx, p2)
    else:
        # 插到最后一个非 sectPr 元素之后（sectPr 之前）
        idx = len(list(body))
        for i, el in enumerate(list(body)):
            if el.tag == WR + "sectPr":
                idx = i
                break
        body.insert(idx, p2)
    return 2


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
        _insert_sect_refs(doc_root)
        parts["word/document.xml"] = _xml_str(doc_root).encode("utf-8")

        # 页眉 / 页脚
        parts["word/header1.xml"] = _xml_str(
            _header_footer("hdr", _WM_HEADER)).encode("utf-8")
        parts["word/footer1.xml"] = _xml_str(
            _header_footer("ftr", _WM_HEADER)).encode("utf-8")

        # Content_Types 注册
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

        # document.xml.rels 注册
        rels_path = "word/_rels/document.xml.rels"
        rels_xml = parts.get(rels_path, b"")
        if _HDR_REL_ID.encode("utf-8") not in rels_xml:
            rels_root = ET.fromstring(rels_xml) if rels_xml else ET.Element(PRR + "Relationships")
            for rid, target, typ in ((_HDR_REL_ID, "header1.xml",
                                      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"),
                                     (_FTR_REL_ID, "footer1.xml",
                                      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer")):
                rel = ET.SubElement(rels_root, PRR + "Relationship")
                rel.set("Id", rid)
                rel.set("Type", typ)
                rel.set("Target", target)
            parts[rels_path] = _xml_str(rels_root).encode("utf-8")

        # 重写 zip（注意：header1/footer1 是新增条目，不在原 names 里，需一并写入）
        extra = ["word/header1.xml", "word/footer1.xml"]
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
