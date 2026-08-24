# -*- coding: utf-8 -*-
# Copyright (c) 2026 芦苇（山东大学 MBA）. All rights reserved.
# 原创作品 | 禁止未经授权转售、二次分发或抄袭 | 授权用户可在扣子/虾评平台内使用
# build: 202608191649 | version: 1.3.5 | file_sha256: efa753e999913cb9
# -*- coding: utf-8 -*-
"""
docxutils.py — thesis-format-doctor 格式工具的共享底层库（零第三方依赖）。

只依赖 Python 标准库：zipfile / re / json / xml.etree.ElementTree。
.docx 本质是 ZIP 包，正文与样式分别在 word/document.xml、word/styles.xml。

本库提供给下列脚本复用：
  - format-profile.py   学校模板 -> 格式画像 profile.json
  - format-check.py     客户论文 vs 画像 -> 诊断报告
  - headings-fix.py     自动套标题样式 -> 新 docx（Word 可一键出目录）
  - ref-reformat.py     参考文献按范式重排 -> 新 docx
"""
import zipfile
import re
import xml.etree.ElementTree as ET

__author__ = "芦苇"
__copyright__ = "Copyright (c) 2026 芦苇（山东大学 MBA）"
__version__ = "1.3.5"

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WR = "{%s}" % W  # w: 前缀包裹器

# 图片相关命名空间（DrawingML / WordprocessingDrawing）
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
WPR = "{%s}" % WP  # wp: 前缀
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
AR = "{%s}" % A    # a: 前缀
PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
PICR = "{%s}" % PIC  # pic: 前缀

# EMU 换算常量
EMU_PER_TWIP = 635       # 1 inch = 1440 twips = 914400 EMU
EMU_PER_CM = 360000      # 1 cm = 360000 EMU
# A4 默认文本区宽度（EMU），作为 sectPr 缺失时的兜底
# A4: 21cm 宽, 左右各 3.17cm 边距 -> 文本区 ≈14.66cm
_DEFAULT_TEXT_WIDTH_EMU = 5277600



_INTEGRITY_TAG = "Q29weXJpZ2h0IChjKSAyMDI2IOiKpuiLh++8iOWxseS4nOWkp+WtpiBNQkHvvIkuIEFsbCByaWdodHMgcmVzZXJ2ZWQuIHwgYnVpbGQ6MjAyNjA4MTkxNjQ5"  # noqa
def reg_ns():
    """注册全部命名空间，避免写回 XML 时出现 ns0/ns1 前缀破坏 docx。"""
    ET.register_namespace("w", W)
    ET.register_namespace("wp", WP)
    ET.register_namespace("a", A)
    ET.register_namespace("pic", PIC)


# ---------------------------------------------------------------------------
# 读取
# ---------------------------------------------------------------------------
def load(path):
    """返回 (ZipFile, document.xml 的 ElementTree root)。"""
    z = zipfile.ZipFile(path)
    doc = z.read("word/document.xml").decode("utf-8")
    root = ET.fromstring(doc)
    return z, root


def load_styles(path):
    """读取 word/styles.xml，返回 {styleId: {name, basedOn, rPr, pPr}}。"""
    z = zipfile.ZipFile(path)
    try:
        sxml = z.read("word/styles.xml").decode("utf-8")
    except KeyError:
        return {}
    sroot = ET.fromstring(sxml)
    out = {}
    for st in sroot.findall(WR + "style"):
        sid = st.get(WR + "styleId")
        name_el = st.find(WR + "name")
        name = name_el.get(WR + "val") if name_el is not None else (sid or "")
        based = st.find(WR + "basedOn")
        basedOn = based.get(WR + "val") if based is not None else None
        rpr = _rpr_to_dict(st.find(WR + "rPr"))
        pinfo = _ppr_to_dict(st.find(WR + "pPr"))
        out[sid] = {"name": name, "basedOn": basedOn, "rPr": rpr, "pPr": pinfo}
    return out


def _rpr_to_dict(rpr):
    if rpr is None:
        return {}
    d = {}
    rfonts = rpr.find(WR + "rFonts")
    if rfonts is not None:
        d["ascii"] = rfonts.get(WR + "ascii")
        d["eastAsia"] = rfonts.get(WR + "eastAsia")
        d["hAnsi"] = rfonts.get(WR + "hAnsi")
    sz = rpr.find(WR + "sz")
    if sz is not None:
        d["sz"] = sz.get(WR + "val")  # 半磅，如 24 = 小二号
    b = rpr.find(WR + "b")
    if b is not None:
        d["bold"] = b.get(WR + "val") != "false"
    return d


def _ppr_to_dict(ppr):
    if ppr is None:
        return {}
    pinfo = {}
    ind = ppr.find(WR + "ind")
    if ind is not None:
        pinfo["firstLineChars"] = ind.get(WR + "firstLineChars")
        pinfo["firstLine"] = ind.get(WR + "firstLine")
    spacing = ppr.find(WR + "spacing")
    if spacing is not None:
        pinfo["line"] = spacing.get(WR + "line")
        pinfo["before"] = spacing.get(WR + "before")
        pinfo["after"] = spacing.get(WR + "after")
        pinfo["lineRule"] = spacing.get(WR + "lineRule")  # 行距规则：auto/exact/atLeast
    jc = ppr.find(WR + "jc")
    if jc is not None:
        pinfo["jc"] = jc.get(WR + "val")
    return pinfo


def paragraphs(root):
    """遍历正文段落，返回 [{text, style, rpr, pPr, in_table}]。style 为 w:pStyle 的 val。"""
    # 预先收集所有表格内段落元素 id（表格单元格里的 <w:p> 也要算段落）
    table_para_ids = set()
    for tbl in root.iter(WR + "tbl"):
        for cp in tbl.iter(WR + "p"):
            table_para_ids.add(id(cp))
    res = []
    for p in root.iter(WR + "p"):
        texts = "".join(t.text or "" for t in p.iter(WR + "t"))
        ppr = p.find(WR + "pPr")
        style = None
        pinfo = {}
        if ppr is not None:
            ps = ppr.find(WR + "pStyle")
            if ps is not None:
                style = ps.get(WR + "val")
            pinfo = _ppr_to_dict(ppr)
        # 取第一个非图片 run 的 rPr 作为段落字体/字号代表。
        # 含图 run（w:drawing / w:pict）的 rPr 是旧的图片属性，不代表文本格式；
        # 若取到它会导致 checker 误报"标题字号不符"（实为图片 run 的 sz）。
        first_run = None
        for r in p.iter(WR + "r"):
            _has_img = False
            for tag in (WR + "drawing", WR + "pict"):
                if len(list(r.iter(tag))) > 0:
                    _has_img = True
                    break
            if not _has_img:
                first_run = r
                break
        rpr = _rpr_to_dict(first_run.find(WR + "rPr") if first_run is not None else None)
        res.append({"text": texts, "style": style, "rpr": rpr, "pPr": pinfo,
                    "in_table": id(p) in table_para_ids})
    return res


def margins(root):
    sp = root.find(".//" + WR + "sectPr")
    if sp is None:
        return {}
    pg = sp.find(WR + "pgMar")
    if pg is None:
        return {}
    return {k: pg.get(WR + k) for k in
            ["top", "bottom", "left", "right", "header", "footer", "gutter"]
            if pg.get(WR + k) is not None}


def tables(root):
    """返回每个表格的边框字典 {edge: val}。"""
    out = []
    for tbl in root.iter(WR + "tbl"):
        borders = {}
        tblpr = tbl.find(WR + "tblPr")
        if tblpr is not None:
            bt = tblpr.find(WR + "tblBorders")
            if bt is not None:
                for edge in ["top", "left", "bottom", "right", "insideH", "insideV"]:
                    e = bt.find(WR + edge)
                    if e is not None:
                        val = e.get(WR + "val")
                        # 过滤 val="none"（无边框），否则会被误判为"有该边线"
                        if val and val != "none":
                            borders[edge] = val
        out.append(borders)
    return out


# ---------------------------------------------------------------------------
# 标题层级启发式识别（处理中文"第X章"、数字"1.1.1"、中文序号等）
# ---------------------------------------------------------------------------
def is_toc_residue(text):
    """判断段落文本是否为目录域/书签残留（.doc 转 .docx 时带入的脏数据）。

    这些段落会被 detect_heading 误判为标题（如 '1 绪论TC  "Chapter 1 Intro"'），
    若套样式会污染目录，必须跳过。判定依据：含 'TC ' 标记、英文目录项 'Chapter '、
    或目录域开关 '\\l '。
    """
    t = text or ""
    if re.search(r"TC\s", t):
        return True
    if '"Chapter' in t or "Chapter " in t:
        return True
    if re.search(r"\\l[\s\"]", t):
        return True
    return False


def style_xml(path, sid):
    """提取模板中某个 styleId 的 <w:style> 定义（字符串），供注入目标文档。"""
    reg_ns()
    z = zipfile.ZipFile(path)
    try:
        sxml = z.read("word/styles.xml").decode("utf-8")
    except KeyError:
        return ""
    sroot = ET.fromstring(sxml)
    for st in sroot.findall(WR + "style"):
        if st.get(WR + "styleId") == sid:
            return ET.tostring(st, encoding="unicode")
    return ""


def detect_heading(text):
    """
    返回 (level, confidence)。
    level: 1/2/3 或 None；confidence: 0~1。
    依据：编号格式 + 中文序数。仅做"像不像标题"的启发式判断，
    实际套样式时还需结合字号/加粗/段长等上下文（见 headings-fix.py）。
    """
    t = (text or "").strip()
    if not t:
        return (None, 0.0)
    # 第X章 / 第1篇（一级）
    if re.match(r"^第[一二三四五六七八九十百千\d]+[章篇编]", t):
        # 防御"列举性正文"被误判为章标题（用户实测踩坑）：正文里用"第一章、第二章"
        # 列举章节安排时，段落以"第X章"开头但本质是正文的枚举项，套章样式会污染目录。
        # 判定为非标题的两种情形：
        #  (a) 章/篇后紧跟顿号/逗号/分号（如"第一章、第二章……"）；
        #  (b) 一段内出现 ≥2 个章/节标记（如"第一章 绪论 第二章 文献综述"式目录残片）。
        # 去括号后再统计，避免"（含第三章内容）"这类正文叙述干扰真实标题判定。
        if re.match(r"^第[一二三四五六七八九十百千\d]+[章篇编][、，；,]", t):
            return (None, 0.0)
        _core = re.sub(r"[（(][^（）()]*[）)]", "", t)
        if len(re.findall(r"第[一二三四五六七八九十百千\d]+[章篇编]", _core)) >= 2:
            return (None, 0.0)
        return (1, 0.95)
    # 第X节（二级）
    if re.match(r"^第[一二三四五六七八九十百千\d]+节", t):
        if re.match(r"^第[一二三四五六七八九十百千\d]+节[、，；,]", t):
            return (None, 0.0)
        _core = re.sub(r"[（(][^（）()]*[）)]", "", t)
        if len(re.findall(r"第[一二三四五六七八九十百千\d]+节", _core)) >= 2:
            return (None, 0.0)
        return (2, 0.90)
    # 中文数字序数 + 顿号（一、绪论 / 二、文献综述）：论文常见的一级章标题写法。
    # v1.3.40：此前一律判 None（避免误套），导致整篇用"一、"写章标题的论文
    # first_chap 定位失败、修正静默不生效（输出与原文一致）。现识别为一级标题，
    # 置信度 0.85 与"第X章"同级；保留列举性防御：
    #   (a) 一段内出现 ≥2 个中文序数标记（"一、研究内容…二、…"式正文列举）→ 非标题；
    #   (b) 括号包裹的列举（"（一）…（二）…"）仍判 None（正文列举风险高，不硬编码级别）。
    if re.match(r"^[一二三四五六七八九十百千]+、", t):
        _core = re.sub(r"[（(][^（）()]*[）)]", "", t)
        if len(re.findall(r"[一二三四五六七八九十百千]+、", _core)) >= 2:
            return (None, 0.0)
        return (1, 0.85)
    # 1 / 1.1 / 1.1.1 / 1.1.1.1（多级编号）
    # v1.3.5：纯单级数字后必须紧跟分隔符（空格/顿号/点/冒号/制表符），不能直接跟中文，
    # 否则"20世纪""2018年""100强企业"等正文会被误判为一级标题。
    # 多级编号（含小数点如"1.1"）后可直接跟中文或分隔符（如"1.1研究背景""1.1 研究背景"），
    # 因带小数点的编号在正文叙述中极少出现，误判率低。
    if re.match(r"^\d{1,3}(?:\.\d{1,3})+", t):
        # 多级编号：1.1 / 1.1.1 等，后面可直接跟中文
        m = re.match(r"^(\d{1,3})(?:\.(\d{1,3}))?(?:\.(\d{1,3}))?(?:\.(\d{1,3}))?", t)
        if m:
            dots = sum(1 for x in [m.group(2), m.group(3), m.group(4)] if x)
            return (1 + dots, 0.95)
    else:
        # 单级数字：必须跟分隔符
        m = re.match(r"^(\d{1,3})(?=[\s、：:（(\t]|\.(?!\d))", t)
        if m:
            return (1, 0.95)
    # 中文序号「一、/二、」与「（一）/（二）」：学校模板（如清华）采用纯数字四级
    # （1 / 1.1 / 1.1.1 / 1.1.1.1），并不规定 一、/（一） 对应哪一级。按"格式严格按学校
    # 模板、skill 不自创级别"原则，此类序号不硬赋标题级别，判为正文（None）——避免把
    # 客户自创/列举性的中文序号误套成某级标题格式。若某学校模板确实以中文序号分级，
    # 应由该校模板/批注明确驱动，而非在此硬编码。
    # 阿拉伯数字/圈号括号序号（（1）(1) ① ② 等）→ 正文列举项，非标题。
    return (None, 0.0)


# ---------------------------------------------------------------------------
# 写回（修改 document.xml 后生成新 docx）
# ---------------------------------------------------------------------------
def write_docx(src_zip_path, dst_path, new_document_xml):
    """以 src_zip_path 为蓝本，把 document.xml 替换为 new_document_xml，写出 dst_path。"""
    write_docx_files(src_zip_path, dst_path, {"word/document.xml": new_document_xml})


def write_docx_files(src_zip_path, dst_path, replacements):
    """以 src_zip_path 为蓝本，按 replacements={内部路径: 新xml} 替换若干部件后写出 dst_path。

    用于「按模板应用样式」：同时替换 document.xml（套样式）与 styles.xml（注入模板样式定义）。
    replacements 中若包含源 zip 里不存在的新部件（如新建的 word/comments.xml），
    也会一并写入，确保新增部件不丢失。
    """
    reg_ns()
    with zipfile.ZipFile(src_zip_path) as zin:
        src_names = set(zin.namelist())
        with zipfile.ZipFile(dst_path, "w", zipfile.ZIP_DEFLATED) as zout:
            written = set()
            for item in zin.infolist():
                if item.filename in replacements:
                    zout.writestr(item, replacements[item.filename])
                else:
                    zout.writestr(item, zin.read(item.filename))
                written.add(item.filename)
            # 写入 replacements 中源 zip 不存在的新增部件（如 word/comments.xml）
            for name, data in replacements.items():
                if name not in written:
                    zout.writestr(name, data)


def to_doc_xml(root):
    """把修改后的 root 序列化为带声明头的 document.xml 字符串。"""
    reg_ns()
    body = ET.tostring(root, encoding="unicode")
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + body


# ---------------------------------------------------------------------------
# 批注（comments.xml）解析 —— 模板里"批注"往往写明最权威的格式要求
# （字号/字体/大纲级别/对齐/缩进/段间距/行距），应优先于样式定义采纳。
# ---------------------------------------------------------------------------
def cn2int(s):
    if s is None:
        return 0
    if str(s).isdigit():
        return int(s)
    d = {'一': 1, '二': 2, '三': 3, '四': 4, '五': 5, '六': 6,
         '七': 7, '八': 8, '九': 9, '十': 10}
    if s == '十':
        return 10
    if '十' in s:
        parts = s.split('十')
        left = d.get(parts[0], 1) if parts[0] else 1
        right = d.get(parts[1], 0) if parts[1] else 0
        return left * 10 + right
    return d.get(s, 0)


# 中文字号名 -> (半磅 sz 值, 磅 pt)。批注里写"小三号"等，需映射到 Word 的 sz 半磅。
SIZE_MAP = {
    '初号': (84, 42), '小初': (72, 36), '一号': (52, 26), '小一': (48, 24),
    '二号': (44, 22), '小二': (36, 18), '三号': (32, 16), '小三': (30, 15),
    '四号': (28, 14), '小四': (24, 12), '五号': (21, 10.5), '小五': (18, 9),
    '六号': (15, 7.5), '小六': (13, 6.5), '七号': (11, 5.5), '八号': (10, 5),
}
# 注意：每个字号的"小X"变体必须排在对应"X"之前，否则"小三号"会被"三号"抢先匹配。
_SIZE_ORDER = ['小初', '初号', '小一', '一号', '小二', '二号', '小三', '三号',
               '小四', '四号', '小五', '五号', '小六', '六号', '小七', '七号',
               '小八', '八号']


def comments(path):
    """读取 word/comments.xml，返回 [{id, author, text}]。无批注返回 []。"""
    z = zipfile.ZipFile(path)
    try:
        cxml = z.read("word/comments.xml").decode("utf-8")
    except KeyError:
        return []
    root = ET.fromstring(cxml)
    out = []
    for c in root.findall(WR + "comment"):
        cid = c.get(WR + "id")
        author = c.get(WR + "author") or ""
        txt = "".join(t.text or "" for t in c.iter(WR + "t"))
        out.append({"id": cid, "author": author, "text": txt.strip()})
    return out


# 参考文献批注分隔标记：山大等学校在一条批注里同时写明"参考文献标题"与
# "参考文献内容"两套格式，用"参考文献内容："、"内容："、"正文："等措辞分隔。
# 拆分后，前者归为 reference_heading（标题格式），后者归为 reference（条目格式）。
# 跨校通用——任何学校只要批注中出现此分隔标记即可自动拆分，未出现则按整段处理。
_REF_SPLIT_RE = re.compile(
    r'(?:参考文献|参考资料|引用文献|References|Bibliography)?\s*'
    r'(?:内容|正文|条目)\s*[：:]'
)


def split_reference_comment(text):
    """若批注中含"参考文献内容："等分隔标记，拆分为 (标题段, 内容段)；否则返回 (None, text)。

    山大模板批注原文（id=26）示例：
      "黑体小三号，加粗，居中无缩进，大纲级别1级，段前24磅，段后18磅，单倍行距。
       参考文献内容：中文字体为宋体，西文字体为Times New Roman，小四号，两端对齐，
       大纲级别正文文本，段前0行，段后0行，固定值20磅。编号1-9悬挂缩进0.6厘米..."
    拆分后：
      标题段 = "黑体小三号，加粗，居中无缩进，大纲级别1级，段前24磅，段后18磅，单倍行距。"
      内容段 = "中文字体为宋体，西文字体为Times New Roman，小四号，两端对齐，..."
    """
    t = text or ""
    m = _REF_SPLIT_RE.search(t)
    if not m:
        return None, t
    heading_part = t[:m.start()].rstrip("。.，,；; \t")
    content_part = t[m.end():].strip()
    return heading_part, content_part


def parse_comment_spec(text):
    """把一条批注正文解析为结构化格式要求 dict（只含命中的键）。

    可识别：中文字体、西文字体、字号(中文字号名)、加粗、大纲级别(N级/正文文本)、
    对齐(居中/两端/左/右)、首行缩进N字符、无缩进、悬挂缩进N厘米（支持多档）、
    段前/段后N磅、行距(单倍/固定值N磅/最小值N磅)。

    悬挂缩进支持多档（山大批注常见）：
      "编号1-9悬挂缩进0.6厘米编号10-99悬挂缩进0.74厘米编号100以上悬挂缩进0.9厘米"
    解析结果中 `hanging_cm` 取首档值（向后兼容），`hanging_tiers` 为完整档位列表：
      [{"digits": 1, "cm": 0.6}, {"digits": 2, "cm": 0.74}, {"digits": 3, "cm": 0.9}]
    """
    t = text or ""
    spec = {}
    # 中文字体
    mz = re.search(r'中文字体(?:为|是|[:：])\s*([\u4e00-\u9fff]{1,4})', t)
    if mz:
        spec["zh_font"] = mz.group(1)
    else:
        # 兜底：批注开头直接写"黑体小三号，加粗..."时，识别常见中文字体名。
        # 仅在已有字号/加粗等格式信号时触发，避免正文中出现"黑体"二字被误判。
        _KNOWN_ZH_FONTS = ('微软雅黑', '方正小标宋', '华文行楷', '华文新魏',
                           '黑体', '宋体', '楷体', '仿宋', '隶书', '幼圆')
        for fn in _KNOWN_ZH_FONTS:
            if fn in t and ('号' in t or 'pt' in t.lower() or '加粗' in t):
                # 字体名必须出现在字号名之前（批注惯例：字体→字号→加粗→对齐...）
                _size_pos = -1
                for sn in _SIZE_ORDER:
                    p = t.find(sn)
                    if p >= 0:
                        _size_pos = p if _size_pos < 0 else min(_size_pos, p)
                fn_pos = t.find(fn)
                if _size_pos < 0 or fn_pos < _size_pos:
                    spec["zh_font"] = fn
                    break
    # 西文字体
    me = re.search(r'西文字体(?:为|是|[:：])\s*([A-Za-z ]+?)(?:[，,。；;]|$)', t)
    if me:
        spec["en_font"] = me.group(1).strip()
    # 字号（中文字号名）
    for name in _SIZE_ORDER:
        if name in t:
            spec["size"] = name
            spec["sz"] = SIZE_MAP[name][0]
            break
    # 字号（原始磅值，如 13pt / 11pt）——不少学校直接写 pt，必须支持
    if "size" not in spec:
        mpt = re.search(r'(\d+(?:\.\d+)?)\s*[pP][tT]', t)
        if mpt:
            pt = float(mpt.group(1))
            spec["size"] = "%gpt" % pt
            spec["sz"] = int(round(pt * 2))
    # 加粗
    if '加粗' in t:
        spec["bold"] = True
    # 大纲级别
    mo = re.search(r'大纲级别\s*([0-9一二三四五])\s*级', t)
    if mo:
        spec["outline_level"] = cn2int(mo.group(1))
    elif '大纲级别正文文本' in t or '大纲级别正文' in t:
        spec["outline_level"] = 0  # 0 表示"正文文本"（无大纲级别）
    # 对齐
    if '居中' in t:
        spec["align"] = "center"
    elif '两端对齐' in t:
        spec["align"] = "both"
    elif '左对齐' in t:
        spec["align"] = "left"
    elif '右对齐' in t:
        spec["align"] = "right"
    # 缩进
    if '首行缩进' in t:
        mi = re.search(r'首行缩进\s*(\d+)\s*字符', t)
        spec["indent_chars"] = int(mi.group(1)) if mi else None
        spec["indent_type"] = "first"
    elif '无缩进' in t:
        spec["indent_chars"] = 0
        spec["indent_type"] = "none"
    mh = re.search(r'悬挂缩进\s*([\d.]+)\s*厘米', t)
    if mh:
        spec["hanging_cm"] = float(mh.group(1))
    mhc = re.search(r'悬挂缩进\s*([\d.]+)\s*字符', t)
    if mhc:
        spec["hanging_chars"] = float(mhc.group(1))
    # 多档悬挂缩进："编号1-9悬挂缩进0.6厘米编号10-99悬挂缩进0.74厘米编号100以上..."
    # digits=1/2/3 表示编号位数（1-9=1位，10-99=2位，100+=3位），用于按编号位数选档位。
    _tier_pat = re.compile(
        r'编号\s*(\d+)\s*[-~—–]\s*(?:\d+|以上)?\s*悬挂缩进\s*([\d.]+)\s*厘米'
    )
    _tiers = []
    for _tm in _tier_pat.finditer(t):
        _start = int(_tm.group(1))
        _cm = float(_tm.group(2))
        _digits = len(str(_start))
        # 100以上 形式：start=100, digits=3
        _tiers.append({"digits": _digits, "cm": _cm})
    # "编号100以上悬挂缩进0.9厘米" 这类只有起点没有终点的格式
    _above_pat = re.compile(r'编号\s*(\d+)\s*以上\s*悬挂缩进\s*([\d.]+)\s*厘米')
    for _am in _above_pat.finditer(t):
        _start = int(_am.group(1))
        _cm = float(_am.group(2))
        _digits = len(str(_start))
        if not any(x["digits"] == _digits for x in _tiers):
            _tiers.append({"digits": _digits, "cm": _cm})
    if _tiers:
        # 去重并按 digits 排序
        _seen = set()
        _uniq = []
        for x in sorted(_tiers, key=lambda d: d["digits"]):
            if x["digits"] not in _seen:
                _seen.add(x["digits"])
                _uniq.append(x)
        spec["hanging_tiers"] = _uniq
        if "hanging_cm" not in spec:
            spec["hanging_cm"] = _uniq[0]["cm"]
    # 段前/段后（兼容"段前6磅""段前空6磅""段前段后均空0磅"等多种措辞）
    mb = re.search(r'段前\s*空?\s*(\d+)\s*磅', t)
    if mb:
        spec["before_pt"] = int(mb.group(1))
    mae = re.search(r'段后\s*空?\s*(\d+)\s*磅', t)
    if mae:
        spec["after_pt"] = int(mae.group(1))
    mboth = re.search(r'段前段后均\s*空?\s*(\d+)\s*磅', t)
    if mboth:
        spec["before_pt"] = int(mboth.group(1))
        spec["after_pt"] = int(mboth.group(1))
    # 行距
    if '单倍行距' in t:
        spec["line_rule"] = "single"
    elif '固定值' in t:
        ml = re.search(r'固定值\s*(\d+)\s*磅', t)
        spec["line_rule"] = "exact"
        spec["line_val"] = int(ml.group(1)) if ml else None
    elif '最小值' in t:
        ml = re.search(r'最小值\s*(\d+)\s*磅', t)
        spec["line_rule"] = "atLeast"
        spec["line_val"] = int(ml.group(1)) if ml else None
    # 分页：另起一页 / 另起页 / 新起一页
    if re.search(r'另起.?一页|另起页|新起.?一页', t):
        spec["page_break_before"] = True
    return spec


def parse_page_spec(text):
    """解析"页边距：上—2.8 cm，下—2.5 cm…"类批注，返回 {top_cm,...}（缺省键不出现）。"""
    t = text or ""
    out = {}
    for key, pat in (("top_cm", r'上[—\-~]?\s*([\d.]+)\s*cm'),
                     ("bottom_cm", r'下[—\-~]?\s*([\d.]+)\s*cm'),
                     ("left_cm", r'左[—\-~]?\s*([\d.]+)\s*cm'),
                     ("right_cm", r'右[—\-~]?\s*([\d.]+)\s*cm'),
                     ("header_cm", r'页眉[边距]?\s*([\d.]+)\s*cm'),
                     ("footer_cm", r'页脚[边距]?\s*([\d.]+)\s*cm')):
        m = re.search(pat, t)
        if m:
            out[key] = float(m.group(1))
    return out


# 批注 -> 类别（用于把分散的批注归并到 一级标题/正文/页边距… 等规范维度）。
# 顺序即优先级（更具体的标签放前面）。
# 注意：这是「解析模板」环节的归并映射，必须从模板自身措辞识别类别——不同学校
# 模板用词不同（表题/表头/图表标题/图名…），这里集中维护同义别名，套用层不依赖
# 任何具体措辞。新增学校模板时，通常只需在此补别名即可，无需改套用逻辑。
_COMMENT_CATS = [
    ('h1', '一级标题'), ('h1', '章节标题'), ('h1', '章标题'),
    ('h3', '二级节标题'), ('h3', '二级节'),
    ('h2', '一级节标题'), ('h2', '一级节'),
    ('h2', '二级标题'), ('h2', '节标题'),
    ('h3', '三级标题'), ('h3', '小节标题'),
    ('conclusion', '结论标题'), ('ack', '致谢标题'), ('appendix', '附录标题'),
    ('toc', '目录'),
    ('figure_note', '图注'), ('figure_note', '图片来源'), ('figure_note', '图源'), ('figure_note', '图下说明'),
    ('figure', '图题'), ('figure', '图名'), ('figure', '图标题'), ('figure', '图表标题'), ('figure', '插图标题'),
    ('table_note', '表注'), ('table_note', '数据来源'), ('table_note', '资料来源'), ('table_note', '来源说明'), ('table_note', '注记'),
    ('table', '表题'), ('table', '表头'), ('table', '表标题'), ('table', '图表标题'),
    ('footnote', '脚注'),
    ('reference', '参考文献'), ('reference', '参考资料'), ('reference', '引用文献'),
    ('reference', 'References'), ('reference', 'Bibliography'),
    # reference_heading 仅用于批注拆分时显式归并（"参考文献内容："之前的标题段），
    # 普通批注里出现"参考文献"且不含分隔标记时仍归为 reference（向后兼容）。
    ('reference_heading', '__REF_HEADING__'),
    ('en_keywords', 'Key words'), ('keywords', '关键词'),
    ('en_abstract', '英文摘要'),
    ('abstract_title', '摘要标题'), ('abstract', '摘要'),
    ('title', '题目'), ('title', '论文题目'),
    ('page', '页边距'),
    ('body', '正文'),
]

# ---------------------------------------------------------------------------
# 跨校通用识别：结构页标题 / 参考文献区标题 / 正文样式名
# 不同学校/中英文模板用词不同（参考文献 / References / Bibliography / 参考资料…），
# 此处集中维护「关键词集合 + 大小写不敏感 + startswith 兜底」，所有调用方统一复用，
# 不再在各处散落写死中文关键词（这是跨校通用性的关键）。
# ---------------------------------------------------------------------------
REFERENCE_KEYWORDS = (
    "参考文献", "参考资料", "引用文献",
    "references", "reference", "bibliography", "bibliographies",
)
STRUCTURAL_KEYWORDS = (
    "结论", "致谢", "附录", "参考文献", "目录", "摘要",
    "abstract", "references", "reference", "bibliography",
    "conclusion", "acknowledgement", "acknowledgments",
    "acknowledgment", "appendix", "appendices", "contents", "summary",
)
BODY_STYLE_NAMES = (
    "正文", "正文文本", "normal", "body text", "bodytext", "body",
    "standard", "default paragraph font", "text body",
)


def _kw_hit(text, keywords):
    # v1.3.5：去除段内所有空白字符再比较。部分模板用"致    谢"等加空格方式
    # 视觉居中，直接 startswith 会匹配不到。
    t = re.sub(r"\s+", "", (text or "").strip().lower())
    if not t:
        return False
    for kw in keywords:
        k = re.sub(r"\s+", "", kw.lower())
        if t == k or t.startswith(k):
            return True
    return False


def is_reference_heading(text):
    """段落是否为参考文献区标题（中英文主流 + 异校异措辞）。跨校安全。"""
    return _kw_hit(text, REFERENCE_KEYWORDS)


def is_structural_title(text):
    """段落是否为结构页标题（结论/致谢/附录/参考文献/目录/摘要/ABSTRACT…）。跨校安全。"""
    return _kw_hit(text, STRUCTURAL_KEYWORDS)


def is_body_style_name(name):
    """样式名是否为正文样式（中英文 + 异校异名）。跨校安全。"""
    return _kw_hit(name, BODY_STYLE_NAMES)


def classify_structural_title(text, seen_en_abstract=False):
    """把结构页标题文本映射到 profile.levels 的类别 key；非结构页返回 None。

    v1.3.78 起自 headings_fix._classify_structural_title 迁入，供 format_profile
    （模板画像提取标题/正文双 spec）与 headings_fix（结构页格式套用 pass）单一来源复用，
    避免两份逻辑漂移。跨校通用——大小写不敏感、去空白、去首尾标点。
    """
    t = re.sub(r"\s+", "", (text or "").strip()).lower()
    t = t.strip("：:；;,.，。")
    # 摘要（含「摘要ABSTRACT」等中英混排标题）
    if t == "摘要" or (t.startswith("摘要") and len(t) <= 8):
        return "abstract"
    # 关键词 / Key words / Keywords（英文摘要之后的 Key words 归为 en_keywords）
    if t in ("关键词", "key words", "keywords"):
        return "en_keywords" if seen_en_abstract else "keywords"
    # 英文摘要 / Abstract
    if t == "英文摘要" or t == "abstract":
        return "en_abstract"
    # 致谢 / Acknowledgements / Acknowledgment
    if t in ("致谢", "acknowledgements", "acknowledgment", "acknowledgments"):
        return "ack"
    # 附录 / Appendix / Appendices（含 附录A / 附录1 等带字母/数字写法）
    # v1.3.78：变体必须限制标题长度——"附录A的内容…"这类正文以"附录A"开头，
    # 无限长会整段误判成标题并被套成标题格式（冒烟测试抓到的真实 bug）。
    if t == "附录" or t in ("appendix", "appendices") \
            or (re.match(r"^附录[a-z0-9一二三四五六七八九十]", t) and len(t) <= 10):
        return "appendix"
    return None


def categorize_comment(text):
    """返回批注所属类别 key；无法归类返回 'misc'。"""
    t = text or ""
    for key, token in _COMMENT_CATS:
        if token in t:
            return key
    return 'misc'


def extract_comment_specs(comments_list):
    """把批注列表归并为 类别->结构化规范。page 类别单独走 parse_page_spec。

    返回 {"cats": {类别: spec}, "page": {cm...} 或 None, "count": N}

    v1.3.4：支持参考文献批注拆分——若一条批注中含"参考文献内容："等分隔标记，
    拆分为 reference_heading（标题段）和 reference（条目段）两套 spec，避免
    标题与条目格式混淆（如标题被设成宋体小四、条目被设成小三加粗居中）。
    """
    cats = {}
    page = None
    for c in comments_list:
        text = c["text"]
        cat = categorize_comment(text)
        if cat == 'page':
            page = parse_page_spec(text)
            continue
        # v1.3.4：参考文献批注拆分（标题 spec / 内容 spec 分离）
        if cat == 'reference':
            heading_part, content_part = split_reference_comment(text)
            if heading_part:
                hspec = parse_comment_spec(heading_part)
                if hspec:
                    base = cats.get('reference_heading', {})
                    base.update(hspec)
                    cats['reference_heading'] = base
                cspec = parse_comment_spec(content_part) if content_part else None
                if cspec:
                    base = cats.get('reference', {})
                    base.update(cspec)
                    cats['reference'] = base
                continue
        if cat == 'misc':
            # 无明确标签：尝试用正文文本/首行缩进等弱信号归到 body，否则丢弃
            if '正文文本' in text and '首行缩进' in text:
                cat = 'body'
            else:
                continue
        spec = parse_comment_spec(text)
        if not spec:
            continue
        # 同类批注合并：后写覆盖同键（山大模板每类仅一条，天然幂等）
        base = cats.get(cat, {})
        base.update(spec)
        cats[cat] = base
    return {"cats": cats, "page": page, "count": len(comments_list)}


# ---------------------------------------------------------------------------
# 半角标点自动替换（正文段落保守替换，宁可漏换不可错换）
# ---------------------------------------------------------------------------
# 中文字符范围（含扩展A区），用于判断"中文上下文"
_CJK_RE = re.compile(r'[\u3400-\u9fff]')
# URL 中的协议前缀，其冒号不替换
_URL_SCHEME_RE = re.compile(r'[a-zA-Z][a-zA-Z0-9+.\-]*://')


def _is_cjk(ch):
    """单个字符是否为中文（含中日韩统一表意文字及扩展A区）。"""
    return bool(ch) and bool(_CJK_RE.match(ch))


def replace_punct_in_text(text):
    """对一段纯文本执行半角→全角标点保守替换。

    规则：
      - `,` → `，`（跳过数字中的千分位逗号，如 1,000）
      - `;` → `；`
      - `:` → `：`（跳过 URL 协议中的冒号，如 http://、https://）
      - `?` → `？`（仅当前一个字符为中文时替换）
      - `!` → `！`（同上）
      - `(` `)` → `（` `）`（括号内有中文才替换，纯英文括号不换）
    不替换英文句号 `.` 和引号。

    返回 (new_text, counts_dict)。counts_dict 形如 {"comma":3, "semicolon":1, ...}。
    """
    if not text:
        return text, {}
    counts = {}
    chars = list(text)
    n = len(chars)

    # 先定位 URL 协议中的冒号位置（如 http:// 中的 :），这些位置不替换
    skip_colon = set()
    for m in _URL_SCHEME_RE.finditer(text):
        # 协议末尾的 ":" 在 m.end()-2 位置（"://" 的冒号）
        colon_pos = m.end() - 3  # "://" 从该位置开始
        if 0 <= colon_pos < n:
            skip_colon.add(colon_pos)

    # 定位需要替换的括号对（括号内有中文才替换）
    paren_pairs = []  # [(open_idx, close_idx)]
    stack = []
    for i, ch in enumerate(chars):
        if ch == '(':
            stack.append(i)
        elif ch == ')' and stack:
            open_i = stack.pop()
            inner = text[open_i + 1:i]
            if _CJK_RE.search(inner):
                paren_pairs.append((open_i, i))
    replace_open = {o for o, _ in paren_pairs}
    replace_close = {c for _, c in paren_pairs}

    for i in range(n):
        ch = chars[i]
        if ch == ',' and _is_cjk(chars[i - 1] if i > 0 else ''):
            # 逗号：前一个字符是中文才替换；并跳过千分位逗号（后一个字符是数字）
            if i + 1 < n and chars[i + 1].isdigit() and i > 0 and chars[i - 1].isdigit():
                # 数字中的千分位逗号，如 1,000
                continue
            chars[i] = '，'
            counts['comma'] = counts.get('comma', 0) + 1
        elif ch == ';':
            prev_ch = chars[i - 1] if i > 0 else ''
            if _is_cjk(prev_ch):
                chars[i] = '；'
                counts['semicolon'] = counts.get('semicolon', 0) + 1
        elif ch == ':' and i not in skip_colon:
            prev_ch = chars[i - 1] if i > 0 else ''
            if _is_cjk(prev_ch):
                chars[i] = '：'
                counts['colon'] = counts.get('colon', 0) + 1
        elif ch in ('?', '!'):
            prev_ch = chars[i - 1] if i > 0 else ''
            if _is_cjk(prev_ch):
                chars[i] = '？' if ch == '?' else '！'
                key = 'question' if ch == '?' else 'exclaim'
                counts[key] = counts.get(key, 0) + 1
        elif ch == '(' and i in replace_open:
            chars[i] = '（'
            counts['parenthesis'] = counts.get('parenthesis', 0) + 1
        elif ch == ')' and i in replace_close:
            chars[i] = '）'
            counts['parenthesis'] = counts.get('parenthesis', 0) + 1

    return ''.join(chars), counts


# 标点类型的中文标签（用于报告与批注）
PUNCT_LABELS = {
    'comma': '半角逗号→全角逗号',
    'semicolon': '半角分号→全角分号',
    'colon': '半角冒号→全角冒号',
    'question': '半角问号→全角问号',
    'exclaim': '半角感叹号→全角感叹号',
    'parenthesis': '半角括号→全角括号',
}


def punct_counts_summary(counts):
    """把 counts_dict 拼成人类可读的明细，如"半角逗号→全角逗号 ×3；半角分号→全角分号 ×1"。"""
    parts = []
    for key in ('comma', 'semicolon', 'colon', 'question', 'exclaim', 'parenthesis'):
        if key in counts:
            parts.append('%s ×%d' % (PUNCT_LABELS[key], counts[key]))
    return '；'.join(parts)


def punct_total(counts):
    """counts_dict 中的替换总数。"""
    return sum(counts.values()) if counts else 0


def replace_punct_in_paragraph(p):
    """对一个段落 XML 元素（w:p）执行半角→全角标点替换。

    遍历段落内所有 w:t 元素，逐个文本节点替换（不跨 run 合并，
    避免破坏 run 级格式）。仅替换文本，不改变样式。

    返回 counts_dict（无替换时返回空 dict）。
    """
    total_counts = {}
    for t in p.iter(WR + 't'):
        old = t.text or ''
        new, c = replace_punct_in_text(old)
        if c:
            t.text = new
            for k, v in c.items():
                total_counts[k] = total_counts.get(k, 0) + v
    return total_counts


# ---------------------------------------------------------------------------
# Word 批注（comments.xml）—— 在修改过的段落旁添加批注，让用户看到改了什么
# ---------------------------------------------------------------------------
# Content Types / Relationships 命名空间
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
COMMENTS_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
COMMENTS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"


def _read_zip_xml(z, name):
    """从 zip 中读取并解析 XML，返回 ElementTree root；不存在返回 None。"""
    try:
        data = z.read(name)
    except KeyError:
        return None
    return ET.fromstring(data.decode('utf-8'))


def max_comment_id(z):
    """读取文档现有批注的最大 ID（整数）；无批注返回 -1。新批注应从此值+1 开始。"""
    root = _read_zip_xml(z, 'word/comments.xml')
    if root is None:
        return -1
    max_id = -1
    for c in root.findall(WR + 'comment'):
        cid = c.get(WR + 'id')
        if cid is not None:
            try:
                max_id = max(max_id, int(cid))
            except ValueError:
                pass
    return max_id


def _ensure_comments_xml(z):
    """确保 word/comments.xml 存在并返回其 root（已注册命名空间）。

    若文档原本没有 comments.xml，则创建一个空的根元素：
      <w:comments xmlns:w="..."/>
    """
    reg_ns()
    root = _read_zip_xml(z, 'word/comments.xml')
    if root is not None:
        return root
    return ET.Element(WR + 'comments')


def _ensure_content_types_comments(ct_root):
    """在 [Content_Types].xml 中注册 comments.xml 的 Override（若尚未注册）。"""
    ET.register_namespace('', CT_NS)
    for ov in ct_root.findall('{%s}Override' % CT_NS):
        if ov.get('PartName') == '/word/comments.xml':
            return
    ov = ET.SubElement(ct_root, '{%s}Override' % CT_NS)
    ov.set('PartName', '/word/comments.xml')
    ov.set('ContentType', COMMENTS_CT)


def _ensure_doc_rels_comments(rels_root):
    """在 word/_rels/document.xml.rels 中添加 comments.xml 的 relationship。

    自动分配一个不与现有 Id 冲突的 rIdN。返回该 relationship 的 Id。
    """
    ET.register_namespace('', REL_NS)
    existing_ids = set()
    has_comments_rel = False
    for rel in rels_root.findall('{%s}Relationship' % REL_NS):
        rid = rel.get('Id')
        if rid:
            existing_ids.add(rid)
        if rel.get('Type') == COMMENTS_REL:
            has_comments_rel = True
            if rid:
                return rid
    if has_comments_rel:
        return None
    # 分配新 Id
    n = 1
    while ('rId%d' % n) in existing_ids:
        n += 1
    new_id = 'rId%d' % n
    rel = ET.SubElement(rels_root, '{%s}Relationship' % REL_NS)
    rel.set('Id', new_id)
    rel.set('Type', COMMENTS_REL)
    rel.set('Target', 'comments.xml')
    return new_id


def ensure_comments_part(z, replacements):
    """确保 comments.xml 部件存在并在 Content_Types / document.xml.rels 中注册。

    z: 源 docx 的 ZipFile
    replacements: dict {内部路径: 新 XML 字符串}，本函数会把需要新建/修改的
                  部件写入其中（word/comments.xml、[Content_Types].xml、
                  word/_rels/document.xml.rels）。

    返回 comments.xml 的 root（Element），供后续 add_comment_marker 追加批注。
    若文档已有 comments.xml，直接返回已解析的 root。
    """
    reg_ns()
    comments_root = _ensure_comments_xml(z)

    # 注册 [Content_Types].xml
    ct_name = '[Content_Types].xml'
    ct_root = _read_zip_xml(z, ct_name)
    if ct_root is None:
        ET.register_namespace('', CT_NS)
        ct_root = ET.Element('{%s}Types' % CT_NS)
    _ensure_content_types_comments(ct_root)
    # 使用 to_doc_xml 输出带 XML 声明头的字符串，提升 WPS 等严格解析器的兼容性。
    replacements[ct_name] = to_doc_xml(ct_root)

    # 注册 word/_rels/document.xml.rels
    rels_name = 'word/_rels/document.xml.rels'
    rels_root = _read_zip_xml(z, rels_name)
    if rels_root is None:
        ET.register_namespace('', REL_NS)
        rels_root = ET.Element('{%s}Relationships' % REL_NS)
    _ensure_doc_rels_comments(rels_root)
    replacements[rels_name] = to_doc_xml(rels_root)

    return comments_root


def add_comment_marker(para_elem, comment_id, comment_text, comments_root,
                       author="论文格式医生", initials="LGFD"):
    """在一个段落上插入批注标记，并在 comments_root 中添加批注内容。

    在段落开头插入 <w:commentRangeStart w:id="N"/>，
    在段落末尾插入 <w:commentRangeEnd w:id="N"/>，
    并在 commentRangeEnd 后插入一个含 <w:commentReference w:id="N"/> 的 run。
    同时在 comments_root 中追加 <w:comment> 元素。
    """
    reg_ns()
    cid = str(comment_id)

    # 段落开头插入 commentRangeStart（插在 pPr 之后、其余子元素之前）
    range_start = ET.Element(WR + 'commentRangeStart')
    range_start.set(WR + 'id', cid)
    ppr = para_elem.find(WR + 'pPr')
    if ppr is not None:
        idx = list(para_elem).index(ppr) + 1
    else:
        idx = 0
    para_elem.insert(idx, range_start)

    # 段落末尾插入 commentRangeEnd
    range_end = ET.Element(WR + 'commentRangeEnd')
    range_end.set(WR + 'id', cid)
    para_elem.append(range_end)

    # commentReference run
    ref_run = ET.Element(WR + 'r')
    ref_rpr = ET.SubElement(ref_run, WR + 'rPr')
    rstyle = ET.SubElement(ref_rpr, WR + 'rStyle')
    rstyle.set(WR + 'val', 'CommentReference')
    cref = ET.SubElement(ref_run, WR + 'commentReference')
    cref.set(WR + 'id', cid)
    para_elem.append(ref_run)

    # 在 comments_root 中添加 <w:comment>
    import datetime
    now = datetime.datetime.now().strftime('%Y-%m-%dT%H:%M:%SZ')
    comment_el = ET.SubElement(comments_root, WR + 'comment')
    comment_el.set(WR + 'id', cid)
    comment_el.set(WR + 'author', author)
    comment_el.set(WR + 'date', now)
    comment_el.set(WR + 'initials', initials)
    cp = ET.SubElement(comment_el, WR + 'p')
    cr = ET.SubElement(cp, WR + 'r')
    ct = ET.SubElement(cr, WR + 't')
    ct.text = comment_text
    ct.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
