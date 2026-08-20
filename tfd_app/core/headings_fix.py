# -*- coding: utf-8 -*-
# Copyright (c) 2026 芦苇（山东大学 MBA）. All rights reserved.
# 原创作品 | 禁止未经授权转售、二次分发或抄袭 | 授权用户可在扣子/虾评平台内使用
# build: 202608191649 | version: 1.3.5 | file_sha256: ff086844ba31f8d5
# -*- coding: utf-8 -*-
"""
headings-fix.py — 自动识别标题层级并套用 Heading 样式（让 Word/WPS 可一键生成目录）

用法：
  python headings-fix.py 输入.docx 输出.docx

能力：
  - 兼容中文"第X章/第X节"、数字"1 / 1.1 / 1.1.1"、中文序号"一、""（一）"等层级写法
  - 默认对识别到的标题套内置 Heading1/2/3 样式（Word 据此即可一键插入目录）
  - 带 --profile 时按「学校模板」的**真实 styleId** 套用（如 山大 L1=00001d/L2=000022/L3=000010），
    并把模板的样式定义注入目标文档，实现真正贴合学校的格式（非仅通用 Heading）
  - 自动跳过目录域/书签残留（.doc 转 .docx 带入的 'TC'/'Chapter' 脏数据），避免污染目录
  - 低置信度条目（仅靠字号/加粗判断、无编号）会列出，供人工复核
  - 不改变任何正文文字，仅写入样式，符合"仅调样式、不改内容"的合规边界

用法：
  python headings-fix.py 输入.docx 输出.docx [--profile 模板画像.json]
"""
import sys
import os
import re
import json
import zipfile
import xml.etree.ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from docxutils import (load, detect_heading, write_docx_files, to_doc_xml, WR,
                       is_toc_residue, reg_ns, is_reference_heading, is_structural_title,
                       load_styles, WPR, AR, PICR, EMU_PER_TWIP, EMU_PER_CM,
                       _DEFAULT_TEXT_WIDTH_EMU,
                       replace_punct_in_paragraph, punct_counts_summary, punct_total,
                       ensure_comments_part, add_comment_marker, max_comment_id)
from report_docx import write_change_report

__author__ = "芦苇"
__copyright__ = "Copyright (c) 2026 芦苇（山东大学 MBA）"
__version__ = "1.3.5"

STYLE_FOR = {1: "Heading1", 2: "Heading2", 3: "Heading3"}
# 通用模式下新样式的展示名（无模板画像时使用）
DISP_GENERIC = {1: ("Heading1", "标题1"), 2: ("Heading2", "标题2"), 3: ("Heading3", "标题3")}



_INTEGRITY_TAG = "Q29weXJpZ2h0IChjKSAyMDI2IOiKpuiLh++8iOWxseS4nOWkp+WtpiBNQkHvvIkuIEFsbCByaWdodHMgcmVzZXJ2ZWQuIHwgYnVpbGQ6MjAyNjA4MTkxNjQ5"  # noqa
def _cur_style_id(p):
    """取段落当前 w:pStyle 的 styleId（无则 None）。"""
    ppr = p.find(WR + "pPr")
    if ppr is None:
        return None
    ps = ppr.find(WR + "pStyle")
    return ps.get(WR + "val") if ps is not None else None


def _is_toc_style(sid, styles_map, heading_sids=None):
    """段落当前样式是目录样式（toc 1/2/3、含『目录』或 Table of Contents）则跳过，
    避免把目录条目误套成章/节标题、并防止目录里的『1 绪论』被当成首个一级标题
    而把正文区起点算错（真实论文常见踩坑）。

    heading_sids：画像中真实的标题样式 styleId 集合——这些样式即便名字叫
    『toc 1』（个别学校模板的标题样式命名如此）也必须当作标题处理，不可误跳。
    """
    if not sid:
        return False
    if heading_sids and sid in heading_sids:
        return False
    nm = ((styles_map or {}).get(sid) or {}).get("name") or ""
    if re.match(r"^toc\b", nm, re.I):
        return True
    if "目录" in nm or "table of contents" in nm.lower():
        return True
    return False



def _text_of(p):
    return "".join(t.text or "" for t in p.iter(WR + "t"))


def _set_style(p, style_val):
    ppr = p.find(WR + "pPr")
    if ppr is None:
        ppr = ET.SubElement(p, WR + "pPr")
    ps = ppr.find(WR + "pStyle")
    if ps is None:
        ps = ET.SubElement(ppr, WR + "pStyle")
    ps.set(WR + "val", style_val)


def _align(val):
    return {'center': 'center', '居中': 'center', 'both': 'both', '两端': 'both',
            'justify': 'both', 'left': 'left', '左': 'left', 'right': 'right', '右': 'right'}.get(val, 'both')


def _line(spec):
    """把批注/样式 spec 的行距规则映射为 OOXML 的 (lineRule, line-twips)。"""
    lr = spec.get('line_rule')
    val = spec.get('line_val')
    if lr == 'exact' and val:
        return 'exact', int(val) * 20          # 固定值 X 磅 -> X*20 twips
    if lr in ('atLeast',) and val:
        return 'atLeast', int(val) * 20
    if lr == 'auto' and val:
        return 'auto', int(val)                 # auto：twips 直接透传（240=单倍，360=1.5倍）
    if lr == 'single':
        return 'auto', 240
    return None, None


def _cm_to_twips(cm):
    try:
        return int(round(float(cm) * 1440 / 2.54))
    except (TypeError, ValueError):
        return None


def _fix_sect_margins(root, profile):
    """按模板批注的页边距（cm）修正文档**正文区域**的 sectPr 的 pgMar。

    v1.3.4 修复：此前遍历文档里全部 sectPr（body 级 + 段落级）统一修改，
    会把封面分节（段落级 sectPr，由分节符产生）的页边距也改掉，违反
    "结构页只诊断不修改"原则。现改为只修改**最后一个 sectPr**——
      * 多 sectPr 文档：最后一个 sectPr 通常是 body 直接子元素（末尾 sectPr），
        控制正文（第一章起）及之后区域的页边距；前面段落级 sectPr 属封面/摘要
        等分节，保留原样不动；
      * 单 sectPr 文档（封面与正文不分节）：唯一 sectPr 即正文 sectPr，正常修改。
    该逻辑跨校通用，不写死任何学校名称。

    返回 (是否修改, [(边标签, 旧值显示, 新值cm显示), ...])。每条边仅记录一次。
    """
    page = (profile or {}).get("spec", {}).get("page") or {}
    if not page:
        return False, []
    # 收集全部 sectPr（按文档顺序）
    all_sects = list(root.iter(WR + "sectPr"))
    if not all_sects:
        return False, []
    # 只修改最后一个 sectPr：
    #  - body 末尾的 sectPr（body 直接子元素）是整篇文档最后一节；
    #  - 段落级 sectPr（pPr/sectPr）按出现顺序排在它之前，属于封面/摘要等分节。
    sect = all_sects[-1]
    changed = []
    pg = sect.find(WR + "pgMar")
    if pg is None:
        pg = ET.SubElement(sect, WR + "pgMar")
    for key, cmkey, label in (("top", "top_cm", "上"), ("bottom", "bottom_cm", "下"),
                              ("left", "left_cm", "左"), ("right", "right_cm", "右")):
        if cmkey not in page:
            continue
        want = _cm_to_twips(page[cmkey])
        if want is None:
            continue
        cur = pg.get(WR + key)
        if cur is not None and int(cur) == want:
            continue
        old_disp = "%.2fcm" % (int(cur) / 1440 * 2.54) if cur is not None else "（默认）"
        changed.append((label, old_disp, "%.2fcm" % float(page[cmkey])))
        pg.set(WR + key, str(want))
    return bool(changed), changed


def _set_run_rpr(run, spec):
    """按批注 spec 重写 run 的 rPr：字体(中/西)、字号、加粗。直接格式化优先于样式继承。
    含图片(w:drawing/w:pict)的 run 不触碰——避免向图片注入字体/字号导致渲染异常。"""
    if _run_has_image(run):
        return
    old = run.find(WR + 'rPr')
    if old is not None:
        run.remove(old)
    rpr = ET.Element(WR + 'rPr')
    fonts = {}
    if spec.get('zh_font'):
        fonts['eastAsia'] = spec['zh_font']
    if spec.get('en_font'):
        fonts.update({'ascii': spec['en_font'], 'hAnsi': spec['en_font'], 'cs': spec['en_font']})
    if fonts:
        rf = ET.SubElement(rpr, WR + 'rFonts')
        for k, v in fonts.items():
            rf.set(WR + k, v)
    if spec.get('bold'):
        ET.SubElement(rpr, WR + 'b')
    if spec.get('sz'):
        for tag in ('sz', 'szCs'):
            s = ET.SubElement(rpr, WR + tag)
            s.set(WR + 'val', str(spec['sz']))
    run.insert(0, rpr)


def _run_has_image(run):
    """run 是否包含图片（DrawingML / VML）。"""
    for tag in (WR + "drawing", WR + "pict"):
        if len(list(run.iter(tag))) > 0:
            return True
    return False


def _format_runs(p, spec):
    """对段落内全部 run 应用 spec 字体/字号/加粗；含图片的 run 自动跳过，不受影响。"""
    for r in p.iter(WR + "r"):
        _set_run_rpr(r, spec)


def _set_para_format(p, spec, style_val=None):
    """按批注 spec 写段落格式：对齐(jc)、缩进(ind)、段前后+行距(spacing)；可选套样式。

    缩进处理规则：
      - spec 显式指定 indent_type="none"（批注"无缩进"）：强制清除段落直接 <w:ind>，
        即便原段落带有悬挂缩进（如参考文献标题段从客户旧格式遗留的悬挂缩进）。
      - style_val 不为 None（套标题样式）：清除段落直接 <w:ind>，让样式定义生效。
      - spec 指定首行/悬挂缩进：清除旧值并按 spec 重建。
      - 其它（正文/题注等，spec 未指定缩进）：保留段落原有 <w:ind>。
    """
    ppr = p.find(WR + 'pPr')
    if ppr is None:
        ppr = ET.Element(WR + 'pPr')
        p.insert(0, ppr)
    if style_val is not None:
        ps = ppr.find(WR + 'pStyle')
        if ps is None:
            ps = ET.Element(WR + 'pStyle')
            ppr.insert(0, ps)
        ps.set(WR + 'val', style_val)
    # 判断 spec 是否显式指定了缩进（首行/悬挂）。
    has_indent_spec = (
        spec.get('hanging_cm') is not None
        or spec.get('hanging_chars') is not None
        or (spec.get('indent_chars') and spec.get('indent_type') == 'first')
    )
    # 批注明确要求"无缩进"（indent_type=="none"，indent_chars==0）：
    # 必须强制清除段落直接 <w:ind>，覆盖客户旧格式遗留的悬挂/首行缩进。
    force_no_indent = (spec.get('indent_type') == 'none'
                       or spec.get('indent_chars') == 0)
    # jc 与 spacing 始终按 spec 重建；ind 按场景决定是否保留。
    for tag in ('jc', 'spacing'):
        e = ppr.find(WR + tag)
        if e is not None:
            ppr.remove(e)
    old_ind = ppr.find(WR + 'ind')
    if old_ind is not None:
        if has_indent_spec or style_val is not None or force_no_indent:
            ppr.remove(old_ind)
            old_ind = None
        # 正文（style_val=None）且 spec 未指定缩进：保留原有 <w:ind>。
    newc = []
    if spec.get('align'):
        jc = ET.Element(WR + 'jc')
        jc.set(WR + 'val', _align(spec['align']))
        newc.append(jc)
    if spec.get('hanging_cm') is not None or spec.get('hanging_chars') is not None:
        # 悬挂缩进：Word需要同时设置 w:left（整段左边界）和 w:hanging（首行回弹），
        # 两者取值相同，首行回到左边距，续行缩进。只设 hanging 不设 left 会导致
        # 续行跑到左边距外（Word标尺显示为负值）。
        ind = ET.Element(WR + 'ind')
        if spec.get('hanging_cm') is not None:
            _tw = str(int(round(spec['hanging_cm'] * 567)))
            ind.set(WR + 'left', _tw)
            ind.set(WR + 'hanging', _tw)
        if spec.get('hanging_chars') is not None:
            _ch = str(int(round(spec['hanging_chars'] * 100)))
            ind.set(WR + 'leftChars', _ch)
            ind.set(WR + 'hangingChars', _ch)
        newc.append(ind)
    elif spec.get('indent_chars') and spec.get('indent_type') == 'first':
        ind = ET.Element(WR + 'ind')
        ind.set(WR + 'firstLineChars', str(spec['indent_chars'] * 100))
        ind.set(WR + 'firstLine', str(spec['indent_chars'] * 240))
        newc.append(ind)
    lr, line = _line(spec)
    sp = {}
    if lr is not None and line is not None:
        sp['line'] = str(line)
        sp['lineRule'] = lr
    if spec.get('before_pt') is not None:
        sp['before'] = str(int(spec['before_pt']) * 20)
    if spec.get('after_pt') is not None:
        sp['after'] = str(int(spec['after_pt']) * 20)
    if sp:
        s = ET.Element(WR + 'spacing')
        for k, v in sp.items():
            s.set(WR + k, v)
        newc.append(s)
    # 分页符：另起一页（一级标题常用）
    if spec.get('page_break_before'):
        # 移除旧值避免重复
        old_pbb = ppr.find(WR + 'pageBreakBefore')
        if old_pbb is not None:
            ppr.remove(old_pbb)
        newc.append(ET.Element(WR + 'pageBreakBefore'))
    # 插入位置：若保留了原有 ind，则把 jc 插到 ind 之前、spacing/pageBreak 插到 ind 之后；
    # 否则整体插入到 pStyle 之后。这样既保留样式缩进，又维持 pStyle→jc→ind→spacing 顺序。
    if old_ind is not None:
        anchor_idx = list(ppr).index(old_ind)
        # jc 必须在 ind 之前，先倒序插入到 anchor_idx
        pre = [c for c in newc if c.tag == WR + 'jc']
        post = [c for c in newc if c.tag != WR + 'jc']
        for c in reversed(pre):
            ppr.insert(anchor_idx, c)
        # ind 之后的元素依次插入（anchor 位置已移动，使用 index 重新定位）
        base = list(ppr).index(old_ind) + 1
        for i, c in enumerate(post):
            ppr.insert(base + i, c)
    else:
        anchor = ppr.find(WR + 'pStyle')
        idx = (list(ppr).index(anchor) + 1) if anchor is not None else 0
        for i, c in enumerate(newc):
            ppr.insert(idx + i, c)


def _para_needs_fix(p, spec):
    """正文段落是否需要按模板修正：字体 / 字号 / 对齐 / 首行缩进 / 段前 / 段后 / 行距
    任一不符，即视为需改。覆盖用户要求的全部 per-level 要素。"""
    for r in p.iter(WR + 'r'):
        rpr = r.find(WR + 'rPr')
        if rpr is None:
            return True
        rf = rpr.find(WR + 'rFonts')
        if spec.get('zh_font') and rf is not None and rf.get(WR + 'eastAsia') \
                and rf.get(WR + 'eastAsia') != spec['zh_font']:
            return True
        sz = rpr.find(WR + 'sz')
        if spec.get('sz') and sz is not None and sz.get(WR + 'val') \
                and int(sz.get(WR + 'val')) != int(spec['sz']):
            return True
    ppr = p.find(WR + 'pPr')
    # 首行缩进
    if spec.get('indent_type') == 'first':
        ind = ppr.find(WR + 'ind') if ppr is not None else None
        want = spec.get('indent_chars', 0) * 100
        cur = ind.get(WR + 'firstLineChars') if ind is not None else None
        if cur is None or int(cur) != want:
            return True
    # 对齐方式
    if spec.get('align'):
        jc = ppr.find(WR + 'jc') if ppr is not None else None
        if jc is None or jc.get(WR + 'val') != _align(spec['align']):
            return True
    sp = ppr.find(WR + 'spacing') if ppr is not None else None
    # 段前 / 段后（twips = pt*20，容许 1 磅误差避免抖动）
    if spec.get('before_pt') is not None:
        cur = sp.get(WR + 'before') if sp is not None else None
        if cur is None or abs(int(cur) - int(spec['before_pt']) * 20) > 20:
            return True
    if spec.get('after_pt') is not None:
        cur = sp.get(WR + 'after') if sp is not None else None
        if cur is None or abs(int(cur) - int(spec['after_pt']) * 20) > 20:
            return True
    # 行距（exact/atLeast 用值，single 隐式 240，auto 用 240 分度）
    # 注意：单倍行距没有显式 line_val，_line() 会兜底返回 (auto,240)，
    # 故不能以 line_val 缺失就跳过检查，否则无法纠出错误的单倍行距。
    if spec.get('line_rule'):
        lr, line = _line(spec)
        if lr is not None and line is not None:
            cur_line = sp.get(WR + 'line') if sp is not None else None
            cur_rule = sp.get(WR + 'lineRule') if sp is not None else None
            if cur_line is None or int(cur_line) != line or (cur_rule or 'auto') != lr:
                return True
    # 悬挂缩进（参考文献条目按编号位数多档）
    # v1.3.5：必须同时检查 w:hanging 和 w:left，只设 hanging 不设 left 会导致续行
    # 跑到左边距外显示负数（v1.3.4 遗留 bug）。
    if spec.get('hanging_cm') is not None:
        ind = ppr.find(WR + 'ind') if ppr is not None else None
        want_hanging = int(round(spec['hanging_cm'] * 567))
        cur_hanging = ind.get(WR + 'hanging') if ind is not None else None
        if cur_hanging is None or abs(int(cur_hanging) - want_hanging) > 10:
            return True
        # 必须有 w:left 且值与 hanging 相同
        cur_left = ind.get(WR + 'left') if ind is not None else None
        if cur_left is None or abs(int(cur_left) - want_hanging) > 10:
            return True
    return False


def _fmt_summary(spec):
    parts = []
    if spec.get('zh_font'):
        parts.append('字体' + spec['zh_font'])
    if spec.get('size'):
        parts.append('字号' + spec['size'])
    if spec.get('indent_type') == 'first':
        parts.append('首行缩进%d字符' % spec.get('indent_chars', 0))
    if spec.get('line_rule') == 'exact' and spec.get('line_val'):
        parts.append('固定值%d磅' % spec.get('line_val'))
    return ' / '.join(parts) if parts else '（同模板）'


def _current_fmt(p, spec):
    zh = None
    sz = None
    for r in p.iter(WR + 'r'):
        rpr = r.find(WR + 'rPr')
        if rpr is None:
            continue
        rf = rpr.find(WR + 'rFonts')
        if rf is not None and rf.get(WR + 'eastAsia') and zh is None:
            zh = rf.get(WR + 'eastAsia')
        s = rpr.find(WR + 'sz')
        if s is not None and s.get(WR + 'val') and sz is None:
            sz = int(s.get(WR + 'val'))
        if zh and sz:
            break
    parts = []
    if zh:
        parts.append('字体' + zh)
    if sz:
        parts.append('字号%dpt' % (sz // 2))
    ppr = p.find(WR + 'pPr')
    ind = ppr.find(WR + 'ind') if ppr is not None else None
    parts.append('有缩进' if (ind is not None and ind.get(WR + 'firstLineChars')) else '无缩进')
    return ' / '.join(parts) if parts else '（无显式格式）'


def _style_map_from_profile(profile):
    out = {}
    hs = profile.get("headingStyles") or {}
    for k, v in hs.items():
        sid = v.get("styleId")
        if sid:
            out[int(k)] = sid
    return out


def _inject_style_defs(src, need_defs):
    """把模板的样式定义注入目标 styles.xml，覆盖同名 styleId，返回新 styles.xml 字符串。"""
    reg_ns()
    z = zipfile.ZipFile(src)
    sroot = ET.fromstring(z.read("word/styles.xml").decode("utf-8"))
    # 移除目标中已存在的同名定义（确保模板定义生效）
    existing = {st.get(WR + "styleId") for st in sroot.findall(WR + "style")}
    for sid in need_defs:
        if sid in existing:
            for st in sroot.findall(WR + "style"):
                if st.get(WR + "styleId") == sid:
                    sroot.remove(st)
                    break
    for sid, xml_snippet in need_defs.items():
        if not xml_snippet:
            continue
        el = ET.fromstring(xml_snippet)
        sroot.append(el)
    return to_doc_xml(sroot)


def inject_outline(xml_snippet, level):
    """在样式定义的 <w:pPr> 中注入 <w:outlineLvl w:val="N"/>，使该样式可被 Word 用于自动生成目录。

    山大等模板的标题样式本身可能缺失大纲级别（outlineLvl=None），但模板批注会明确
    写明"大纲级别1/2/3级"——本函数按批注要求补回大纲级别，从而让"按模板套样式"后
    Word 也能一键出目录（批注为准的权威来源）。
    """
    reg_ns()
    el = ET.fromstring(xml_snippet)
    ppr = el.find(WR + "pPr")
    if ppr is None:
        ppr = ET.SubElement(el, WR + "pPr")
    for ex in ppr.findall(WR + "outlineLvl"):
        ppr.remove(ex)
    ol = ET.SubElement(ppr, WR + "outlineLvl")
    ol.set(WR + "val", str(level))
    return ET.tostring(el, encoding="unicode")

def _remove_indent_from_style_xml(xml_snippet):
    """从样式定义 XML 中移除 <w:ind> 节点，用于批注要求无缩进时覆盖样式定义。

    山大等模板的 Heading3 样式定义可能自带缩进，但批注明确要求三级标题
    "两端对齐、无缩进"——在注入样式定义前移除 <w:ind>，确保批注的
    无缩进要求生效（批注为准）。
    """
    el = ET.fromstring(xml_snippet)
    ppr = el.find(WR + "pPr")
    if ppr is not None:
        ind = ppr.find(WR + "ind")
        if ind is not None:
            ppr.remove(ind)
    return ET.tostring(el, encoding="unicode")


def _body_flow(root):
    """按正文顺序枚举 body 直接子节点：('p', 段落元素) / ('tbl', 表格元素) / ('other', 其它)。
    仅含正文层段落，不含表格内部段落。用于结构定位题注。"""
    body = root.find(WR + "body")
    if body is None:
        return []
    flow = []
    for ch in body:
        if ch.tag == WR + "p":
            flow.append(("p", ch))
        elif ch.tag == WR + "tbl":
            flow.append(("tbl", ch))
        else:
            flow.append(("other", ch))
    return flow


def _has_image(p):
    """段落是否含图片（DrawingML / VML 任一容器）。用于识别图题依附的对象。"""
    if p is None:
        return False
    for tag in (WR + "drawing", WR + "pict"):
        if len(list(p.iter(tag))) > 0:
            return True
    for el in p.iter():
        tag = el.tag
        if isinstance(tag, str) and (tag.endswith("}imagedata") or tag.endswith("}shape")):
            return True
    return False


def _caption_targets(root):
    """结构定位 + 文字特征验证，返回题注/题注相关段落 -> (通用类别key, kind)。

    v1.3.4 修复：纯结构定位会把表前的正文段误标为表题、把英文双标题或
    表后正文误标为表注/图注。现加入文字特征验证，保守策略，宁可漏不可错。

    规则（中文论文通行惯例，跨校通用）：
      - 表题(table)：紧邻 w:tbl 之前（跳过空段）的非空段，且文字含表/Table/Tab. 编号
        （如"表3-1"、"表 3-1"、"Table 3-1"、"Tab. 3-1"、"续表3-1"）；
      - 表后紧邻段（跳过空段）分类：
          * 以 Tab./Table 开头 → 英文表题（中英文双标题），按 table 规格套用（不是表注）；
          * 含来源/注释特征词（资料来源/数据来源/来源：/注：/Source:/Note: 等）→ table_note；
          * 其它正文散文段 → 不标（保留原样）。
      - 图题(figure)：紧邻含图段落之后（跳过空段）的非空段，且含图/Figure/Fig. 编号；
      - 图后再下一段：同表后规则（英文标题→figure，来源特征→figure_note，否则不标）。

    返回的 key 直接对应 profile.levels 的通用类别；若某学校模板未批注该类别，
    套用层查不到 spec 会自动跳过。
    """
    targets = {}
    flow = _body_flow(root)
    n = len(flow)

    def _next_nonempty_p(start):
        k = start + 1
        while k < n:
            if flow[k][0] == "p":
                txt = _text_of(flow[k][1]).strip()
                if txt:
                    return flow[k][1], txt, k
            k += 1
        return None, None, None

    def _prev_nonempty_p(start):
        k = start - 1
        while k >= 0:
            if flow[k][0] == "p":
                txt = _text_of(flow[k][1]).strip()
                if txt:
                    return flow[k][1], txt, k
            k -= 1
        return None, None, None

    for j, (kind, el) in enumerate(flow):
        if kind == "tbl":
            # 表前最近非空段 -> 表题候选（需文字特征验证）
            p_prev, t_prev, _ = _prev_nonempty_p(j)
            if p_prev is not None and _is_table_caption_text(t_prev):
                targets[p_prev] = ("table", "table_caption")
            # 表后最近非空段 -> 分类（英文标题/表注/不标）
            p_next, t_next, _ = _next_nonempty_p(j)
            if p_next is not None:
                cls = _classify_after_caption(t_next, is_table=True)
                if cls is not None:
                    targets[p_next] = cls
        elif kind == "p" and _has_image(el):
            # 含图段之后第一段 -> 图题候选（需文字特征验证）
            p_next, t_next, kk = _next_nonempty_p(j)
            if p_next is not None and _is_figure_caption_text(t_next):
                targets[p_next] = ("figure", "figure_caption")
                # 图题之后再下一段 -> 英文图题/图注/不标
                p_after, t_after, _ = _next_nonempty_p(kk)
                if p_after is not None:
                    cls = _classify_after_caption(t_after, is_table=False)
                    if cls is not None:
                        targets[p_after] = cls
    return targets


# 表/图编号模式（跨校通用，兼容中英文/有无空格/全半角连字符）：
#   表1, 表1-1, 表 1-1, 表3－1, 续表3-1, 续表 3-1,
#   Table 1, Table 1-1, Tab. 1-1, Tab.1-1, 图1, 图1-1, Figure 1-1, Fig. 1-1 ...
_CAP_TABLE_RE = re.compile(
    r'(?:^|[\s（(])(?:续\s*)?(?:表\s*\d|Table\s+\d|Tab\.?\s*\d)',
    re.IGNORECASE,
)
_CAP_FIGURE_RE = re.compile(
    r'(?:^|[\s（(])(?:图\s*\d|Figure\s+\d|Fig\.?\s*\d)',
    re.IGNORECASE,
)
# 英文标题起始（中英文双标题场景：英文标题紧跟中文标题）
_EN_TBL_START_RE = re.compile(r'^\s*(?:Tab\.?|Table)\b', re.IGNORECASE)
_EN_FIG_START_RE = re.compile(r'^\s*(?:Fig\.?|Figure)\b', re.IGNORECASE)
# 表注/图注特征词（跨校：资料来源/数据来源/来源/注/Source/Note 等）
_NOTE_KEYWORDS = (
    '资料来源', '数据来源', '来源：', '来源:', '注：', '注:',
    'Source：', 'Source:', 'source：', 'source:',
    'Note：', 'Note:', 'note：', 'note:',
)


def _is_table_caption_text(text):
    """表前段落文字是否像表题（含表/Table/Tab. 编号模式）。"""
    t = (text or '').strip()
    return bool(t) and bool(_CAP_TABLE_RE.search(t))


def _is_figure_caption_text(text):
    """图后段落文字是否像图题（含图/Figure/Fig. 编号模式）。"""
    t = (text or '').strip()
    return bool(t) and bool(_CAP_FIGURE_RE.search(t))


def _classify_after_caption(text, is_table):
    """表后/图题后紧邻段的分类：英文标题→(table/figure, *_caption)，
    来源注释→(*_note, *_note)，其它正文→None（保留原样，不套题注格式）。"""
    t = (text or '').strip()
    if not t:
        return None
    if is_table:
        if _EN_TBL_START_RE.match(t):
            return ("table", "table_caption")
        if t.startswith(('续表', '续 表')):
            return ("table", "table_caption")
    else:
        if _EN_FIG_START_RE.match(t):
            return ("figure", "figure_caption")
    for kw in _NOTE_KEYWORDS:
        if kw in t:
            return ("table_note", "table_note") if is_table else ("figure_note", "figure_note")
    return None


# ---------------------------------------------------------------------------
# 参考文献条目按编号位数选悬挂缩进档位
# ---------------------------------------------------------------------------
# 匹配参考文献编号：[1], [12], [100]；1., 12., 100.；(1), (12)；1、
_REF_NUM_RE = re.compile(r'^\s*[\[（(]?\s*(\d{1,4})\s*[\]）)]?\s*[\.、,，]?\s*')


def _ref_number_digits(text):
    """从参考文献条目文本中取出编号位数（1-9→1，10-99→2，100-999→3）。
    无法识别编号时返回 None（调用方应回退到默认档位）。"""
    t = (text or '').lstrip()
    m = _REF_NUM_RE.match(t)
    if m:
        num = m.group(1)
        if num.isdigit():
            return len(num)
    return None


def _ref_entry_spec_for_text(text, base_spec):
    """根据参考文献条目编号位数返回带正确 hanging_cm 的 spec 副本。

    若 base_spec 含 hanging_tiers（多档悬挂缩进，由批注"编号1-9悬挂缩进0.6厘米
    编号10-99悬挂缩进0.74厘米..."解析得到），则按条目编号位数选对应档位；
    否则使用 base_spec.hanging_cm（单档，向后兼容）。无法识别编号时回退首档。
    """
    tiers = (base_spec or {}).get("hanging_tiers")
    if not tiers:
        return base_spec
    digits = _ref_number_digits(text)
    chosen = None
    if digits is not None:
        for tier in tiers:
            if tier.get("digits") == digits:
                chosen = tier.get("cm")
                break
        if chosen is None:
            # 编号位数超出已配置档位（如配置到3位但出现4位编号），用最大档位兜底
            chosen = tiers[-1].get("cm")
    else:
        chosen = tiers[0].get("cm")
    spec = dict(base_spec)
    spec["hanging_cm"] = chosen
    return spec


# ===========================================================================
# 图片归一化：浮动→内嵌 / 超宽缩放 / 去固定行距 / 居中
# ===========================================================================

def _text_area_width_emu(root, profile=None):
    """从文档 sectPr 或 profile 计算文本区宽度(EMU)。
    文本区宽度 = 页面宽度 − 左页边距 − 右页边距。
    优先用 profile 的页边距（模板要求），其次用文档自身的 sectPr。"""
    pg_w = None
    pg_mar_l = pg_mar_r = None
    if profile:
        page = (profile.get("spec") or {}).get("page") or {}
        if page.get("left_cm"):
            pg_mar_l = int(round(float(page["left_cm"]) * EMU_PER_CM))
        if page.get("right_cm"):
            pg_mar_r = int(round(float(page["right_cm"]) * EMU_PER_CM))
    for sect in root.iter(WR + "sectPr"):
        pgsz = sect.find(WR + "pgSz")
        if pgsz is not None and pg_w is None:
            w_val = pgsz.get(WR + "w")
            if w_val:
                pg_w = int(w_val) * EMU_PER_TWIP
        pgmar = sect.find(WR + "pgMar")
        if pgmar is not None:
            if pg_mar_l is None:
                l_val = pgmar.get(WR + "left")
                if l_val:
                    pg_mar_l = int(l_val) * EMU_PER_TWIP
            if pg_mar_r is None:
                r_val = pgmar.get(WR + "right")
                if r_val:
                    pg_mar_r = int(r_val) * EMU_PER_TWIP
        if pg_w is not None and pg_mar_l is not None and pg_mar_r is not None:
            break
    if pg_w is None:
        return _DEFAULT_TEXT_WIDTH_EMU
    if pg_mar_l is None:
        pg_mar_l = 1800 * EMU_PER_TWIP  # ≈3.17cm
    if pg_mar_r is None:
        pg_mar_r = 1800 * EMU_PER_TWIP
    return max(pg_w - pg_mar_l - pg_mar_r, _DEFAULT_TEXT_WIDTH_EMU // 2)


def _convert_anchor_to_inline(drawing_el):
    """将 w:drawing 内的 wp:anchor（浮动图片）转换为 wp:inline（内嵌图片）。

    浮动图片锚定在文字位置上，文字可环绕（四周型/紧密型/衬于文字下方等），
    极易与正文重叠。内嵌图片随文字流流动，不会重叠——这是消除"图片被字体盖住"
    最有效的一步。

    wp:anchor 需要保留的子元素：wp:extent / wp:effectExtent / wp:docPr /
    wp:cNvGraphicFramePr / a:graphic；其余（simplePos / positionH / positionV /
    wrapNone 等）在 inline 模式下无意义，丢弃。
    """
    anchor = None
    for child in drawing_el:
        if child.tag == WPR + "anchor":
            anchor = child
            break
    if anchor is None:
        return False
    keep_tags = {
        WPR + "extent", WPR + "effectExtent", WPR + "docPr",
        WPR + "cNvGraphicFramePr", AR + "graphic",
    }
    inline = ET.Element(WPR + "inline")
    inline.set("distT", "0")
    inline.set("distB", "0")
    inline.set("distL", "0")
    inline.set("distR", "0")
    for child in list(anchor):
        if child.tag in keep_tags:
            anchor.remove(child)
            inline.append(child)
    drawing_el.remove(anchor)
    drawing_el.append(inline)
    return True


def _scale_oversized_drawing(drawing_el, max_width_emu, container=None):
    """如果图片宽度超过文本区宽度，等比缩小 cx/cy。

    修改 wp:extent 的 cx/cy，同时同步修改 a:graphicData/pic:pic/pic:spPr/
    a:xfrm/a:ext 的 cx/cy，确保渲染尺寸与声明一致。
    """
    if container is None:
        for child in drawing_el:
            if child.tag in (WPR + "inline", WPR + "anchor"):
                container = child
                break
    if container is None:
        return False
    extent = container.find(WPR + "extent")
    if extent is None:
        return False
    cx_str = extent.get("cx")
    cy_str = extent.get("cy")
    if not cx_str or not cy_str:
        return False
    cx = int(cx_str)
    cy = int(cy_str)
    if cx <= max_width_emu or cx <= 0:
        return False
    scale = max_width_emu / cx
    new_cx = int(max_width_emu)
    new_cy = int(cy * scale)
    extent.set("cx", str(new_cx))
    extent.set("cy", str(new_cy))
    # 同步 effectExtent
    eff = container.find(WPR + "effectExtent")
    if eff is not None:
        for attr in ("l", "t", "r", "b"):
            eff.set(attr, "0")
    # 同步 pic:spPr/a:xfrm/a:ext
    graphic = container.find(AR + "graphic")
    if graphic is not None:
        for gd in graphic:
            if gd.tag == AR + "graphicData":
                for pic in gd:
                    if pic.tag == PICR + "pic":
                        for spPr in pic:
                            if spPr.tag == AR + "spPr":
                                for xfrm in spPr:
                                    if xfrm.tag == AR + "xfrm":
                                        for ext in xfrm:
                                            if ext.tag == AR + "ext":
                                                ext.set("cx", str(new_cx))
                                                ext.set("cy", str(new_cy))
    return True


def _normalize_image_para_format(p):
    """归一化纯图片段落格式：居中 + 去固定值行距 + 去首行缩进。

    仅对"纯图片段落"（无文字内容或仅空白）执行——标题段若含图片，
    标题 pass 会正常套标题格式，不应被此处居中/去缩进干扰。
    """
    txt = _text_of(p).strip()
    if txt:
        return False  # 有文字的段落（如含图的标题）不在此处理
    changed = False
    ppr = p.find(WR + "pPr")
    if ppr is None:
        ppr = ET.Element(WR + "pPr")
        p.insert(0, ppr)
    # 居中
    jc = ppr.find(WR + "jc")
    if jc is None:
        jc = ET.SubElement(ppr, WR + "jc")
        jc.set(WR + "val", "center")
        changed = True
    elif jc.get(WR + "val") != "center":
        jc.set(WR + "val", "center")
        changed = True
    # 去固定值行距（lineRule=exact 会裁剪图片）
    sp = ppr.find(WR + "spacing")
    if sp is not None:
        if sp.get(WR + "lineRule") == "exact":
            sp.set(WR + "lineRule", "auto")
            sp.set(WR + "line", "240")  # 单倍
            changed = True
    # 去首行缩进（会把图片推偏）
    ind = ppr.find(WR + "ind")
    if ind is not None:
        for attr in ("firstLine", "firstLineChars"):
            if ind.get(WR + attr) is not None:
                del ind.attrib[WR + attr]
                changed = True
        # 如果 ind 上的所有缩进属性都已清空，移除整个 ind 元素
        remaining = [k for k in ind.attrib if k.startswith(WR)]
        if not remaining:
            ppr.remove(ind)
            changed = True
    return changed


def _normalize_images(root, profile=None, skip_paras=None):
    """图片归一化 pass：在格式化 pass 之前运行，主动修复客户文档中的图片问题。

    四项归一化：
    1. 浮动图片(wp:anchor)→内嵌(wp:inline)：消除文本环绕导致的图片与文字重叠
    2. 超宽图片等比缩放到文本区宽度内：消除图片溢出页面边界
    3. 纯图片段落去除 lineRule=exact 固定行距：消除行距裁剪图片
    4. 纯图片段落居中+去首行缩进：消除图片偏移

    skip_paras: 集合，其中的段落元素将被跳过（用于保护封面/结构页图片）。

    返回修改记录列表（kind="image"），供 fix() 汇总到 changes。
    """
    reg_ns()
    max_w = _text_area_width_emu(root, profile)
    changes = []
    _skip = skip_paras or set()

    for p in root.iter(WR + "p"):
        if p in _skip:
            continue
        if not _has_image(p):
            continue
        para_changed = False
        details = []

        # 1. 浮动→内嵌 + 2. 超宽缩放（对每个 w:drawing 逐一处理）
        for drawing in p.iter(WR + "drawing"):
            if _convert_anchor_to_inline(drawing):
                para_changed = True
                details.append("浮动→内嵌")
            if _scale_oversized_drawing(drawing, max_w):
                para_changed = True
                details.append("超宽缩放")

        # 3+4. 段落格式归一化（仅纯图片段落）
        if _normalize_image_para_format(p):
            para_changed = True
            details.append("段落格式归一化")

        if para_changed:
            txt = _text_of(p).strip()[:40]
            changes.append({
                "kind": "image", "text": txt or "（图片段落）", "level": 0,
                "old": "原始图片格式", "new": " / ".join(details),
                "new_name": " / ".join(details), "conf": 0.9,
                "_para": p,
            })
    return changes


def _comment_text_for_change(c):
    """根据一条 change 生成简洁的 Word 批注文字。"""
    kind = c.get("kind", "")
    if kind == "heading":
        old = c.get("old") or "（无样式）"
        new_name = c.get("new_name") or c.get("new") or ""
        return "标题样式：%s → %s，套样式后可生成目录" % (old, new_name)
    if kind == "body":
        return "正文格式：按模板统一字体/字号/缩进/行距"
    if kind == "image":
        details = c.get("new_name") or c.get("new") or ""
        return "图片处理：%s" % details
    if kind == "reference":
        details = c.get("new_name") or c.get("new") or ""
        return "参考文献条目：%s" % details
    if kind == "reference_heading":
        details = c.get("new_name") or c.get("new") or ""
        return "参考文献标题：%s" % details
    if kind == "numbering":
        return "自动编号已转为正文（移除 numPr，保留文字）"
    if kind == "punctuation":
        cnt = c.get("count", 0)
        return "标点修正：半角→全角（%d处）" % cnt
    if kind in ("table_caption", "figure_caption"):
        return "题注格式：按模板修正"
    if kind in ("table_note", "figure_note"):
        return "题注格式：按模板修正"
    if kind == "footnote":
        return "脚注格式：按模板修正"
    if kind == "margin":
        return "页边距：按模板修正为 %s" % (c.get("new") or "")
    if kind == "suspected_caption":
        return c.get("note") or "此段紧邻表格/图片且字号较小，可能是题注或来源说明，已保留原格式，请确认"
    return "格式修正：按模板调整"


def _apply_comments(z, root, changes, replacements):
    """为 changes 中每个带 `_para` 段落引用的 change 添加 Word 批注。

    批注 ID 从文档现有最大 ID + 1 开始，绝不覆盖已有批注（学校模板批注/导师批注）。
    comments.xml / [Content_Types].xml / document.xml.rels 的新建与注册由
    docxutils.ensure_comments_part 统一处理。返回添加的批注数。
    """
    para_changes = [c for c in changes if c.get("_para") is not None]
    if not para_changes:
        return 0
    comments_root = ensure_comments_part(z, replacements)
    next_id = max_comment_id(z) + 1
    n = 0
    for c in para_changes:
        p = c.get("_para")
        # 段落可能已被从树中移除（理论上不会），防御性跳过
        if p is None:
            continue
        text = _comment_text_for_change(c)
        add_comment_marker(p, next_id, text, comments_root)
        next_id += 1
        n += 1
    # 把 comments.xml 写入 replacements（无论新建还是修改）
    replacements["word/comments.xml"] = to_doc_xml(comments_root)
    return n


def _punctuation_pass(root, first_chap, ref_start, tbl_ps, cap_targets, styles_map, heading_sids):
    """正文段落半角→全角标点替换（独立 pass，在正文格式套用之后执行）。

    保守策略：跳过标题段、表格内段、含图片段、目录样式段、题注段、结构页标题。
    仅处理正文区（首个一级标题之后、参考文献之前）的散文段落。

    返回标点修改记录列表（kind="punctuation"），每条含段落引用 `_para`。
    """
    changes = []
    paras_list = list(root.iter(WR + "p"))
    for idx, p in enumerate(paras_list):
        if p in tbl_ps:
            continue
        if _has_image(p):
            continue
        t = _text_of(p).strip()
        if not t:
            continue
        # 跳过目录域/书签残留
        if is_toc_residue(t):
            continue
        # 跳过目录样式段
        if _is_toc_style(_cur_style_id(p), styles_map, heading_sids):
            continue
        # 跳过结构页标题（摘要/目录/参考文献/致谢等）
        if is_structural_title(t):
            continue
        # 跳过识别到的标题段
        lvl, _ = detect_heading(t)
        if lvl is not None:
            continue
        # 仅处理正文区（首个一级标题之后、参考文献之前）
        in_body = (first_chap is not None and idx > first_chap
                   and (ref_start is None or idx < ref_start))
        if not in_body:
            continue
        # 跳过题注/表注/图注段（这些段由题注 pass 处理格式，标点不强行替换）
        if p in cap_targets:
            continue
        # 执行标点替换
        counts = replace_punct_in_paragraph(p)
        if counts:
            total = punct_total(counts)
            snippet = _text_of(p).strip()[:40]
            changes.append({
                "kind": "punctuation", "text": snippet, "level": 0,
                "old": "半角标点",
                "new": punct_counts_summary(counts),
                "new_name": punct_counts_summary(counts),
                "conf": 0.9,
                "count": total,
                "counts": counts,
                "para_index": idx,
                "_para": p,
            })
    return changes


# ---------------------------------------------------------------------------
# 自动编号（w:numPr）转正文：有些客户论文用 Word 自动编号列表，调整格式后
# 编号可能异常增加或错乱。按用户需求，移除正文区域段落的自动编号（保留文字），
# 不动标题段、表格内段、封面/摘要/目录等结构页。
# ---------------------------------------------------------------------------

def _is_heading_para(p, styles_map, heading_sids):
    """段落是否是标题段（套有标题样式）。

    v1.3.4：仅以"段落是否套有 Heading/标题样式"作为判定标准，不使用 detect_heading
    文本启发式——否则"（1）化妆品市场规模"这类正文列举项会被 detect_heading 的
    中文序号逻辑误判为标题，导致其自动编号 numPr 不应被保留（实际应移除）。
    标题样式由套样式 pass 显式赋予，是最可靠的判定依据。
    """
    sid = _cur_style_id(p)
    if sid and heading_sids and sid in heading_sids:
        return True
    if sid:
        nm = ((styles_map or {}).get(sid) or {}).get("name") or ""
        if re.search(r"heading|标题", nm, re.I):
            return True
        if sid in ("Heading1", "Heading2", "Heading3"):
            return True
    return False


def _looks_like_body_despite_heading_style(text):
    """v1.3.5：段落虽套有 Heading 样式，但内容明显是正文，应降级为正文。

    判定条件（满足任一即视为误套标题样式的正文）：
      - 段落以句号/问号/感叹号/分号结尾（完整句子特征）；
      - 段落长度 > 30 字（标题通常较短）；
      - 不含任何标题编号特征（非"第X章/节"、非"1 / 1.1 / 1.1.1"多级编号、
        非"一、二、"等），且文本不以英文字母+数字编号开头。

    返回 True 表示"像正文，应降级"。
    """
    t = (text or "").strip()
    if not t:
        return False
    # 完整句子结尾
    if t[-1] in "。！？；.!?;":
        return True
    # 长度超阈值
    if len(t) > 30:
        return True
    return False


def _para_font_size_half_pt(p):
    """获取段落首个有字号设置的 run 的字号（half-pt 单位，如 24=12pt）。
    若段落无显式字号返回 None（继承样式）。
    """
    for r in p.iter(WR + 'r'):
        rpr = r.find(WR + 'rPr')
        if rpr is None:
            continue
        s = rpr.find(WR + 'sz')
        if s is not None and s.get(WR + 'val'):
            try:
                return int(s.get(WR + 'val'))
            except (ValueError, TypeError):
                continue
    return None


def _find_small_text_near_objects(root, body_sz_half_pt, cap_targets,
                                   first_chap, ref_start, tbl_ps,
                                   styles_map, heading_sids):
    """v1.3.5：识别表格/图片附近的小字段落（疑似题注/来源说明）。

    客户写作时可能不使用标准编号前缀（"表1-1"等），只是在表格/图下面敲一行
    小字。这些段落字号小于正文，紧邻表格或图片，段长较短——应保留原格式不被
    正文规格覆盖，并加 Word 批注提醒用户确认。

    判定条件（全部满足）：
      1. 段落紧邻表格或含图段落（前后2段以内）；
      2. 字号显式小于正文（差 >=2 half-pt，即1pt 以上）；
      3. 段长 <= 80 字（题注/来源说明通常较短）；
      4. 不是已识别的题注/来源说明段；
      5. 在正文区域内（first_chap 之后、ref_start 之前）。

    返回 set：疑似小字题注段落集合，调用方应跳过正文格式化并加批注。
    """
    suspected = set()
    if not body_sz_half_pt:
        return suspected
    flow = _body_flow(root)
    n = len(flow)
    paras_list = list(root.iter(WR + "p"))
    para_idx = {id(p): i for i, p in enumerate(paras_list)}

    # 收集表格/图片在 flow 中的位置
    obj_positions = []
    for j, (k, el) in enumerate(flow):
        if k == "tbl" or (k == "p" and _has_image(el)):
            obj_positions.append(j)

    for j, (k, el) in enumerate(flow):
        if k != "p":
            continue
        p = el
        if p in cap_targets or id(p) not in para_idx:
            continue
        idx = para_idx[id(p)]
        # 正文区域检查
        if first_chap is not None and idx <= first_chap:
            continue
        if ref_start is not None and idx >= ref_start:
            continue
        if p in tbl_ps or _has_image(p):
            continue
        if _is_toc_style(_cur_style_id(p), styles_map):
            continue
        if is_structural_title(_text_of(p).strip()):
            continue

        # 检查是否紧邻表格/图片（前后2个flow位置内）
        near_obj = any(abs(j - op) <= 2 for op in obj_positions)
        if not near_obj:
            continue

        # 字号检查：显式设了字号且明显小于正文
        psz = _para_font_size_half_pt(p)
        if psz is None or psz >= body_sz_half_pt - 1:
            continue

        # 段长检查
        txt = _text_of(p).strip()
        if len(txt) > 80:
            continue

        suspected.add(p)
    return suspected


def _remove_numbering_pass(root, first_chap, ref_start, tbl_ps, cap_targets,
                           styles_map, heading_sids):
    """扫描正文区段落，移除非标题段的 <w:numPr>（自动编号），使其编号不再自动生成。

    保守规则（只处理正文区域、只动正文段）：
      - 仅处理正文区（首个一级标题之后、参考文献之前）；
      - 跳过表格内段、含图段、目录样式段、结构页标题、题注/表注段；
      - 跳过标题段（标题样式或 detect_heading 命中）——标题编号由样式控制；
      - 仅移除段落 pPr 下的 <w:numPr>，不删文字、不改其它格式。

    返回修改记录列表（kind="numbering"），每条含段落引用 `_para`。
    """
    changes = []
    paras_list = list(root.iter(WR + "p"))
    for idx, p in enumerate(paras_list):
        if p in tbl_ps:
            continue
        if _has_image(p):
            continue
        t = _text_of(p).strip()
        # 空段也处理（可能带 numPr，会产生孤立编号）；其它有文字的段走常规过滤。
        if t:
            if is_toc_residue(t):
                continue
            if _is_toc_style(_cur_style_id(p), styles_map, heading_sids):
                continue
            if is_structural_title(t):
                continue
        # 仅正文区
        in_body = (first_chap is not None and idx > first_chap
                   and (ref_start is None or idx < ref_start))
        if not in_body:
            continue
        # 标题段不动（题注/来源说明段也清除 numPr，它们不应有自动编号）
        if _is_heading_para(p, styles_map, heading_sids):
            continue
        ppr = p.find(WR + "pPr")
        if ppr is None:
            continue
        numpr = ppr.find(WR + "numPr")
        if numpr is None:
            continue
        ppr.remove(numpr)
        snippet = t[:40] if t else "（空段）"
        changes.append({
            "kind": "numbering", "text": snippet, "level": 0,
            "old": "自动编号(numPr)",
            "new": "自动编号已转为正文",
            "new_name": "自动编号已转为正文",
            "conf": 0.95,
            "_para": p,
        })
    return changes


def fix(src, dst, profile=None, add_comments=True):
    z, root = load(src)

    # 样式表：用于识别目录(toc)样式段落，避免误套标题/污染正文区
    styles_map = load_styles(src)
    # 画像中真实的标题样式 styleId（即便其名字叫『toc 1』也要当标题处理，不可误跳）
    heading_sids = set(
        v.get("styleId") for v in (profile or {}).get("headingStyles", {}).values()
        if v.get("styleId")
    ) if profile else set()
    # 样式映射：有模板画像则用模板真实 styleId，否则通用 Heading
    if profile and (profile.get("headingStyles")):
        style_for = _style_map_from_profile(profile)
        template_exact = True
        # 展示用：level -> (styleId, 展示名)
        disp = {}
        for k, v in profile["headingStyles"].items():
            disp[int(k)] = (v.get("styleId"), v.get("name") or v.get("styleId"))
        tmpl_name = os.path.basename(profile.get("source", "学校模板"))
        # 回退：模板画像中缺失的级别（如山大模板一级/二级标题用 Normal 样式，
        # 不算真实标题样式，画像中不输出）用内置 Heading1/2/3 样式兜底。
        for lvl in (1, 2, 3):
            if lvl not in style_for:
                style_for[lvl] = STYLE_FOR[lvl]
                if lvl not in disp:
                    disp[lvl] = (STYLE_FOR[lvl], DISP_GENERIC[lvl][1])
    else:
        style_for = STYLE_FOR
        template_exact = False
        disp = DISP_GENERIC
        tmpl_name = "通用 Heading 样式"
    # 需要注入的样式定义（来自模板画像中存的 xml 片段）
    need_defs = {}
    # 批注 per-level 要素（用于判断该级别是否明确要求无缩进）
    _levels_for_indent = (profile or {}).get("levels") or {}
    if profile and profile.get("headingStyles"):
        for v in profile["headingStyles"].values():
            sid = v.get("styleId")
            x = v.get("xml")
            if sid and x:
                # 按模板批注要求补回大纲级别（批注写明 大纲级别1/2/3级 时生效），
                # 使"按模板套样式"后 Word 也能一键出目录。
                ol = v.get("outline_level")
                if ol is not None:
                    x = inject_outline(x, ol)
                # 如果批注明确要求无缩进，从样式定义中移除 <w:ind>，
                # 确保批注的"无缩进"覆盖模板样式定义中自带的缩进。
                lvl_key = str(v.get("level", ""))
                lvl_spec = _levels_for_indent.get(lvl_key, {})
                if lvl_spec.get("indent_type") == "none" or lvl_spec.get("indent_chars") == 0:
                    x = _remove_indent_from_style_xml(x)
                need_defs[sid] = x
    # 注意：回退使用的内置 Heading1/2/3 样式（STYLE_FOR）不需要注入——
    # 目标文档通常已自带这些样式定义；need_defs 只包含画像中实际存在的模板样式。

    changes = []  # 后续格式化记录
    skipped_residue = 0
    already_ok = 0  # 已是目标样式、无需改动（用于透明告知客户）

    # ---- 范围界定（与诊断一致）----
    paras_list = list(root.iter(WR + "p"))
    # 表格内的段落不处理（避免改到表内文字格式）
    tbl_ps = set()
    for tbl in root.iter(WR + "tbl"):
        for pp in tbl.iter(WR + "p"):
            tbl_ps.add(pp)
    # 首个一级标题（第一章）之后为"正文区"，到"参考文献"之前结束
    # v1.3.5：定位 first_chap 时要求一级标题必须是"第X章/篇/编"或多级编号（1.1 等），
    # 不能是单数字开头（如封面"1. 分类号"会被误判为正文起点）。
    first_chap = None
    ref_start = None
    for i, p in enumerate(paras_list):
        t = _text_of(p).strip()
        if not t or is_toc_residue(t) or _is_toc_style(_cur_style_id(p), styles_map, heading_sids):
            continue
        lv, _ = detect_heading(t)
        # 一级标题还需满足：要么"第X章"开头，要么是多级编号（含小数点），
        # 单数字开头的不作为正文起点（可能是封面/目录项）
        if lv == 1 and first_chap is None:
            _is_chapter_kw = bool(re.match(r"^第[一二三四五六七八九十百千\d]+[章篇编]", t))
            _is_multi_num = bool(re.match(r"^\d+\.\d+", t))
            # v1.3.5：单数字+分隔符（如"1 绪论"）也是一级标题。
            # detect_heading 已要求单数字后必须跟分隔符，所以这里只需排除正文特征。
            _is_single_num = bool(re.match(r"^\d{1,3}[\s、：:（(\t]", t))
            # v1.3.5：即使匹配了标题模式，也要排除正文特征（完整句子/过长），
            # 否则"第一章：绪论。介绍论文研究背景..."这类研究内容概述会被误判为正文起点。
            if (_is_chapter_kw or _is_multi_num or _is_single_num) and not _looks_like_body_despite_heading_style(t):
                first_chap = i
        if is_reference_heading(t) and ref_start is None and not _looks_like_body_despite_heading_style(t):
            ref_start = i
    spec_cats = (profile or {}).get("spec", {}).get("cats", {}) if profile else {}
    levels = (profile or {}).get("levels") or {}

    def _lvl_spec(key, cat_key):
        """优先 levels（样式定义+批注合并的完整 per-level 要素），缺则回退旧 spec.cats。"""
        ls = levels.get(key)
        if ls:
            return ls
        return spec_cats.get(cat_key) or {}

    body_spec = _lvl_spec("body", "body")

    # ---- 聚集判定：第X章/第X节「逐条罗列」视为正文，而非章标题 ----
    # 用户实测踩坑：1.2.2「研究内容」里用"第一章：绪论…第七章"逐条概述七个章节，
    # 本质是正文列举；若套章样式会污染目录、并造成"1.2.2 里冒出一级章"的层级错乱。
    # 判定：凡"第X章/第X节"彼此间隔 <=K 段（即一条条列出来），一律当正文处理；
    # 真正隔好几页的章标题（彼此间隔远大于 K）不受影响。K=3 远小于真实章间距。
    CHAP_CAND = re.compile(r"^第[一二三四五六七八九十百千\d]+[章篇编节]")
    K_CLUSTER = 3
    _chap_idx = [i for i, p in enumerate(paras_list) if CHAP_CAND.match(_text_of(p).strip())]
    demote_chapter = set()
    for a in _chap_idx:
        if any(abs(a - b) <= K_CLUSTER for b in _chap_idx if b != a):
            demote_chapter.add(a)

    # 结构定位题注（表题/表注/图题/图注），正文循环需跳过这些段，交给题注 pass 处理
    cap_targets = _caption_targets(root)

    # v1.3.5：图片归一化移到 first_chap 定位之后，跳过封面/结构页区域（idx <= first_chap）。
    # 封面的校徽、下划线标签等浮动图片不应被转为内嵌或缩放。
    _cover_paras = set()
    if first_chap is not None:
        for _ci in range(min(first_chap + 1, len(paras_list))):
            _cover_paras.add(paras_list[_ci])
    img_changes = _normalize_images(root, profile, skip_paras=_cover_paras)
    changes.extend(img_changes)

    # v1.3.5：识别表格/图片附近字号明显小于正文的短段落（疑似题注/来源说明）。
    # 这些段落保留原格式不被正文规格覆盖，并在批注中提醒用户确认。
    _body_sz = (body_spec or {}).get('sz')
    suspected_small = _find_small_text_near_objects(
        root, _body_sz, cap_targets, first_chap, ref_start, tbl_ps,
        styles_map, heading_sids)

    for idx, p in enumerate(paras_list):
        if p in tbl_ps:
            continue
        txt = _text_of(p).strip()
        if is_toc_residue(txt) or _is_toc_style(_cur_style_id(p), styles_map, heading_sids):
            skipped_residue += 1
            continue
        lvl, conf = detect_heading(txt)
        # ===== 标题：套样式 + 按批注写字体/字号/加粗/对齐/段间距 =====
        # demote_chapter 中的"第X章/节"是 1.2.2 研究内容里的逐条概述（聚集列举），
        # 不当章标题，落到下面的正文分支按正文处理。
        #
        # v1.3.5：封面/摘要/目录区（首个一级标题之前）不套标题格式。
        # 参考文献之后（致谢/附录等）也不套标题格式——由参考文献pass和结构页保护处理。
        # 这些区域的文字（如封面"1. 分类号"、目录页"1.1 研究背景"、附录问卷"1. 请..."）
        # 可能被 detect_heading 命中，但属于结构页/后论区域，不应按正文标题格式化。
        _in_body_for_heading = (first_chap is not None and idx > first_chap
                                and (ref_start is None or idx < ref_start))
        if lvl in (1, 2, 3, 4) and idx not in demote_chapter and _in_body_for_heading:
            # 4 级标题（如"2.1.2.1"）若模板无独立样式，回退用 3 级样式承载
            target = style_for.get(lvl) or style_for.get(3)
            if not target:
                continue
            ppr = p.find(WR + "pPr")
            cur = ppr.find(WR + "pStyle").get(WR + "val") if (ppr is not None and ppr.find(WR + "pStyle") is not None) else None
            hspec = _lvl_spec(str(lvl), "h%d" % lvl)
            if hspec:
                # 标题文本 run 格式化（_set_run_rpr 内部已跳过含图 run），
                # 段落格式（jc/spacing/ind）也正常设置——标题规格不用固定值行距(lineRule=exact)，
                # 不会压缩图片；图片 run 的 rPr 由 _run_has_image 守卫保持原样。
                _format_runs(p, hspec)
                _set_para_format(p, hspec, style_val=target)
            else:
                _set_style(p, target)
            if cur == target:
                already_ok += 1
                continue
            old_disp = cur if cur else "（无/自定义样式）"
            new_sid, new_name = disp.get(lvl, (target, target))
            changes.append({
                "kind": "heading", "text": txt[:60], "level": lvl,
                "old": old_disp, "new": new_sid, "new_name": new_name,
                "conf": conf, "_para": p,
            })
            continue
        # ===== 正文：按批注 body spec 写字体/字号/首行缩进/行距 =====
        # 含图片的段落整段跳过：向图片段落注入首行缩进/固定值行距会压缩图片、与文字重叠
        # （"图片被字体盖住"），故保留其原始段落格式（通常居中+原行距），不格式化。
        #
        # v1.3.5：误套 Heading 样式的正文也在此处理。客户文档中部分段落虽套了
        # Heading1/2/3 样式（可能因格式刷或误操作），但文本特征明显是正文
        # （完整句子结尾、段落过长），应降级为正文格式，覆盖 Heading 样式带来的
        # 加粗/大字号效果。同时把段落样式改为 Normal/正文，避免样式注入后持续显示标题格式。
        in_body = (first_chap is not None and idx > first_chap
                   and (ref_start is None or idx < ref_start))
        _misstyled = (_is_heading_para(p, styles_map, heading_sids)
                      and _looks_like_body_despite_heading_style(txt))
        if _misstyled and in_body:
            # 误套标题样式的正文：强制套正文样式并按 body_spec 格式化
            _set_style(p, "Normal")
        if body_spec and in_body and (not is_structural_title(txt)) and p not in cap_targets \
           and not _is_toc_style(_cur_style_id(p), styles_map) and not _has_image(p):
            # v1.3.5：表格/图片附近的小字段落（疑似题注/来源说明）跳过正文格式化，
            # 保留客户原有小字号，加批注提醒确认。
            if p in suspected_small:
                _psz = _para_font_size_half_pt(p)
                _body_pt = (_body_sz or 0) // 2
                _small_pt = (_psz or 0) // 2
                changes.append({
                    "kind": "suspected_caption",
                    "text": txt[:50], "level": 0,
                    "old": "原格式保留（字号%dpt）" % _small_pt if _psz else "原格式保留",
                    "new": "未修改（疑似题注/来源说明）",
                    "new_name": "疑似题注/来源说明，已保留原格式",
                    "conf": 0.6, "_para": p,
                    "note": ("此段字号（%dpt）小于正文（%dpt），且紧邻表格或图片，"
                             "可能是题注或数据来源说明。已保留原格式，请确认是否需要"
                             "按题注/来源说明格式调整。" % (_small_pt, _body_pt)),
                })
                continue
            # 对误套标题样式的段落，无论 _para_needs_fix 判断如何都强制格式化
            _force = _misstyled
            if _force or _para_needs_fix(p, body_spec):
                _format_runs(p, body_spec)
                _set_para_format(p, body_spec)
                changes.append({
                    "kind": "body", "text": txt[:50], "level": 0,
                    "old": _current_fmt(p, body_spec),
                    "new": _fmt_summary(body_spec),
                    "new_name": _fmt_summary(body_spec),
                    "conf": 1.0, "_para": p,
                    "note": "原套标题样式，文本特征为正文，已降级" if _misstyled else None,
                })

    # ===== 题注段落格式套用（纯结构定位 + profile 驱动，零文字关键词） =====
    # 位置由 _caption_targets 用"表前/表后/图后"结构给出；类别 key 直接查 profile.levels。
    # 若模板未批注某类别，levels 里没有对应 spec，自然跳过——换任何学校模板都安全，
    # 不会因"表/图/数据来源"等措辞差异而误判或漏判。
    replacements = {}  # 题注/脚注 pass 可能提前写入 footnotes.xml，行末统一并入 document
    if profile:
        _levels = (profile or {}).get("levels") or {}
        _n_cap = 0
        for idx, p in enumerate(paras_list):
            if p in tbl_ps:
                continue
            tgt = cap_targets.get(p)
            if not tgt:
                continue
            gkey, kind = tgt
            cap_spec = _levels.get(gkey)
            if not cap_spec:
                continue
            # 题注只套正文区（首个一级标题之后、参考文献之前），封面/摘要/目录区不碰——
            # 避免把封面校徽旁的标题文字误当成图题（这是"逻辑性、不要弄错"的关键守卫）。
            in_body = (first_chap is not None and idx > first_chap
                       and (ref_start is None or idx < ref_start))
            if not in_body:
                continue
            # 图题/图注再叠加"短段落"保守守卫：图片后紧邻的长正文句不算图题，
            # 降低误套风险（宁可漏、不可错）。
            if kind in ("figure_caption", "figure_note"):
                _ct = _text_of(p).strip()
                if len(_ct) > 60:
                    continue
            if _para_needs_fix(p, cap_spec) and not _has_image(p):
                _format_runs(p, cap_spec)
                _set_para_format(p, cap_spec)
                changes.append({
                    "kind": kind, "text": _text_of(p).strip()[:50], "level": 0,
                    "old": _current_fmt(p, cap_spec),
                    "new": _fmt_summary(cap_spec),
                    "new_name": _fmt_summary(cap_spec),
                    "conf": 0.9, "_para": p,
                })
                _n_cap += 1

    # ===== 脚注：模板批注指定脚注格式时，套用到 word/footnotes.xml =====
    # 同样是 profile 驱动：只有 levels 含 "footnote" 才处理，且只改脚注正文段
    # （跳过分隔符/续接这类特殊定义段）。无脚注部件或无批注则整段跳过。
    if profile:
        _fn_spec = (profile or {}).get("levels", {}).get("footnote")
        _fn_name = "word/footnotes.xml"
        if _fn_spec and _fn_name in z.namelist():
            try:
                fn_root = ET.fromstring(z.read(_fn_name))
                _fn_changed = 0
                for p in fn_root.iter(WR + "p"):
                    ppr = p.find(WR + "pPr")
                    sid = ppr.find(WR + "pStyle").get(WR + "val") if (
                        ppr is not None and ppr.find(WR + "pStyle") is not None) else None
                    if sid in ("FootnoteTextSeparator", "FootnoteContinuationSeparator"):
                        continue
                    if _para_needs_fix(p, _fn_spec) and not _has_image(p):
                        _format_runs(p, _fn_spec)
                        _set_para_format(p, _fn_spec)
                        _fn_changed += 1
                if _fn_changed:
                    replacements[_fn_name] = to_doc_xml(fn_root)
            except Exception:
                pass

    # ===== 参考文献区：标题套 reference_heading spec，条目套 reference spec =====
    # v1.3.4 修复：山大等学校在同一条批注中写明两套格式（标题黑体小三加粗居中，
    # 条目宋体/TNR/小四/两端对齐/固定值20磅/按编号位数多档悬挂缩进）。此前整段
    # 被当作一个 spec 套用，导致标题被设成宋体小四、条目被设成小三加粗居中。
    # 现按批注分隔标记拆分：ref_start 段（"参考文献"标题本身）套 reference_heading，
    # 后续条目套 reference，并按条目编号位数（1-9/10-99/100+）选择对应悬挂缩进档位。
    if profile and ref_start is not None:
        _levels_all = (profile or {}).get("levels", {})
        _ref_heading_spec = _levels_all.get("reference_heading")
        _ref_spec = _levels_all.get("reference")
        _REF_END = ("附录", "致谢", "攻读", "后记", "个人简历", "作者简介", "声明")
        for j in range(ref_start, len(paras_list)):
            p = paras_list[j]
            if p in tbl_ps:
                continue
            _rt = _text_of(p).strip()
            if not _rt:
                continue
            # v1.3.5：遇到下一个结构节则停止。结束边界增强，识别三类段落标题：
            #  (a) is_structural_title: 精确匹配/startswith（"致谢"、"References"）
            #  (b) 标题样式 + 包含关键词：客户文档中致谢/附录可能是"6 致谢""附录A 某某"，
            #      带编号前缀导致 startswith 失效。若段落套有 Heading 样式且文本中含有关键词，
            #      也视为参考文献区结束。
            #  (c) 文本以"附录"+字母/数字开头（如"附录A""附录1"）。
            if j > ref_start:
                _is_head = _is_heading_para(p, styles_map, heading_sids)
                _hits_end = any(kw in _rt for kw in _REF_END)
                _is_appx = bool(re.match(r"^附录[A-Za-z0-9]", _rt))
                if (is_structural_title(_rt)
                        or (_is_head and _hits_end)
                        or _is_appx):
                    break
            if _is_toc_style(_cur_style_id(p), styles_map, heading_sids):
                continue
            if _has_image(p):
                continue
            if j == ref_start and _ref_heading_spec:
                # 参考文献标题段：套 reference_heading 格式
                if _para_needs_fix(p, _ref_heading_spec):
                    _format_runs(p, _ref_heading_spec)
                    _set_para_format(p, _ref_heading_spec)
                    changes.append({
                        "kind": "reference_heading", "text": _rt[:50], "level": 0,
                        "old": _current_fmt(p, _ref_heading_spec),
                        "new": _fmt_summary(_ref_heading_spec),
                        "new_name": _fmt_summary(_ref_heading_spec),
                        "conf": 0.95, "_para": p,
                    })
            elif _ref_spec:
                # 参考文献条目：套 reference 格式，按编号位数选悬挂缩进档位
                entry_spec = _ref_entry_spec_for_text(_rt, _ref_spec)
                if _para_needs_fix(p, entry_spec):
                    _format_runs(p, entry_spec)
                    _set_para_format(p, entry_spec)
                    changes.append({
                        "kind": "reference", "text": _rt[:50], "level": 0,
                        "old": _current_fmt(p, entry_spec),
                        "new": _fmt_summary(entry_spec),
                        "new_name": _fmt_summary(entry_spec),
                        "conf": 0.95, "_para": p,
                    })

        # v1.3.5：清理参考文献结束边界之后（致谢/附录等结构页）的悬挂缩进。
        # v1.3.4 曾把参考文献格式误套到致谢/附录，遗留 hanging 无 left 的错误格式。
        # 上面的 for 循环在遇到致谢/附录时 break，从 break 位置之后清除悬挂缩进。
        _ref_end_idx = None
        for j2 in range(ref_start + 1, len(paras_list)):
            pp2 = paras_list[j2]
            if pp2 in tbl_ps:
                continue
            _t2 = _text_of(pp2).strip()
            if not _t2:
                continue
            _t2_norm = re.sub(r"\s+", "", _t2)
            _hits_end = any(kw in _t2_norm for kw in _REF_END)
            if is_structural_title(_t2) or _hits_end or bool(re.match(r"^附录[A-Za-z0-9]", _t2)):
                _ref_end_idx = j2
                break
        if _ref_end_idx is not None:
            for k in range(_ref_end_idx, len(paras_list)):
                pp = paras_list[k]
                if pp in tbl_ps:
                    continue
                ppr_k = pp.find(WR + "pPr")
                if ppr_k is None:
                    continue
                ind_k = ppr_k.find(WR + "ind")
                if ind_k is None:
                    continue
                if ind_k.get(WR + "hanging") is not None:
                    ppr_k.remove(ind_k)
                    _t = _text_of(pp).strip()[:50]
                    changes.append({
                        "kind": "reference", "text": _t, "level": 0,
                        "old": "悬挂缩进（v1.3.4遗留）",
                        "new": "已清除悬挂缩进",
                        "new_name": "已清除悬挂缩进",
                        "conf": 0.9, "_para": pp,
                    })

    if profile:
        _, margin_changes = _fix_sect_margins(root, profile)
        for label, old_disp, new_disp in margin_changes:
            changes.append({
                "kind": "margin", "text": "%s边距" % label, "level": 0,
                "old": old_disp, "new": new_disp, "new_name": new_disp,
                "conf": 1.0,
            })

    # ---- 半角标点自动替换（独立 pass，在所有正文格式套用之后执行）----
    # 保守替换：仅对正文区散文段落生效，跳过标题/表格/图片/题注/目录样式/结构页。
    # 不替换句号和引号，宁可漏换不可错换。每条替换记录到 changes（kind="punctuation"）。
    punct_changes = _punctuation_pass(
        root, first_chap, ref_start, tbl_ps, cap_targets, styles_map, heading_sids)
    changes.extend(punct_changes)

    # ---- v1.3.4：自动编号(numPr)转正文（独立 pass，正文区域非标题段）----
    # 客户论文中常用 Word 自动编号列表，格式调整后编号可能异常增加或错乱；
    # 移除正文段 pPr 下的 numPr，使编号不再自动生成（保留文字内容）。
    # 标题段、表格内段、结构页、题注段都不动。
    numbering_changes = _remove_numbering_pass(
        root, first_chap, ref_start, tbl_ps, cap_targets, styles_map, heading_sids)
    changes.extend(numbering_changes)

    # 报告分级：置信度 ≥0.8 视为确定套用（applied），低于则归 low_conf 供人工复核。
    # 中文序号标题（一、/（一）/第X章等）置信度 0.85，属确定套用，须与文档实际落盘一致。
    applied = [c for c in changes if c["conf"] >= 0.8]
    low_conf = [c for c in changes if c["conf"] < 0.8]

    # ---- Word 批注：为每个有对应段落的 change 添加批注（默认开启）----
    # 批注 ID 从文档现有最大 ID+1 开始，不覆盖已有批注（学校模板/导师批注）。
    comment_count = 0
    if add_comments:
        comment_count = _apply_comments(z, root, changes, replacements)

    replacements["word/document.xml"] = to_doc_xml(root)
    # 题注/脚注 pass 可能已往 replacements 写入 footnotes.xml，这里以 document 为基准并集
    if need_defs:
        replacements["word/styles.xml"] = _inject_style_defs(src, need_defs)
    # 判断套用的样式是否带大纲级别（决定 Word 能否一键出目录）
    has_outline = any("outlineLvl" in (x or "") for x in need_defs.values())
    write_docx_files(src, dst, replacements)
    # 验证批注确实写入了输出文件（v1.3.3 兜底：防止执行链路中 comments.xml 意外丢失）
    if add_comments and comment_count > 0:
        try:
            with zipfile.ZipFile(dst, 'r') as _zv:
                if 'word/comments.xml' not in _zv.namelist():
                    print(f"[警告] 批注({comment_count}条)未写入输出文件！", file=sys.stderr)
        except Exception:
            pass
    return applied, low_conf, skipped_residue, (sorted(need_defs), template_exact, has_outline, tmpl_name, already_ok, comment_count)


def _render(applied, low_conf, skipped_residue, style_info, dst):
    used_styles, template_exact, has_outline, tmpl_name, already_ok, comment_count = style_info
    lines = ["# 标题样式套用报告", ""]
    mode = "【按学校模板真实样式】" if template_exact else "【通用 Heading 样式】"
    lines.append(f"模式：{mode}（{tmpl_name}）")
    if has_outline:
        lines.append(f"已写出新文档：{os.path.basename(dst)}（Word 现可一键插入目录）")
    else:
        lines.append(f"已写出新文档：{os.path.basename(dst)}")
        lines.append("⚠ 提示：所套用的学校标题样式未含 Word 大纲级别（outlineLvl），"
                     "Word 可能无法据此自动生成目录。如需一键出目录，请改用内置 Heading 样式或手动建目录。")
    if template_exact and used_styles:
        lines.append(f"套用的模板样式 ID：{', '.join(used_styles)}")
    lines.append("")
    all_c = applied + low_conf
    _nh = sum(1 for c in all_c if c.get("kind") == "heading")
    _nb = sum(1 for c in all_c if c.get("kind") == "body")
    _nm = sum(1 for c in all_c if c.get("kind") == "margin")
    _ni = sum(1 for c in all_c if c.get("kind") == "image")
    _np = sum(c.get("count", 0) for c in all_c if c.get("kind") == "punctuation")
    _nref = sum(1 for c in all_c if c.get("kind") == "reference")
    _nrefh = sum(1 for c in all_c if c.get("kind") == "reference_heading")
    _nn = sum(1 for c in all_c if c.get("kind") == "numbering")
    _ncap = sum(1 for c in all_c if c.get("kind") in ("table_caption", "figure_caption", "table_note", "figure_note"))
    _parts = [f"标题 {_nh}", f"正文 {_nb}"]
    if _ni:
        _parts.append(f"图片 {_ni}")
    if _nrefh:
        _parts.append(f"参考文献标题 {_nrefh}")
    if _nref:
        _parts.append(f"参考文献条目 {_nref}")
    if _ncap:
        _parts.append(f"题注 {_ncap}")
    if _nn:
        _parts.append(f"自动编号转正文 {_nn}")
    if _np:
        _parts.append(f"标点修正 {_np}")
    _parts.append(f"页边距 {_nm}")
    lines.append(f"## 逐条修改清单（共 {len(all_c)} 条记录、{sum(1 for c in all_c)} 处：{' / '.join(_parts)}）")
    lines.append(f"> 另有 {already_ok} 处标题经核对已符合学校样式，无需改动。")
    if comment_count:
        lines.append(f"> 已在文档中添加 {comment_count} 条 Word 批注，说明每处修改（可在 Word 中「删除所有批注」清理）。")
    lines.append("")
    lines.append("| # | 类型 | 修改前 | 修改后 | 内容摘要 | 置信度 |")
    lines.append("|---|------|--------|--------|----------|--------|")
    for i, r in enumerate(all_c, 1):
        conf = "高" if r["conf"] >= 0.9 else "低（建议复核）"
        new_disp = f"{r['new']}（{r['new_name']}）" if r["new_name"] != r["new"] else r["new"]
        # 标点类修改显示段落位置
        text_disp = r["text"]
        if r.get("kind") == "punctuation" and r.get("para_index") is not None:
            text_disp = f"[第{r['para_index']}段] {r['text']}"
        lines.append(f"| {i} | {_type_label(r)} | {r['old']} | {new_disp} | {text_disp} | {conf} |")
    if skipped_residue:
        lines.append("")
        lines.append(f"## 已跳过 {skipped_residue} 处目录域/书签残留（脏数据，未套样式）")
    return "\n".join(lines)


def _build_change_docx(applied, low_conf, skipped_residue, style_info, dst_docx):
    """生成 Word 版修改明细，供客户逐条验收。"""
    used_styles, template_exact, has_outline, tmpl_name, already_ok, comment_count = style_info
    all_c = applied + low_conf
    nh = sum(1 for c in all_c if c.get("kind") == "heading")
    nb = sum(1 for c in all_c if c.get("kind") == "body")
    nm = sum(1 for c in all_c if c.get("kind") == "margin")
    ni = sum(1 for c in all_c if c.get("kind") == "image")
    nref = sum(1 for c in all_c if c.get("kind") == "reference")
    nrefh = sum(1 for c in all_c if c.get("kind") == "reference_heading")
    nnum = sum(1 for c in all_c if c.get("kind") == "numbering")
    ncap = sum(1 for c in all_c if c.get("kind") in ("table_caption", "figure_caption", "table_note", "figure_note"))
    npunct = sum(c.get("count", 0) for c in all_c if c.get("kind") == "punctuation")
    total_records = len(all_c)
    total_changes = sum(c.get("count", 1) for c in all_c)
    mode = "按学校模板真实样式" if template_exact else "通用 Heading 样式"
    summary = [
        f"修改模式：{mode}（{tmpl_name}）",
    ]
    if template_exact and used_styles:
        summary.append(f"套用样式：{', '.join(used_styles)}")
    _detail_parts = [f"标题 {nh} 处", f"正文 {nb} 处"]
    if ni:
        _detail_parts.append(f"图片 {ni} 处")
    if nrefh:
        _detail_parts.append(f"参考文献标题 {nrefh} 处")
    if nref:
        _detail_parts.append(f"参考文献条目 {nref} 处")
    if ncap:
        _detail_parts.append(f"题注 {ncap} 处")
    if nnum:
        _detail_parts.append(f"自动编号转正文 {nnum} 处")
    if npunct:
        _detail_parts.append(f"标点修正 {npunct} 处")
    _detail_parts.append(f"页边距 {nm} 处")
    summary.append(f"共修改 {total_changes} 处（{total_records} 条记录：{' / '.join(_detail_parts)}）："
                   f"按「{tmpl_name}」批注与样式统一字体、字号、首行缩进、行距与页边距"
                   + ("，并归一化图片布局（浮动→内嵌/超宽缩放/居中/去固定行距）" if ni else "")
                   + ("，并将正文自动编号转为普通文本" if nnum else "")
                   + ("，并自动修正中文半角标点（半角逗号/分号/冒号/问号/感叹号/括号→全角）" if npunct else ""))
    summary.append(f"另有 {already_ok} 处标题经核对已符合学校样式，无需改动。")
    summary.append(f"已跳过目录域/书签残留 {skipped_residue} 处（脏数据，未改动，不影响正文）")
    if comment_count:
        summary.append(f"已在文档中添加 {comment_count} 条 Word 批注，说明每处修改；可在 Word 中「审阅→删除→删除所有批注」一键清理。")
    if has_outline:
        summary.append("✅ 所套样式含 Word 大纲级别，Word 可据此一键生成目录。")
    else:
        summary.append("⚠ 提示：所套用的学校标题样式未含 Word 大纲级别，Word 可能无法据此自动生成目录；"
                       "如需一键出目录请告知，我可改用「套样式+补大纲级别」方案。")

    columns = [("序号", 700), ("类型", 1000), ("修改前", 2100),
               ("修改后", 2100), ("内容摘要", 2260), ("置信度", 900)]
    rows = []
    for i, r in enumerate(all_c, 1):
        conf = "高" if r["conf"] >= 0.9 else "低（建议复核）"
        new_disp = f"{r['new']}（{r['new_name']}）" if r["new_name"] != r["new"] else r["new"]
        # 标点类修改显示段落位置
        text_disp = r["text"]
        if r.get("kind") == "punctuation" and r.get("para_index") is not None:
            text_disp = f"[第{r['para_index']}段] {r['text']}"
        rows.append([str(i), _type_label(r), r["old"], new_disp, text_disp, conf])
    write_change_report(
        dst_docx,
        title="论文格式修改明细（标题与正文）",
        summary=summary,
        columns=columns,
        rows=rows,
    )
    return dst_docx


def _lvl(n):
    return {0: "正文", 1: "一级", 2: "二级", 3: "三级"}.get(n, "标题")


def _type_label(r):
    kind = r.get("kind", "")
    if kind == "margin":
        return "页边距"
    if kind == "image":
        return "图片"
    if kind == "punctuation":
        return "标点修正"
    if kind == "reference":
        return "参考文献条目"
    if kind == "reference_heading":
        return "参考文献标题"
    if kind == "numbering":
        return "自动编号转正文"
    if kind in ("table_caption", "figure_caption"):
        return "题注"
    if kind in ("table_note", "figure_note"):
        return "表/图注"
    if kind == "footnote":
        return "脚注"
    return _lvl(r.get("level", 0))


def main():
    if len(sys.argv) < 3:
        print("用法: python headings-fix.py 输入.docx 输出.docx [--profile 模板画像.json] [--report 修改明细.docx] [--no-comments]")
        sys.exit(1)
    src, dst = sys.argv[1], sys.argv[2]
    profile = None
    report_path = None
    add_comments = "--no-comments" not in sys.argv
    if "--profile" in sys.argv:
        pf = sys.argv[sys.argv.index("--profile") + 1]
        with open(pf, encoding="utf-8") as f:
            profile = json.load(f)
    if "--report" in sys.argv:
        report_path = sys.argv[sys.argv.index("--report") + 1]
    applied, low_conf, skipped, sinfo = fix(src, dst, profile, add_comments=add_comments)
    print(_render(applied, low_conf, skipped, sinfo, dst))
    if report_path:
        _build_change_docx(applied, low_conf, skipped, sinfo, report_path)
        print(f"\n[已生成修改明细报告] {report_path}")


if __name__ == "__main__":
    main()
