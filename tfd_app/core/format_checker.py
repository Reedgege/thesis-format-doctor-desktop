# -*- coding: utf-8 -*-
# Copyright (c) 2026 芦苇（山东大学 MBA）. All rights reserved.
# 原创作品 | 禁止未经授权转售、二次分发或抄袭 | 授权用户可在扣子/虾评平台内使用
# build: 202608191649 | version: 1.3.5 | file_sha256: 22815c6048934f7a
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
文档格式自动校验脚本（纯标准库，零第三方依赖，复用 docxutils 单一真源）

用法：
  python format-checker.py 论文.docx
  python format-checker.py 论文.docx --json        # 输出 JSON，便于平台展示/计费
  python format-checker.py 论文.docx -o report.md  # 报告写文件

【通用规范体检】——不依赖学校模板，按学术通用红线检查。
  注意：通用检查只能覆盖"各校普遍要求"的规则（全角标点、标题层级、目录能力、
  三线表、文献格式、字体/字号/缩进/页边距的明显异常）。学校特有的精确数值
  （如上边距必须 2.5cm、正文必须宋体小四、行距固定 20 磅、指定标题字号）
  需【上传模板】做"模板驱动诊断"才能逐项核对。无模板时本报告对"学校特有
  规则"存在漏报可能，建议补传模板复检。

检查维度（10 项）：
  1. 章节标题层级与编号连续性（第N章 / 1.1 / 1.1.1 缺章、跳号）
  2. 标题样式与"一键出目录"能力（内置 Heading vs 自定义样式）
  3. 三线表规范（仅 顶/底/栏目 三条横线）
  4. 中文字体分布（≥3 种疑似随意混用）
  5. 字号分布（正文与标题是否明显分层）
  6. 首行缩进（中文正文通常 2 字符 = firstLineChars 200）
  7. 页边距（上下/左右对称）
  8. 中文半角标点（排除参考文献区，降低误报）
  9. 参考文献 [n] 连续性 + GB/T 7714 出版年
 10. 图/表题编号与表格数匹配
"""
import sys
import os
import re
import json
from collections import Counter, defaultdict

_HERE = os.path.dirname(os.path.abspath(__file__))
if _HERE not in sys.path:
    sys.path.insert(0, _HERE)
import docxutils

__author__ = "芦苇"
__copyright__ = "Copyright (c) 2026 芦苇（山东大学 MBA）"
__version__ = "1.3.5"

W = docxutils.WR  # 注意：元素标签需用 Clark 记法 {ns}tag，docxutils.WR 才是带花括号的前缀
CJK = re.compile(r'[\u3400-\u9fff]')
HALF_PUNCT = set('.,;:!?()')  # 不含 [] {}（文献引用标记非标点错误）与引号（中英文混用常见）
SEV_ORDER = {'高': 0, '中': 1, '低': 2, 'info': 3}
SEV_ICON = {'高': '🔴', '中': '🟡', '低': '🟢', 'info': 'ℹ️'}
# 结构页标题（结论/致谢/附录等），detect_heading 未覆盖，但属标题而非正文，不对其套用正文缩进/字号规范。
# 不再写死中英文关键词：统一用 docxutils.is_structural_title（跨校，覆盖中英文主流写法）。


_INTEGRITY_TAG = "Q29weXJpZ2h0IChjKSAyMDI2IOiKpuiLh++8iOWxseS4nOWkp+WtpiBNQkHvvIkuIEFsbCByaWdodHMgcmVzZXJ2ZWQuIHwgYnVpbGQ6MjAyNjA4MTkxNjQ5"  # noqa
def cn2int(s):
    if s.isdigit():
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


def find_ref_start(paras):
    for idx, p in enumerate(paras):
        t = (p.get('text') or '').strip()
        if docxutils.is_reference_heading(t):
            return idx
    return None


def collect_stats(z, root, paras):
    """一次性收集统计信息（供问题判定与报告画像共用）。"""
    font_c = Counter()
    sz_c = Counter()
    for r in root.iter(W + 'r'):
        rpr = r.find(W + 'rPr')
        if rpr is None:
            continue
        rf = rpr.find(W + 'rFonts')
        if rf is not None:
            ea = rf.get(W + 'eastAsia')
            if ea:
                font_c[ea] += 1
        sz = rpr.find(W + 'sz')
        if sz is not None and sz.get(W + 'val'):
            sz_c[sz.get(W + 'val')] += 1
    firstline_c = Counter()
    heading_builtin = 0
    heading_custom = 0
    ref_start = find_ref_start(paras)
    body_total = 0
    body_no_indent = 0
    for idx, p in enumerate(paras):
        st = p.get('style')
        ppr = p.get('pPr') or {}
        fl = ppr.get('firstLineChars') or ppr.get('firstLine')
        firstline_c[str(fl)] += 1
        t = (p.get('text') or '').strip()
        lvl = docxutils.detect_heading(t)[0]
        is_heading = lvl is not None or (st and st.startswith('Heading'))
        in_ref = ref_start is not None and idx >= ref_start
        if lvl is not None:
            if st and st.startswith('Heading'):
                heading_builtin += 1
            else:
                heading_custom += 1
        # 仅对"正文候选段"（非标题/非参考文献区/非表格内）评估首行缩进，避免误报
        if not is_heading and not in_ref and not p.get('in_table'):
            body_total += 1
            if fl is None:
                body_no_indent += 1
    margins = docxutils.margins(root)
    tables = docxutils.tables(root)
    media = sum(1 for n in z.namelist() if n.startswith('word/media/') and not n.endswith('/'))
    return {
        'font_c': font_c, 'sz_c': sz_c, 'firstline_c': firstline_c,
        'heading_builtin': heading_builtin, 'heading_custom': heading_custom,
        'body_total': body_total, 'body_no_indent': body_no_indent,
        'margins': margins, 'tables': tables, 'media': media,
        'para_count': len(paras),
    }


def check_headings(paras):
    issues = []
    seen_chapters = []
    numeric = []
    for p in paras:
        t = (p.get('text') or '').strip()
        if not t:
            continue
        m = re.match(r'^第([0-9一二三四五六七八九十]+)章', t)
        if m:
            seen_chapters.append((cn2int(m.group(1)), t))
        mn = re.match(r'^(\d+)(?:\.(\d+))?(?:\.(\d+))?(?:\.(\d+))?[\s、.．]', t)
        if mn:
            nums = tuple(int(x) for x in mn.groups() if x is not None)
            numeric.append((len(nums), nums, t))
    if seen_chapters:
        maxc = max(n for n, _ in seen_chapters)
        present = {n for n, _ in seen_chapters}
        for c in range(1, maxc + 1):
            if c not in present:
                issues.append(('高', f'缺少第{c}章标题（检测到章节跨度 1–{maxc} 但有缺）'))
    by_parent = defaultdict(list)
    for lvl, nums, _t in numeric:
        if lvl == 1:
            continue
        by_parent[nums[:-1]].append(nums[-1])
    for parent, kids in by_parent.items():
        ks = sorted(set(kids))
        full = list(range(ks[0], ks[-1] + 1))
        for g in full:
            if g not in ks:
                issues.append(('中', f'编号跳号：缺少 {".".join(map(str, parent + (g,)))}（上级 {".".join(map(str, parent))} 下）'))
    return issues


def check_heading_styles(stats):
    issues = []
    tot = stats['heading_builtin'] + stats['heading_custom']
    if tot > 0 and stats['heading_custom'] > 0:
        issues.append(('中', f'检测到 {tot} 处文本层级标题，其中 {stats["heading_custom"]} 处使用「自定义样式」、仅 {stats["heading_builtin"]} 处使用 Word 内置 Heading 样式 → 文档大概率无法直接自动生成目录，建议用 headings-fix.py 套标题样式。'))
    return issues


def check_three_line(tables):
    issues = []
    for i, b in enumerate(tables, 1):
        present = set(b.keys())
        ok = present >= {'top', 'bottom', 'insideH'} and not (present & {'left', 'right', 'insideV'})
        if not ok:
            missing = [k for k in ('top', 'bottom', 'insideH') if k not in present]
            extra = [k for k in ('left', 'right', 'insideV') if k in present]
            detail = []
            if missing:
                detail.append('缺 ' + ','.join(missing))
            if extra:
                detail.append('多出 ' + ','.join(extra))
            issues.append(('中', f'表 {i} 非标准三线表（{"；".join(detail)}）。三线表应仅含顶线/底线/栏目线三条横线，无左右边线、无竖线。'))
    return issues


def check_fonts(font_c):
    issues = []
    # 阈值≥3：标题黑体+正文宋体属正常分层，不误报；≥3 种才疑似随意混用
    if len(font_c) >= 3:
        top = '、'.join(f'{k}({v})' for k, v in font_c.most_common(6))
        issues.append(('中', f'中文字体疑似随意混用：检测到 {len(font_c)} 种（{top}）。中文正文通常应统一（如正文宋体、英文 Times New Roman），标题可用黑体。'))
    return issues


def check_sizes(sz_c):
    issues = []
    # 阈值>3：标题(18pt)+正文(12pt)属正常分层，不误报
    if len(sz_c) > 3:
        top = '、'.join(f'{int(s)//2}pt' for s, _ in sz_c.most_common(6))
        issues.append(('低', f'字号种类较多（{len(sz_c)} 种：{top}），正文与标题应明显分层（如正文小四/12pt、一级标题小二/18pt）。'))
    return issues


def check_indent(body_total, body_no_indent):
    issues = []
    # 仅统计"正文候选段"：标题/参考文献/表格内段落本就不该缩进，混入会放大误报
    if body_total >= 5 and body_no_indent == body_total:
        issues.append(('中', f'正文段落（共 {body_total} 段）均未设置首行缩进，中文正文通常需首行缩进 2 字符（firstLineChars=200）。'))
    elif body_total >= 10 and body_no_indent > body_total * 0.9:
        issues.append(('中', f'绝大多数正文段落（{body_no_indent}/{body_total}）未设置首行缩进，中文正文通常需首行缩进 2 字符。'))
    return issues


def check_margins(margins):
    issues = []
    if not margins:
        return issues

    def cm(v):
        try:
            return round(int(v) / 1440 * 2.54, 2)
        except Exception:
            return None

    top, bot = cm(margins.get('top')), cm(margins.get('bottom'))
    left, right = cm(margins.get('left')), cm(margins.get('right'))
    if top is not None and bot is not None and abs(top - bot) >= 0.5:
        issues.append(('中', f'页边距上下不对称（上 {top}cm / 下 {bot}cm），多数学校要求上下一致（常见 2.5–3cm）。'))
    if left is not None and right is not None and abs(left - right) >= 0.5:
        issues.append(('中', f'页边距左右不对称（左 {left}cm / 右 {right}cm），通常要求左右一致。'))
    return issues


def check_punct(paras):
    issues = []
    count = 0
    examples = []
    ref_start = find_ref_start(paras)
    for idx, p in enumerate(paras):
        t = p.get('text') or ''
        if ref_start is not None and idx >= ref_start:
            continue  # 参考文献区合法半角（作者. 题名. 出处）豁免，避免误报
        if not CJK.search(t):
            continue
        for i, ch in enumerate(t):
            if ch in HALF_PUNCT:
                prev = t[i - 1] if i > 0 else ''
                nxt = t[i + 1] if i + 1 < len(t) else ''
                if CJK.search(prev) or CJK.search(nxt):
                    count += 1
                    if len(examples) < 5:
                        s = max(0, i - 8)
                        e = min(len(t), i + 8)
                        examples.append(t[s:e].replace('\n', ' '))
                    break
    if count:
        issues.append(('中', f'中文正文检测到 {count} 处半角标点（应为全角 ，。；：！？（））。示例：' + ' │ '.join(examples)))
    return issues


def check_references(paras):
    issues = []
    ref_start = find_ref_start(paras)
    if ref_start is None:
        # 未检测到参考文献章节：属内容缺失而非格式缺陷，不计入可整改问题，
        # 静默跳过检查，避免误导 issue_count / 严重度分布。
        return issues
    block = [(p.get('text') or '').strip() for p in paras[ref_start + 1: ref_start + 200]]
    block = [t for t in block if t]
    nums = []
    no_year = 0
    total = 0
    for t in block:
        if re.match(r'^第[0-9一二三四五六七八九十]+章', t):
            break
        m = re.search(r'\[(\d+)\]', t)
        if m:
            nums.append(int(m.group(1)))
            total += 1
            if not re.search(r'(19|20)\d{2}', t):
                no_year += 1
    if nums:
        mn, mx = min(nums), max(nums)
        gaps = [n for n in range(mn, mx + 1) if n not in nums]
        if gaps:
            issues.append(('中', f'参考文献编号不连续，缺 [n]：{gaps[:10]}'))
        if no_year:
            issues.append(('低', f'有 {no_year}/{total} 条参考文献未检测到出版年（GB/T 7714 通常需含年份）'))
    return issues


def check_captions(paras, tables):
    issues = []
    tab = sum(1 for p in paras if re.search(r'表\s*[0-9一二三四五六七八九十]+\s*[-\.]\s*[0-9一二三四五六七八九十]+', p.get('text', '')))
    if tab < len(tables):
        issues.append(('低', f'检测到 {len(tables)} 个表格，但仅 {tab} 处"表X-X"题注，部分表可能缺题注或未按章编号。'))
    return issues


def _cm_to_twips(cm):
    try:
        return int(round(float(cm) * 1440 / 2.54))
    except Exception:
        return None


def _eff_level_spec(profile, key, cat_key):
    """取 per-level 要素规范：优先 profile.levels（样式定义+批注合并），回退旧 spec.cats。"""
    levels = profile.get("levels") or {}
    ls = levels.get(key)
    if ls:
        return ls
    cats = (profile.get("spec") or {}).get("cats") or {}
    return cats.get(cat_key) or {}


def _want_line(cs):
    """把 per-level spec 的行距规则换算成 OOXML 的 (line-twips, lineRule)。"""
    lr = cs.get("line_rule")
    val = cs.get("line_val")
    if lr == "exact" and val:
        return int(val) * 20, "exact"
    if lr == "atLeast" and val:
        return int(val) * 20, "atLeast"
    if lr == "auto" and val:
        return int(val), "auto"
    if lr == "single":
        return 240, "auto"
    return None, None


# 与 headings-fix 完全一致的「聚集判定」：彼此间隔 <=K 的"第X章/节"视为正文列举
# （如 1.2.2 研究内容里逐条概述七个章节），不当真实标题；诊断时跳过其标题级要素核查，
# 避免对概述段误报"应为黑体/小三/居中"。K 取段落索引间隔，与 fix 的 root.iter 顺序一致。
CHAP_CAND = re.compile(r"^第[一二三四五六七八九十百千\d]+[章篇编节]")
K_CLUSTER = 3


def _demoted_chapter_indices(paras):
    """返回应当「降级为正文」的第X章段落索引集合（复刻 headings-fix 聚集判定）。"""
    chap_idx = [i for i, p in enumerate(paras)
                if CHAP_CAND.match((p.get("text") or "").strip())]
    demote = set()
    for a in chap_idx:
        if any(abs(a - b) <= K_CLUSTER for b in chap_idx if b != a):
            demote.add(a)
    return demote


def template_driven_checks(paras, stats, profile):
    """模板驱动诊断：以模板批注（profile.spec，用户明确"以批注为准"）为权威规范，
    逐项核对客户论文的页边距、标题样式归属、标题/正文字体字号与缩进。

    仅报告"与批注要求不符"的项；无法从 run/段落属性读取到的（如继承自样式、
    无显式设置）不做误判。返回 [(severity, msg)]。
    """
    issues = []
    spec = profile.get("spec") or {}
    cats = spec.get("cats") or {}
    page = spec.get("page")

    # ---- 1. 页边距（批注 cm -> twips 比对）----
    if page:
        m = stats.get("margins") or {}
        # 容忍 ±0.12cm（约 ±3 磅）的误差
        tol = _cm_to_twips(0.12) or 17
        for key, cmkey, label in (("top", "top_cm", "上"),
                                  ("bottom", "bottom_cm", "下"),
                                  ("left", "left_cm", "左"),
                                  ("right", "right_cm", "右")):
            if cmkey not in page:
                continue
            want = _cm_to_twips(page[cmkey])
            got = m.get(key)
            if want is None or got is None:
                continue
            if abs(int(got) - want) > tol:
                issues.append(("中", f'页边距{label}边应为 {page[cmkey]}cm（批注要求），'
                                      f'当前约 {round(int(got)/1440*2.54,2)}cm，不符。'))

    # ---- 2/3. 标题与正文的字体/字号/缩进/样式归属 ----
    ref_start = find_ref_start(paras)
    # 正文核查范围：从"第一章/1 绪论"之后到"参考文献"之前，避免把封面/目录/摘要误判为正文。
    first_chap = None
    for i, p in enumerate(paras):
        tt = (p.get("text") or "").strip()
        if not tt:
            continue
        if docxutils.is_toc_residue(tt):
            continue
        lv, _ = docxutils.detect_heading(tt)
        if lv == 1:
            first_chap = i
            break
    _demoted = _demoted_chapter_indices(paras)
    for idx, p in enumerate(paras):
        t = (p.get("text") or "").strip()
        if not t:
            continue
        # 目录域/书签残留（.doc 转 .docx 带入）一律跳过，避免污染诊断
        if docxutils.is_toc_residue(t):
            continue
        lvl, _ = docxutils.detect_heading(t)
        # 聚集的"第X章"概述段（1.2.2 研究内容式）已降级为正文，跳过标题级要素核查
        if idx in _demoted:
            continue
        in_ref = ref_start is not None and idx >= ref_start
        in_body = (first_chap is not None and idx > first_chap
                   and (ref_start is None or idx < ref_start))
        rpr = p.get("rpr") or {}
        ppr = p.get("pPr") or {}

        if lvl in (1, 2, 3):
            hs = (profile.get("headingStyles") or {}).get(str(lvl)) or {}
            exp_sid = hs.get("styleId")
            exp_name = hs.get("name") or exp_sid
            cur = p.get("style")
            if exp_sid and cur and cur != exp_sid:
                issues.append(("中", f'标题「{t[:24]}」未使用学校指定样式（应为 {exp_name} / {exp_sid}），'
                                      f'当前样式 {cur}。'))
            cs = _eff_level_spec(profile, str(lvl), "h%d" % lvl)
            lab = "一二三"[lvl - 1] + "级标题"
            # 字体/字号仅在段落有显式 run 设置时才比对，避免对继承值误判
            if cs.get("zh_font") and rpr.get("eastAsia") and rpr.get("eastAsia") != cs["zh_font"]:
                issues.append(("低", f'{lab}「{t[:24]}」中文字体应为 {cs["zh_font"]}（模板），'
                                      f'实为 {rpr.get("eastAsia")}。'))
            if cs.get("sz") and rpr.get("sz") and str(rpr.get("sz")) != str(cs["sz"]):
                issues.append(("低", f'{lab}「{t[:24]}」字号应为 {cs.get("size")}'
                                      f'（sz={cs["sz"]}，模板），实为 sz={rpr.get("sz")}。'))
            # 对齐方式
            if cs.get("align") and ppr.get("jc") and ppr.get("jc") != cs["align"]:
                issues.append(("低", f'{lab}「{t[:24]}」对齐方式应为 {cs["align"]}（模板），'
                                      f'实为 {ppr.get("jc")}。'))
            # 段前 / 段后（容许 1 磅误差）
            if cs.get("before_pt") is not None and ppr.get("before") is not None:
                if abs(int(ppr.get("before")) - cs["before_pt"] * 20) > 20:
                    issues.append(("低", f'{lab}「{t[:24]}」段前应为 {cs["before_pt"]} 磅（模板），'
                                          f'实为 {int(ppr.get("before")) // 20} 磅。'))
            if cs.get("after_pt") is not None and ppr.get("after") is not None:
                if abs(int(ppr.get("after")) - cs["after_pt"] * 20) > 20:
                    issues.append(("低", f'{lab}「{t[:24]}」段后应为 {cs["after_pt"]} 磅（模板），'
                                          f'实为 {int(ppr.get("after")) // 20} 磅。'))
            # 行距
            if cs.get("line_rule") and cs.get("line_val"):
                want_line, want_rule = _want_line(cs)
                got_line = ppr.get("line")
                got_rule = ppr.get("lineRule") or "auto"
                if got_line is None or int(got_line) != want_line or got_rule != want_rule:
                    issues.append(("低", f'{lab}「{t[:24]}」行距应为 {cs["line_rule"]} '
                                          f'{cs.get("line_val")}（模板）。'))
        elif in_body and (not in_ref) and (not p.get("in_table")) and lvl is None \
                and not docxutils.is_structural_title(t):
            # 仅对"正文正文区"的散文段落核查（排除封面/目录/摘要/声明/参考文献）
            cs = _eff_level_spec(profile, "body", "body")
            if cs.get("zh_font") and rpr.get("eastAsia") and rpr.get("eastAsia") != cs["zh_font"]:
                issues.append(("低", f'正文「{t[:24]}…」中文字体应为 {cs["zh_font"]}（模板），'
                                      f'实为 {rpr.get("eastAsia")}。'))
            if cs.get("sz") and rpr.get("sz") and str(rpr.get("sz")) != str(cs["sz"]):
                issues.append(("低", f'正文「{t[:24]}…」字号应为 {cs.get("size")}'
                                      f'（sz={cs["sz"]}，模板），实为 sz={rpr.get("sz")}。'))
            # 首行缩进（2 字符 = firstLineChars 200）
            if cs.get("indent_chars"):
                want_chars = cs["indent_chars"] * 100
                got = ppr.get("firstLineChars")
                if got is None and not (ppr.get("firstLine")):
                    issues.append(("中", f'正文「{t[:24]}…」未设置首行缩进，模板要求首行缩进 {cs["indent_chars"]} 字符。'))
                elif got is not None and abs(int(got) - want_chars) > 20:
                    issues.append(("中", f'正文「{t[:24]}…」首行缩进应为 {cs["indent_chars"]} 字符，'
                                      f'当前约 {int(got)//100} 字符。'))
            # 对齐方式
            if cs.get("align") and ppr.get("jc") and ppr.get("jc") != cs["align"]:
                issues.append(("低", f'正文「{t[:24]}…」对齐方式应为 {cs["align"]}（模板），'
                                      f'实为 {ppr.get("jc")}。'))
            # 段前 / 段后
            if cs.get("before_pt") is not None and ppr.get("before") is not None:
                if abs(int(ppr.get("before")) - cs["before_pt"] * 20) > 20:
                    issues.append(("低", f'正文「{t[:24]}…」段前应为 {cs["before_pt"]} 磅（模板），'
                                          f'实为 {int(ppr.get("before")) // 20} 磅。'))
            if cs.get("after_pt") is not None and ppr.get("after") is not None:
                if abs(int(ppr.get("after")) - cs["after_pt"] * 20) > 20:
                    issues.append(("低", f'正文「{t[:24]}…」段后应为 {cs["after_pt"]} 磅（模板），'
                                          f'实为 {int(ppr.get("after")) // 20} 磅。'))
            # 行距
            if cs.get("line_rule") and cs.get("line_val"):
                want_line, want_rule = _want_line(cs)
                got_line = ppr.get("line")
                got_rule = ppr.get("lineRule") or "auto"
                if got_line is None or int(got_line) != want_line or got_rule != want_rule:
                    issues.append(("低", f'正文「{t[:24]}…」行距应为 {cs["line_rule"]} '
                                          f'{cs.get("line_val")}（模板）。'))
    return issues


# ---------------------------------------------------------------------------
# 结构页只读诊断（封面/摘要/声明/目录）——不修改文件，仅提示
# ---------------------------------------------------------------------------
# 封面识别关键词（出现任意一个即认为该区域是封面）
_COVER_KEYWORDS = ("毕业论文", "学位论文", "分类号", "密级", "UDC", "学校代码")
# 封面应包含的要素（关键词, 显示名）
_COVER_FIELDS = [
    (("大学", "学院", "学校", "单位", "研究院"), "学校/单位名"),
    (("论文题目", "题目", "题名"), "论文题目"),
    (("作者姓名", "姓名", "作者", "研究生姓名"), "作者姓名"),
    (("指导教师", "导师", "导师姓名", "指导老师"), "导师姓名"),
    (("专业", "学科", "学科专业", "专业学位"), "专业/学科"),
    (("日期", "年", "月"), "日期"),
]
# 声明识别关键词
_DECLARATION_KEYWORDS = ("原创性声明", "学位论文版权使用授权书", "版权使用授权",
                         "学术诚信声明", "原创声明")
# 摘要关键词（段落级，strip 后精确或包含）
_ABSTRACT_KEYWORDS = ("摘要", "摘 要", "ABSTRACT", "Abstract")
_KEYWORDS_LABELS = ("关键词", "Key Words", "Keywords", "KEY WORDS", "关键字")


def check_structural_pages(paras, stats):
    """对封面/摘要/声明/目录做只读诊断，返回 [(severity, msg)] 和尾部说明字符串。

    扫描前30个段落识别结构页（目录可能在更靠后的位置，通过 Heading 样式判断）。
    所有问题均为 info/低危，不自动修改，建议对照学校模板手动确认。
    """
    issues = []
    scan = paras[:30]
    scan_texts = [(p.get('text') or '').strip() for p in scan]
    joined = '\n'.join(scan_texts)

    # ---- 1. 封面 ----
    cover_hit = [t for t in scan_texts if any(kw in t for kw in _COVER_KEYWORDS)]
    if cover_hit:
        missing_fields = []
        for kws, label in _COVER_FIELDS:
            if not any(any(kw in t for kw in kws) for t in scan_texts):
                missing_fields.append(label)
        if missing_fields:
            issues.append(('低', f'封面区域可能缺少：{ "、".join(missing_fields) }'
                                  '（结构页各校差异大，建议对照学校模板确认）。'))
        else:
            issues.append(('info', '封面要素（学校/题目/作者/导师/专业/日期）基本齐全。'))
    else:
        issues.append(('info', '未在前30段检测到典型封面关键词（毕业论文/分类号/密级/UDC/学校代码等），'
                                '若文档确实含封面请忽略此提示。'))

    # ---- 2. 摘要 ----
    abstract_idx = None
    for i, t in enumerate(scan_texts):
        if not t:
            continue
        for kw in _ABSTRACT_KEYWORDS:
            if t == kw or t.startswith(kw):
                abstract_idx = i
                break
        if abstract_idx is not None:
            break
    if abstract_idx is not None:
        # 检查关键词（在摘要标题之后的若干段内查找）
        look_ahead = scan_texts[abstract_idx + 1: abstract_idx + 15]
        has_keywords = any(any(kw in t for kw in _KEYWORDS_LABELS) for t in look_ahead)
        if not has_keywords:
            issues.append(('低', '摘要区域未检测到「关键词/Key Words」标记，'
                                  '中文摘要通常需附3–5个关键词。'))
        # 摘要正文长度（去掉"摘要"标题行和关键词行，估算正文字数）
        body_texts = []
        for t in look_ahead:
            if any(kw in t for kw in _KEYWORDS_LABELS):
                continue
            if t.startswith('ABSTRACT') or t.startswith('Abstract'):
                continue
            body_texts.append(t)
        abstract_body = ''.join(body_texts)
        # 去除空白后估算中文字符数
        abstract_len = len(re.sub(r'\s', '', abstract_body))
        if abstract_len < 100:
            issues.append(('低', f'摘要正文较短（约 {abstract_len} 字，<100字），'
                                  '多数学校要求中文摘要 300–500 字以上。'))
        elif has_keywords:
            issues.append(('info', f'摘要区域检测到关键词，摘要正文约 {abstract_len} 字。'))
    else:
        issues.append(('info', '未在前30段检测到「摘要/Abstract」标题。'))

    # ---- 3. 声明 ----
    has_declaration = any(any(kw in t for kw in _DECLARATION_KEYWORDS) for t in scan_texts)
    if not has_declaration:
        issues.append(('低', '未检测到「原创性声明/学位论文版权使用授权书」，'
                              '多数学校要求论文包含学术诚信声明页。'))
    else:
        issues.append(('info', '检测到声明页（原创性声明/版权授权书）。'))

    # ---- 4. 目录（是否能生成目录）----
    heading_styled = stats.get('heading_builtin', 0) + stats.get('heading_custom', 0)
    if heading_styled == 0:
        issues.append(('低', '未检测到使用 Word 标题样式（Heading1/2/3）的段落，'
                              'Word 可能无法自动生成目录。建议使用 headings-fix.py 套标题样式。'))
    else:
        issues.append(('info', f'检测到 {heading_styled} 处使用标题样式的段落，Word 可据此生成目录。'))

    return issues


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    path = sys.argv[1]
    json_out = '--json' in sys.argv
    out = None
    if '-o' in sys.argv:
        out = sys.argv[sys.argv.index('-o') + 1]
    profile = None
    if '--profile' in sys.argv:
        pf = sys.argv[sys.argv.index('--profile') + 1]
        with open(pf, encoding='utf-8') as f:
            profile = json.load(f)
    z, root = docxutils.load(path)
    paras = docxutils.paragraphs(root)
    st = collect_stats(z, root, paras)

    all_issues = []
    all_issues += check_headings(paras)
    # 有模板画像时，标题样式归属由 template_driven_checks 按学校指定样式判定，
    # 不应再用通用"自定义样式告警"误报（按模板套用 toc 样式本就是正确做法）。
    if not profile:
        all_issues += check_heading_styles(st)
    all_issues += check_three_line(st['tables'])
    # 模板驱动模式下，字体/字号/缩进由 template_driven_checks 按批注精确核对，
    # 关闭通用"字体随意混用/字号过多/正文缩进"的启发式检查，避免封面/标题引发的误报。
    if not profile:
        all_issues += check_fonts(st['font_c'])
        all_issues += check_sizes(st['sz_c'])
        all_issues += check_indent(st['body_total'], st['body_no_indent'])
    # 页边距在模板驱动模式下由 template_driven_checks 按批注精确核对，关闭通用近似检查避免重复告警。
    if not profile:
        all_issues += check_margins(st['margins'])
    all_issues += check_punct(paras)
    all_issues += check_references(paras)
    all_issues += check_captions(paras, st['tables'])
    # 结构页只读诊断（封面/摘要/声明/目录）——不修改文件，独立分组展示
    struct_issues = check_structural_pages(paras, st)
    if profile:
        all_issues += template_driven_checks(paras, st, profile)
    all_issues.sort(key=lambda x: SEV_ORDER.get(x[0], 3))

    if json_out:
        obj = {
            'file': path,
            'table_count': len(st['tables']),
            'image_count': st['media'],
            'para_count': st['para_count'],
            'margins': st['margins'],
            'heading_builtin': st['heading_builtin'],
            'heading_custom': st['heading_custom'],
            'font_count': dict(st['font_c']),
            'sz_count': dict(st['sz_c']),
            'template_driven': profile is not None,
            'issues': [{'severity': s, 'msg': m} for s, m in all_issues],
            'issue_count': len(all_issues),
            'severity_counts': dict(Counter(s for s, _ in all_issues)),
            'structural_pages': [{'severity': s, 'msg': m} for s, m in struct_issues],
        }
        if profile:
            obj['spec_comment_count'] = profile.get('comment_count')
            obj['spec_headings'] = {
                l: {'font': (profile.get('headingStyles') or {}).get(l, {}).get('font'),
                    'size': (profile.get('headingStyles') or {}).get(l, {}).get('size'),
                    'outline': (profile.get('headingStyles') or {}).get(l, {}).get('outline_level')}
                for l in ('1', '2', '3')
            }
        print(json.dumps(obj, ensure_ascii=False, indent=2))
        return

    def cm(v):
        try:
            return f'{round(int(v) / 1440 * 2.54, 2)}cm'
        except Exception:
            return str(v)

    L = []
    L.append(f'# 论文格式校验报告：{os.path.basename(path)}')
    L.append('')
    if profile:
        L.append('> ✅ 本报告为【模板驱动诊断】——依据您上传的学校模板**批注**中的权威规范逐项核对')
        L.append(f'> （模板 `{profile.get("source","")}` 共提取批注 {profile.get("comment_count",0)} 条；用户要求以批注写明为准）。')
        L.append('> 以下"与批注不符"项即为需整改点；未列出的维度表示已符合模板要求。')
    else:
        L.append('> ⚠️ 本报告为【通用规范体检】——未提供学校/单位格式模板，仅按学术通用红线检查。')
        L.append('> 学校特有的**精确数值**（如上边距必须 2.5cm、正文必须宋体小四、行距固定 20 磅、指定标题字号）')
        L.append('> 需【上传模板】做"模板驱动诊断"才能逐项核对，否则可能漏报。建议补传模板复检。')
    L.append('')
    L.append(f'- 正文段落：**{st["para_count"]}** ｜ 可选表格：**{len(st["tables"])}** ｜ 图片：**{st["media"]}** ｜ 检出问题：**{len(all_issues)}**')
    if all_issues:
        _sc = Counter(s for s, _ in all_issues)
        L.append(f'- 严重程度分布：🔴高危 **{_sc.get("高", 0)}** · 🟡中危 **{_sc.get("中", 0)}** · 🟢低危 **{_sc.get("低", 0)}**')
    L.append('')
    L.append('## 一、问题清单（按严重度）')
    L.append('')
    if not all_issues:
        L.append('✅ 未发现明显格式红线问题。')
    else:
        for s, m in all_issues:
            L.append(f'- {SEV_ICON[s]} **[{s}]** {m}')
    L.append('')
    L.append('## 二、当前格式画像（统计，供参考/复检）')
    L.append('')
    if st['font_c']:
        top = '、'.join(f'`{k}`×{v}' for k, v in st['font_c'].most_common(8))
        L.append(f'- 中文字体分布：{top}')
    else:
        L.append('- 中文字体：未提取到显式设置（可能继承自样式默认值）。')
    if st['sz_c']:
        top = '、'.join(f'{int(s)//2}pt(sz={s})×{v}' for s, v in st['sz_c'].most_common(8))
        L.append(f'- 字号分布：{top}')
    if st['firstline_c']:
        top = '、'.join(f'firstLineChars={k}×{v}' for k, v in st['firstline_c'].most_common(6))
        L.append(f'- 首行缩进：{top}')
    if st['margins']:
        parts = ' ｜ '.join(f'{k}={cm(v)}' for k, v in st['margins'].items())
        L.append(f'- 页边距：{parts}')
    else:
        L.append('- 页边距：未提取到。')
    L.append(f'- 标题样式：内置 Heading **{st["heading_builtin"]}** 处 / 自定义样式 **{st["heading_custom"]}** 处（影响一键出目录）')
    L.append('')
    # ---- 结构页只读诊断（封面/摘要/声明/目录）----
    L.append('## 三、结构页检查（只读诊断，不自动修改）')
    L.append('')
    for s, m in struct_issues:
        L.append(f'- {SEV_ICON.get(s, "")} **[{s}]** {m}')
    L.append('')
    L.append('> 📌 结构页（封面/摘要/声明/目录）各校差异大，建议对照学校模板手动确认，技能暂不自动修正。')
    L.append('')
    L.append('---')
    L.append('> 报告由 thesis-format-doctor · 格式校验脚本生成（仅诊断不修改文档）。')
    if all_issues:
        L.append('> 💡 以上问题均需整改。存在高危/中危项时，可使用 headings-fix.py 套标题样式出目录、ref-reformat.py 按 GB/T 7714 重排参考文献。')
    else:
        L.append('> ✨ 格式已达标，无需修正。')
    report = '\n'.join(L)
    if out:
        with open(out, 'w', encoding='utf-8') as f:
            f.write(report)
        print(f'已写入 {out}')
    else:
        print(report)


if __name__ == '__main__':
    main()
