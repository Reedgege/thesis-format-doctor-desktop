# -*- coding: utf-8 -*-
"""论文格式医生 · 学生版 · 无界面回归测试

覆盖所有引擎入口（逻辑 bug 真正藏身处），不依赖 tkinter / 网络：
  - engine.normalize_input（.docx 直通；不碰原文件）
  - engine.run_check（无画像 / 带画像）
  - engine.md_to_docx（报告转 Word）
  - engine.run_fix_headings（一键套标题样式 + 修改报告）
  - engine.run_reformat_refs（参考文献 GB/T 7714 重排）
  - engine.run_build_profile（学校模板画像提取）
  - engine.cleanup_conv_dirs（转换临时目录清理）
  - watermark.apply_watermark（试用输出水印）
  - trial / license 激活边界（全部指向临时目录，不碰真实 ~/.tfd_license）

用法：python tests/test_headless_regression.py
退出码 0=全过，1=有失败。

与导师版 tests/test_headless_regression.py 的差异：本仓库没有 group_report /
xlsx_report / run_annotate（导师版专属的全组汇总与只批注模式），故不覆盖。
"""
import os
import sys
import json
import shutil
import tempfile
import traceback
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, ".."))   # tfd-desktop/
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from tfd_app import engine
from tfd_app import watermark
from tfd_app import trial
from tfd_app import license as lic

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
    body = "".join([
        _para("摘要"),
        _para("本文研究毕业论文格式自动修正方法，按学校模板套用，验证引擎稳定性。"),
        _para("关键词"),
        _para("神经网络；模板驱动；格式修正"),
        _para("1 绪论"),
        _para("这是正文内容，用于验证正文区的格式套用与隔离。"),
        _para("1.1 研究背景"),
        _para("近年来，高校对毕业论文格式要求日益严格。"),
        _table(),
        _para("2 文献综述"),
        _para("已有研究从多个角度探讨了格式自动化。"),
        _para("参考文献"),
        _para("[1] 张三. 毕业论文格式规范研究[J]. 高等教育, 2020, 12(3): 45-50."),
        _para("[2] 李四. 自动化排版系统设计[D]. 北京: 某大学, 2021."),
        _para("致谢"),
        _para("感谢导师的悉心指导。"),
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
    """写一份最小但合法的画像（含 levels + headingStyles 规范式），供模板驱动路径使用。"""
    prof = {
        "source": "test_template.docx",
        "levels": {
            "abstract": {"zh_font": "宋体", "en_font": "Times New Roman", "sz": "22",
                         "bold": True, "align": "center"},
            "keywords": {"zh_font": "宋体", "en_font": "Times New Roman", "sz": "22",
                         "bold": True, "align": "left"},
            "ack": {"zh_font": "黑体", "en_font": "Times New Roman", "sz": "22",
                    "bold": True, "align": "center"},
        },
        "headingStyles": {
            "1": {"styleId": "Heading1", "name": "标题 1", "font": "黑体",
                  "size": "小三", "outline_level": 1},
            "2": {"styleId": "Heading2", "name": "标题 2", "font": "黑体",
                  "size": "四号", "outline_level": 2},
            "3": {"styleId": "Heading3", "name": "标题 3", "font": "黑体",
                  "size": "小四", "outline_level": 3},
        },
        "body": {"styleId": "Normal", "font": "宋体", "size": "小四"},
        "page": {},
        "spec": {},
    }
    with open(path, "w", encoding="utf-8") as f:
        json.dump(prof, f, ensure_ascii=False, indent=2)


# ---------------------------------------------------------------- 测试框架
checks = []


def ck(name, ok, msg=""):
    checks.append((name, ok, msg))
    print(("  OK  " if ok else "  XX  ") + name + ((" | " + msg) if msg else ""))


def _is_valid_docx(p):
    if not os.path.isfile(p):
        return False
    try:
        with zipfile.ZipFile(p) as z:
            return "word/document.xml" in z.namelist()
    except Exception:
        return False


def run_all():
    tmp = tempfile.mkdtemp(prefix="tfd_regress_")
    try:
        src = os.path.join(tmp, "thesis.docx")
        profile = os.path.join(tmp, "profile.json")
        build_thesis(src)
        build_profile_json(profile)
        ck("构造测试 docx 合法", _is_valid_docx(src))

        # 0) normalize_input：已是 .docx → 原样直通，不动原文件
        try:
            norm, note = engine.normalize_input(src)
            ck("normalize_input(.docx) 原样直通", norm == src and note is None, "note=%r" % note)
        except Exception:
            ck("normalize_input(.docx) 原样直通", False, traceback.format_exc().splitlines()[-1])

        # 1) 体检：无画像
        try:
            rep = engine.run_check(src)
            ck("run_check(无画像) 不崩且产出报告",
               isinstance(rep, str) and len(rep) > 0, "len=%d" % len(rep))
        except Exception:
            ck("run_check(无画像) 不崩且产出报告", False, traceback.format_exc().splitlines()[-1])

        # 2) 体检：带画像
        try:
            rep2 = engine.run_check(src, profile_path=profile)
            ck("run_check(带画像) 不崩且产出报告",
               isinstance(rep2, str) and len(rep2) > 0, "len=%d" % len(rep2))
        except Exception:
            ck("run_check(带画像) 不崩且产出报告", False, traceback.format_exc().splitlines()[-1])

        # 3) 报告 md -> docx
        rep_docx = os.path.join(tmp, "report.docx")
        try:
            engine.md_to_docx("# 检查报告\n\n- 一条问题\n", rep_docx)
            ck("md_to_docx 产出合法 docx", _is_valid_docx(rep_docx))
        except Exception:
            ck("md_to_docx 产出合法 docx", False, traceback.format_exc().splitlines()[-1])

        # 4) 一键修正（套标题样式），带修改报告
        dst_fix = os.path.join(tmp, "fixed.docx")
        rep_fix = os.path.join(tmp, "fix_report.docx")
        try:
            md = engine.run_fix_headings(src, dst_fix, profile_path=profile,
                                         report_docx=rep_fix, add_comments=False)
            ck("run_fix_headings 不崩且产出 docx",
               _is_valid_docx(dst_fix), "md_len=%d" % len(md or ""))
            ck("run_fix_headings 产出修改报告 docx", _is_valid_docx(rep_fix))
        except Exception:
            ck("run_fix_headings 不崩且产出 docx", False,
               traceback.format_exc().splitlines()[-1])

        # 5) 参考文献重排
        dst_ref = os.path.join(tmp, "refs.docx")
        try:
            rep3 = engine.run_reformat_refs(src, dst_ref)
            ck("run_reformat_refs 不崩且产出 docx",
               _is_valid_docx(dst_ref), "rep_len=%d" % len(rep3 or ""))
        except Exception:
            ck("run_reformat_refs 不崩且产出 docx", False,
               traceback.format_exc().splitlines()[-1])

        # 6) 提取学校模板画像
        prof_out = os.path.join(tmp, "extracted.json")
        try:
            txt = engine.run_build_profile(src, prof_out)
            ck("run_build_profile 不崩且产出 json",
               os.path.isfile(prof_out) and len(txt) > 0)
        except Exception:
            ck("run_build_profile 不崩且产出 json", False,
               traceback.format_exc().splitlines()[-1])

        # 7) 转换临时目录清理（幂等、不抛）
        try:
            engine.cleanup_conv_dirs()
            ck("cleanup_conv_dirs 幂等不抛", True)
        except Exception:
            ck("cleanup_conv_dirs 幂等不抛", False, traceback.format_exc().splitlines()[-1])

        # 8) 水印（作用在副本上，不破坏原文件）
        wm_src = os.path.join(tmp, "wm_src.docx")
        shutil.copy(src, wm_src)
        try:
            watermark.apply_watermark(wm_src)
            ck("watermark.apply_watermark 不崩且产出合法 docx", _is_valid_docx(wm_src))
        except Exception:
            ck("watermark.apply_watermark 不崩且产出合法 docx", False,
               traceback.format_exc().splitlines()[-1])

        # 9) trial / license 边界（指向临时目录，不影响真实授权）
        lic.LICENSE_DIR = tmp
        lic.LICENSE_FILE = os.path.join(tmp, "license.json")
        trial.TRIAL_FILE = os.path.join(tmp, "trial.json")
        # 隔离中台试用登记：单测不联网、不污染线上 trials 表（TRIAL_LIMIT=1）
        lic.server_trial_used = lambda mc: False
        lic.claim_server_trial = lambda mc: None
        trial._SERVER_USED.update(val=False, ts=1e18)   # 缓存置为「未用过」，避免联网
        try:
            assert not trial.is_licensed(), "测试环境不应已激活"
            ck("trial 初始剩余=1", trial.trials_left() == 1, "left=%d" % trial.trials_left())
            ck("trial 第1次扣减成功", trial.consume_trial() is True)
            ck("trial 用完第2次被拒", trial.consume_trial() is False)
            # 篡改锁定
            with open(trial.TRIAL_FILE, encoding="utf-8") as fh:
                d = json.load(fh)
            d["used"] = 0
            with open(trial.TRIAL_FILE, "w", encoding="utf-8") as fh:
                json.dump(d, fh, ensure_ascii=False)
            ck("trial 篡改后锁定(剩余0)", trial.trials_left() == 0)
            # 正式版绕开
            lic.save_local_license("TEST-CARD", lic.get_machine_code(), permanent=True)
            ck("激活后 is_licensed=True", trial.is_licensed() is True)
            ck("激活后扣减不消耗", trial.consume_trial() is True)
            summary = lic.license_summary()
            ck("license_summary 返回卡种", isinstance(summary, dict) and summary.get("kind"),
               "kind=%r" % (summary or {}).get("kind"))
        except Exception:
            ck("trial/license 边界", False, traceback.format_exc().splitlines()[-1])

    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    all_ok = all(ok for _, ok, _ in checks)
    print("\n==== 学生版无界面回归: %s ====" % ("全部通过" if all_ok else "存在失败"))
    for name, ok, msg in checks:
        if not ok:
            print("  失败: %s (%s)" % (name, msg))
    return all_ok


if __name__ == "__main__":
    sys.exit(0 if run_all() else 1)
