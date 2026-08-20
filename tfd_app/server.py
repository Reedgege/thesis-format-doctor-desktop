#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
论文格式医生 · 桌面版后端服务
================================
纯 Python 标准库实现：内置 HTTP 服务 + 引擎封装。
双击可执行文件后，自动打开浏览器即可使用，完全离线、零第三方依赖。

引擎来自扣子(Coze) skill「thesis-format-doctor v1.3.5」，
已重命名为可 import 的模块（format_checker / headings_fix / ref_reformat / format_profile）。
"""
import os
import sys
import io
import json
import uuid
import base64
import tempfile
import threading
import webbrowser
import contextlib
import importlib
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse

HERE = os.path.dirname(os.path.abspath(__file__))
CORE = os.path.join(HERE, "core")
# 冻结后用 _MEIPASS 作为资源根目录
BASE = getattr(sys, "_MEIPASS", HERE)
FRONTEND = os.path.join(BASE, "frontend")
ASSETS = os.path.join(BASE, "assets")

if CORE not in sys.path:
    sys.path.insert(0, CORE)

# ----------------------------------------------------------------------------
# 预热导入引擎模块（尽早暴露导入期错误）
# ----------------------------------------------------------------------------
ENGINE = {}

def _load_engine():
    for name in ("docxutils", "format_checker", "headings_fix",
                 "ref_reformat", "format_profile", "report_docx"):
        try:
            ENGINE[name] = importlib.import_module(name)
        except Exception as e:  # pragma: no cover
            print(f"[warn] 引擎模块 {name} 导入失败: {e}", file=sys.stderr)

_load_engine()

# 文件暂存：fid -> 目录
STORE = {}
STORE_LOCK = threading.Lock()


# ----------------------------------------------------------------------------
# 引擎调用封装
# ----------------------------------------------------------------------------
def run_engine(module_name, argv, capture=True):
    """设置 sys.argv 并调用引擎 main()，捕获 stdout。"""
    mod = ENGINE.get(module_name)
    if mod is None:
        return {"ok": False, "error": f"引擎模块未加载: {module_name}"}
    old = sys.argv
    sys.argv = [module_name + ".py"] + list(argv)
    out = io.StringIO()
    try:
        if capture:
            with contextlib.redirect_stdout(out):
                mod.main()
        else:
            mod.main()
        return {"ok": True, "stdout": out.getvalue()}
    except SystemExit as e:
        return {"ok": True, "stdout": out.getvalue(), "exit_code": e.code}
    except Exception as e:
        return {"ok": False, "error": f"{type(e).__name__}: {e}",
                "stdout": out.getvalue()}
    finally:
        sys.argv = old


# ----------------------------------------------------------------------------
# 请求辅助
# ----------------------------------------------------------------------------
def _new_fid_dir():
    fid = uuid.uuid4().hex[:12]
    d = os.path.join(WORK_ROOT, fid)
    os.makedirs(d, exist_ok=True)
    with STORE_LOCK:
        STORE[fid] = d
    return fid, d


def _save_b64(b64, name, dstdir):
    data = base64.b64decode(b64)
    path = os.path.join(dstdir, name)
    with open(path, "wb") as f:
        f.write(data)
    return path


def _json_resp(handler, obj, status=200):
    body = json.dumps(obj, ensure_ascii=False).encode("utf-8")
    handler.send_response(status)
    handler.send_header("Content-Type", "application/json; charset=utf-8")
    handler.send_header("Content-Length", str(len(body)))
    handler.send_header("Access-Control-Allow-Origin", "*")
    handler.end_headers()
    handler.wfile.write(body)


def _read_json_body(handler):
    length = int(handler.headers.get("Content-Length", 0))
    raw = handler.rfile.read(length) if length else b"{}"
    return json.loads(raw.decode("utf-8"))


# ----------------------------------------------------------------------------
# 业务接口
# ----------------------------------------------------------------------------
def api_check(payload):
    fid, d = _new_fid_dir()
    doc = payload.get("file") or {}
    if not doc.get("data"):
        return {"ok": False, "error": "缺少待检测论文文件"}
    _save_b64(doc["data"], "input.docx", d)
    inp = os.path.join(d, "input.docx")
    argv = [inp, "--json"]
    prof = payload.get("profile")
    if prof and prof.get("data"):
        _save_b64(prof["data"], "profile.json", d)
        argv += ["--profile", os.path.join(d, "profile.json")]
    argv += ["-o", os.path.join(d, "report.md")]
    res = run_engine("format_checker", argv)
    if not res.get("ok"):
        return {"ok": False, "error": res.get("error"), "detail": res.get("stdout")}
    try:
        result = json.loads(res["stdout"])
    except Exception:
        return {"ok": False, "error": "引擎输出解析失败", "detail": res.get("stdout")}
    # 引擎在 --json 模式下不会落盘 md，这里据 JSON 另生成一份报告
    md = _render_report_md(result)
    with open(os.path.join(d, "report.md"), "w", encoding="utf-8") as f:
        f.write(md)
    return {"ok": True, "fid": fid, "result": result,
            "report": f"/files/{fid}/report.md"}


def _render_report_md(r):
    sev = r.get("severity_counts") or {}
    L = []
    L.append(f'# 论文格式校验报告：{os.path.basename(r.get("file", ""))}')
    L.append("")
    if r.get("template_driven"):
        L.append("> ✅ 本报告为【模板驱动诊断】——依据上传模板批注中的权威规范逐项核对。")
    else:
        L.append("> ⚠️ 本报告为【通用规范体检】——未提供学校模板，仅按学术通用红线检查。")
    L.append("")
    L.append(f'- 正文段落：**{r.get("para_count", "-")}** ｜ 表格：**{r.get("table_count", "-")}** ｜ 图片：**{r.get("image_count", "-")}** ｜ 检出问题：**{r.get("issue_count", len(r.get("issues", [])))}**')
    sc = sev
    if sc:
        L.append(f'- 严重程度分布：🔴高危 **{sc.get("高",0)}** · 🟡中危 **{sc.get("中",0)}** · 🟢低危 **{sc.get("低",0)}**')
    L.append("")
    L.append("## 一、问题清单（按严重度）")
    L.append("")
    for it in r.get("issues", []):
        L.append(f'- **[{it.get("severity","")}]** {it.get("msg","")}')
    L.append("")
    L.append("## 二、结构页检查（只读，建议对照模板手动确认）")
    L.append("")
    for it in r.get("structural_pages", []):
        L.append(f'- **[{it.get("severity","")}]** {it.get("msg","")}')
    L.append("")
    L.append("---")
    L.append("> 报告由 论文格式医生·桌面版 生成（仅诊断不修改文档）。")
    return "\n".join(L)


def api_fix(payload):
    fid, d = _new_fid_dir()
    doc = payload.get("file") or {}
    if not doc.get("data"):
        return {"ok": False, "error": "缺少待处理论文文件"}
    _save_b64(doc["data"], "input.docx", d)
    inp = os.path.join(d, "input.docx")
    out = os.path.join(d, "fixed.docx")
    argv = [inp, out]
    prof = payload.get("profile")
    if prof and prof.get("data"):
        _save_b64(prof["data"], "profile.json", d)
        argv += ["--profile", os.path.join(d, "profile.json")]
    if payload.get("no_comments"):
        argv += ["--no-comments"]
    if payload.get("make_report", True):
        argv += ["--report", os.path.join(d, "change_report.docx")]
    res = run_engine("headings_fix", argv)
    if not res.get("ok"):
        return {"ok": False, "error": res.get("error"), "detail": res.get("stdout")}
    files = {"fixed": f"/files/{fid}/fixed.docx"}
    if os.path.exists(os.path.join(d, "change_report.docx")):
        files["report"] = f"/files/{fid}/change_report.docx"
    return {"ok": True, "fid": fid, "summary": res["stdout"], "files": files}


def api_reformat(payload):
    fid, d = _new_fid_dir()
    doc = payload.get("file") or {}
    if not doc.get("data"):
        return {"ok": False, "error": "缺少待处理论文文件"}
    _save_b64(doc["data"], "input.docx", d)
    argv = [os.path.join(d, "input.docx"), "-o", os.path.join(d, "fixed.docx")]
    res = run_engine("ref_reformat", argv)
    if not res.get("ok"):
        return {"ok": False, "error": res.get("error"), "detail": res.get("stdout")}
    return {"ok": True, "fid": fid, "summary": res["stdout"],
            "files": {"fixed": f"/files/{fid}/fixed.docx"}}


def api_profile(payload):
    fid, d = _new_fid_dir()
    doc = payload.get("file") or {}
    if not doc.get("data"):
        return {"ok": False, "error": "缺少学校模板文件"}
    _save_b64(doc["data"], "template.docx", d)
    argv = [os.path.join(d, "template.docx"), "-o", os.path.join(d, "profile.json")]
    res = run_engine("format_profile", argv)
    if not res.get("ok"):
        return {"ok": False, "error": res.get("error"), "detail": res.get("stdout")}
    prof_path = os.path.join(d, "profile.json")
    profile = None
    if os.path.exists(prof_path):
        try:
            with open(prof_path, encoding="utf-8") as f:
                profile = json.load(f)
        except Exception:
            profile = None
    return {"ok": True, "fid": fid, "summary": res["stdout"],
            "profile": profile,
            "files": {"profile": f"/files/{fid}/profile.json"}}


# ----------------------------------------------------------------------------
# HTTP 处理
# ----------------------------------------------------------------------------
class Handler(BaseHTTPRequestHandler):
    def log_message(self, *args):  # 静默默认日志
        pass

    def _serve_file(self, fpath, content_type):
        if not os.path.exists(fpath):
            self.send_error(404, "File not found")
            return
        with open(fpath, "rb") as f:
            data = f.read()
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Access-Control-Allow-Origin", "*")
        if content_type.endswith("pdf") or fpath.endswith((".docx", ".json", ".md")):
            self.send_header("Content-Disposition",
                             f'attachment; filename="{os.path.basename(fpath)}"')
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self):
        u = urlparse(self.path)
        path = u.path
        if path in ("/", "/index.html"):
            self._serve_file(os.path.join(FRONTEND, "index.html"),
                             "text/html; charset=utf-8")
            return
        if path.startswith("/static/"):
            name = os.path.basename(path)
            fpath = os.path.join(FRONTEND, name)
            if name.endswith(".js"):
                ct = "application/javascript; charset=utf-8"
            elif name.endswith(".css"):
                ct = "text/css; charset=utf-8"
            else:
                ct = "application/octet-stream"
            self._serve_file(fpath, ct)
            return
        if path.startswith("/files/"):
            # /files/<fid>/<name>
            parts = path.split("/")
            if len(parts) >= 4:
                fid, name = parts[2], parts[3]
                with STORE_LOCK:
                    d = STORE.get(fid)
                if d:
                    fpath = os.path.join(d, name)
                    ct = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    if name.endswith(".json"):
                        ct = "application/json"
                    elif name.endswith(".md"):
                        ct = "text/markdown; charset=utf-8"
                    self._serve_file(fpath, ct)
                    return
            self.send_error(404, "File not found")
            return
        self.send_error(404, "File not found")

    def do_POST(self):
        u = urlparse(self.path)
        route = u.path
        try:
            payload = _read_json_body(self)
        except Exception as e:
            _json_resp(self, {"ok": False, "error": f"请求解析失败: {e}"}, 400)
            return
        try:
            if route == "/api/check":
                _json_resp(self, api_check(payload))
            elif route == "/api/fix":
                _json_resp(self, api_fix(payload))
            elif route == "/api/reformat":
                _json_resp(self, api_reformat(payload))
            elif route == "/api/profile":
                _json_resp(self, api_profile(payload))
            else:
                _json_resp(self, {"ok": False, "error": "未知接口"}, 404)
        except Exception as e:
            _json_resp(self, {"ok": False, "error": f"服务异常: {e}"}, 500)

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "POST, GET, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()


WORK_ROOT = tempfile.mkdtemp(prefix="tfd_")


def run(host="127.0.0.1", port=8765, open_browser=True):
    srv = ThreadingHTTPServer((host, port), Handler)
    url = f"http://{host}:{port}/"
    print(f"论文格式医生已启动：{url}")
    if open_browser:
        try:
            webbrowser.open(url)
        except Exception:
            pass
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        srv.shutdown()


if __name__ == "__main__":
    run()
