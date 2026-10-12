# -*- coding: utf-8 -*-
"""本地回环 HTTP 服务（对接层 v0.1）

职责：托管 web/ 前端 + 提供 /api/health。后续引擎接口按
docs/web-ui-v1/后端接口需求_v1.md 逐个接入。

硬约束：
  1. 纯标准库（http.server），零第三方依赖——这是本产品的硬底线，不引 Flask/FastAPI
  2. 只监听 127.0.0.1，论文不出本机
  3. 随机端口，避免多开冲突
  4. 关窗即退出：用 --user-data-dir 强制浏览器起独立进程，Popen.wait() 反映真实关窗
"""
import json
import os
import socket
import subprocess
import sys
import tempfile
import threading
import time
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer

# 导师版仓库此常量改为 "mentor"
EDITION = "student"
EDITION_LABEL = "学生版"
# 导师版为 True（run_annotate / modify=False / 批量）
SUPPORTS_ANNOTATE = False
SUPPORTS_BATCH = False

WEB_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "web")


def read_version():
    """读仓库 VERSION 文件，读不到给 unknown"""
    here = os.path.dirname(os.path.abspath(__file__))
    for cand in (os.path.join(here, "..", "..", "VERSION"),
                 os.path.join(here, "..", "VERSION")):
        cand = os.path.normpath(cand)
        if os.path.isfile(cand):
            try:
                return open(cand, encoding="utf-8").read().strip().lstrip("v")
            except Exception:
                pass
    return "unknown"


def health_payload():
    modes = ["check", "fix", "profile"]
    if SUPPORTS_ANNOTATE:
        modes.append("annotate")
    return {
        "ok": True,
        "data": {
            "app": "thesis-format-doctor",
            "version": read_version(),
            "edition": EDITION,
            "edition_label": EDITION_LABEL,
            "lang_default": "zh-CN",
            "engine_ready": True,
            "capabilities": {
                "modes": modes,
                "template_sources": ["docx"],
                "specs": [],
                "refs_reformat": False,
                "add_comments": True,
                "batch": SUPPORTS_BATCH,
                "watermark_on_trial": True,
                "purchase_block": True,
                "max_file_mb": 200,
            },
        },
    }


class Handler(SimpleHTTPRequestHandler):
    def __init__(self, *a, **kw):
        super().__init__(*a, directory=WEB_DIR, **kw)

    def do_GET(self):
        if self.path.split("?")[0] == "/api/health":
            body = json.dumps(health_payload(), ensure_ascii=False).encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            self.wfile.write(body)
            return
        # 根路径按 edition 重定向，前端据此决定皮肤（学生版青蓝 / 导师版朱红）
        if self.path in ("/", ""):
            self.path = "/index.html?edition=%s" % ("advisor" if EDITION == "mentor" else "student")
        return super().do_GET()

    def log_message(self, fmt, *args):
        pass  # 静默，不刷屏


def pick_port():
    s = socket.socket()
    s.bind(("127.0.0.1", 0))
    port = s.getsockname()[1]
    s.close()
    return port


def find_browser():
    cands = [
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    ]
    for p in cands:
        if os.path.isfile(p):
            return p
    return None


def main(argv=None):
    argv = sys.argv[1:] if argv is None else argv
    no_browser = "--no-browser" in argv

    port = pick_port()
    httpd = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    threading.Thread(target=httpd.serve_forever, daemon=True).start()

    url = "http://127.0.0.1:%d/?edition=%s" % (port, "advisor" if EDITION == "mentor" else "student")
    print("serving %s on %s" % (WEB_DIR, url), flush=True)

    if no_browser:
        print(url, flush=True)
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            pass
        httpd.shutdown()
        return 0

    browser = find_browser()
    if not browser:
        print("未找到 Edge/Chrome，请手动打开：%s" % url, flush=True)
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            pass
        httpd.shutdown()
        return 1

    profile = tempfile.mkdtemp(prefix="tfd-web-")
    proc = subprocess.Popen([
        browser,
        "--app=%s" % url,
        "--user-data-dir=%s" % profile,          # 独立 profile → 独立进程，关窗可检测
        "--window-size=1040,700",
        "--disable-features=Translate",
        "--no-first-run",
        "--no-default-browser-check",
    ])
    proc.wait()   # 关窗即退出，不留僵尸进程
    httpd.shutdown()
    return 0


if __name__ == "__main__":
    sys.exit(main())
