# -*- coding: utf-8 -*-
"""论文格式医生 · 桌面版 — 启动入口

双击生成的 exe / app / 可执行文件后，本脚本会：
  1. 启动内置本地 Web 服务（纯标准库，离线运行）
  2. 自动打开默认浏览器进入操作界面
文件全程只在你本机处理，不会上传任何服务器。
"""
import os
import sys

# 确保能找到 tfd_app 包与本文件同目录
ROOT = os.path.dirname(os.path.abspath(__file__))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from tfd_app import server


def main():
    host = os.environ.get("TFD_HOST", "127.0.0.1")
    try:
        port = int(os.environ.get("TFD_PORT", "8765"))
    except ValueError:
        port = 8765
    server.run(host=host, port=port, open_browser=True)


if __name__ == "__main__":
    main()
