# -*- coding: utf-8 -*-
"""
卖家发码工具（兜底用）· 论文格式医生
==================================================
仅在「卡密通跑路 / 宕机」时给客户重激活使用，不打包进软件 exe。

用法（在卖家电脑上，需有 Python 3）：
    python keygen.py
按提示输入客户的机器码（让客户在软件激活页"复制机器码"发给你），
回车即生成离线备用码，发给客户粘贴激活。自动记入 keygen_log.txt 台账。

安全：离线码用对称签名（密钥在 license.py 与本项目共用）。破解者即便提取
密钥，也需知道目标机器码才能造码，且需主动联系你才给——价值低、动力小。
"""
import os
import sys
import time

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

from tfd_app import license as lic


def main():
    print("=" * 52)
    print("  论文格式医生 · 离线发码工具（兜底用）")
    print("=" * 52)
    print("本机机器码（卖家）：", lic.get_machine_code())
    print()
    print("让客户在『激活页』点击『复制机器码』，把机器码发给你，")
    print("然后粘贴到下面。")
    print("-" * 52)
    target = input("请输入客户的机器码：").strip()
    if not target:
        print("未输入，已退出。")
        return
    code = lic.generate_offline_code(target)
    print()
    print("生成的离线备用码（发给客户）：")
    print("  " + code)
    print()
    log_path = os.path.join(HERE, "keygen_log.txt")
    with open(log_path, "a", encoding="utf-8") as f:
        f.write("%s | 客户机器码 %s | 离线码 %s\n" % (
            time.strftime("%Y-%m-%d %H:%M"), target, code))
    print("（已记入台账 keygen_log.txt）")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        pass
