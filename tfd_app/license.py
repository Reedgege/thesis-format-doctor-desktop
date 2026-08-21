# -*- coding: utf-8 -*-
"""
授权模块 · 论文格式医生
==================================================================
双轨制：

  主方案（在线）—— 卡密通（keyt.cn）
    客户输入卡密 → 软件联网把"卡密 + 本机机器码"发给卡密通验证
    → 通过则本机写入授权文件 → 之后日常使用完全离线（不联网、不心跳）
    → 卡密通后台支持"一机一码"设备绑定，分享给他人用不了

  兜底方案（离线）—— 卖家离线发码
    万一卡密通跑路/宕机，客户联系卖家，卖家用自己的发码工具
    （offlinecode.py）按客户机器码生成"离线备用码"，客户粘贴即可激活。
    离线备用码用对称签名，安全等级适中，足以防"小白共享"，
    且需主动联系卖家才给，价值低、破解动力小。

机器码：取磁盘序列号 + 主板 UUID（硬件级，重装系统不变，换电脑不同），
        做一次 SHA256 取前 32 位大写，作为本机指纹。

本地授权文件：存于用户主目录 ~/.tfd_license/license.json，日常启动只读它，
             不联网，保证"论文处理全程离线"的卖点。
"""
import os
import sys
import json
import time
import uuid
import socket
import hashlib
import hmac
import subprocess
import urllib.parse
import urllib.request

# Windows 下子进程（wmic/powershell 等控制台程序）默认会弹一个黑框一闪；
# 在 --noconsole 打包的 GUI 程序里必须加 CREATE_NO_WINDOW，让子进程静默执行。
_NO_WINDOW = (getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000)
              if sys.platform.startswith("win") else 0)


def _run(cmd):
    """静默执行子进程命令并返回 stdout 文本（不弹黑框、不闪控制台）。"""
    kwargs = {"stderr": subprocess.DEVNULL, "timeout": 8}
    if sys.platform.startswith("win"):
        kwargs["creationflags"] = _NO_WINDOW
    return subprocess.check_output(cmd, **kwargs).decode("utf-8", "ignore")

# ---------------------------------------------------------------------------
# 卡密通配置 —— 卖家在 keyt.cn 注册开发者、创建应用后，把下面的常量改成你自己的
# ---------------------------------------------------------------------------
KAMI_USER = "reedai0537"      # 卡密通用户名（专属验证地址：keyt.cn/kami/reedai0537/check.php）
KAMI_APP = "lunwengeshi"       # 应用名（后台创建的应用 lunwengeshi）
KAMI_CHECK_URL = "https://www.keyt.cn/kami/{user}/check.php".format(user=KAMI_USER)

# ---------------------------------------------------------------------------
# 本地授权文件位置
# ---------------------------------------------------------------------------
LICENSE_DIR = os.path.join(os.path.expanduser("~"), ".tfd_license")
LICENSE_FILE = os.path.join(LICENSE_DIR, "license.json")

# ---------------------------------------------------------------------------
# 离线备用码密钥（与 offlinecode.py 共用；做了简单分片混淆，提高提取门槛）
# ---------------------------------------------------------------------------
_OFFLINE_KEY = _obf("OykrMyQuIiYzICkpIyYhKjN9f315MzwmKCEzOX4=").encode("utf-8")


def get_machine_code():
    """计算本机指纹（硬件级，尽量稳定）。失败有兜底。"""
    parts = []

    # 磁盘序列号（优先 wmic，失败退 PowerShell）
    try:
        out = _run(["wmic", "diskdrive", "get", "serialnumber"])
        for line in out.splitlines():
            line = line.strip()
            if line and "SerialNumber" not in line:
                parts.append("disk:" + line)
                break
    except Exception:
        pass
    if not parts:
        try:
            out = _run(["powershell", "-NoProfile", "-Command",
                        "(Get-CimInstance Win32_DiskDrive).SerialNumber"]).strip()
            if out:
                parts.append("disk:" + out)
        except Exception:
            pass

    # 主板 / BIOS UUID
    try:
        out = _run(["wmic", "csproduct", "get", "uuid"])
        for line in out.splitlines():
            line = line.strip()
            if line and "UUID" not in line:
                parts.append("uuid:" + line)
                break
    except Exception:
        pass

    # 兜底：主机名 + MAC（换硬件时也会变，仅作最后保险）
    if not parts:
        try:
            parts.append("host:" + socket.gethostname())
            parts.append("mac:" + "%012X" % uuid.getnode())
        except Exception:
            parts.append("rand:" + uuid.uuid4().hex)

    raw = "|".join(parts)
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()[:32].upper()


# ---------------------------------------------------------------------------
# 主方案：卡密通在线验证
# ---------------------------------------------------------------------------
def verify_via_kami(card, machine_code, timeout=40, retries=2):
    """联网校验卡密。返回 (ok: bool, days: int, msg: str)。

    卡密通 check.php 返回约定：以 'ok|' 开头表示通过（其后可能带 天数|分钟）。
    其他内容均视为失败（含 'invalid' / 'expired' / 'bind' 等）。

    注意：
    - 卡密通免费平台实测服务端响应 20 秒+（连接很快、首字节慢），超时放宽到 40s；
    - 强制直连：卡密通是国内服务器，客户电脑的 VPN/代理（环境变量）会把请求
      绕道海外，出现"后台显示在线、软件却一直等不到响应"；直连最快最稳；
    - 超时直接给可操作提示（关 VPN/代理 或 联系卖家要离线码），不让客户干等。
    """
    if KAMI_USER in ("你的卡密通用户名",):
        return False, 0, "卡密通尚未配置：请在 tfd_app/license.py 填写 KAMI_USER 与 KAMI_APP"
    params = urllib.parse.urlencode({"card": card, "mac": machine_code, "app": KAMI_APP})
    url = KAMI_CHECK_URL + "?" + params
    last_err = ""
    for attempt in range(retries):
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "tfd-desktop/1.0"})
            # 无代理直连（无视环境变量 HTTP(S)_PROXY / 系统代理设置）
            opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
            with opener.open(req, timeout=timeout) as resp:
                text = resp.read().decode("utf-8", "ignore").strip()
            if text.startswith("ok|"):
                return True, 0, "验证通过"
            return False, 0, "验证未通过：%s" % text
        except (TimeoutError, socket.timeout):
            return False, 0, ("验证超时：请关闭电脑的 VPN/代理后重试；"
                              "或联系卖家获取离线激活码，无需等待在线验证")
        except Exception as e:
            last_err = str(e)
            if attempt < retries - 1:
                time.sleep(1.0)
    return False, 0, "网络验证失败（请检查网络或稍后重试）：%s" % last_err


def save_local_license(card, machine_code, permanent=True):
    """激活成功后写本地授权（日常离线用）。"""
    os.makedirs(LICENSE_DIR, exist_ok=True)
    data = {
        "machine_code": machine_code,
        "card": hashlib.sha256(card.encode("utf-8")).hexdigest()[:16],
        "issued": int(time.time()),
        "permanent": permanent,
        "expire": None,
        "offline": False,
    }
    with open(LICENSE_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    return LICENSE_FILE


# ---------------------------------------------------------------------------
# 兜底方案：离线备用码
# ---------------------------------------------------------------------------
def _offline_sign(payload):
    return hmac.new(_OFFLINE_KEY, payload.encode("utf-8"), hashlib.sha256).hexdigest()[:24]


def generate_offline_code(machine_code):
    """卖家发码工具用：machine_code + 时间戳 + 签名 → 离线码。"""
    ts = int(time.time())
    payload = "%s|%d" % (machine_code, ts)
    return payload + "|" + _offline_sign(payload)


def verify_offline_code(code, machine_code):
    """校验离线备用码：签名正确 且 绑定本机机器码。"""
    if not code or "|" not in code:
        return False
    payload, _, sig = code.rpartition("|")
    if not payload.startswith(machine_code + "|"):
        return False
    return hmac.compare_digest(_offline_sign(payload), sig)


def save_offline_license(code, machine_code):
    os.makedirs(LICENSE_DIR, exist_ok=True)
    data = {
        "machine_code": machine_code,
        "offline_code": code,
        "issued": int(time.time()),
        "permanent": True,
        "offline": True,
    }
    with open(LICENSE_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


# ---------------------------------------------------------------------------
# 日常：本地授权校验（离线）
# ---------------------------------------------------------------------------
def load_local_license():
    if not os.path.isfile(LICENSE_FILE):
        return None
    try:
        with open(LICENSE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


def check_local_valid():
    """启动时使用：本地授权存在、机器码匹配、未过期 → 放行（不联网）。"""
    lic = load_local_license()
    if not lic:
        return False
    if lic.get("machine_code") != get_machine_code():
        return False  # 授权文件被拷到别的电脑 → 失效
    if lic.get("permanent"):
        return True
    exp = lic.get("expire")
    if exp and int(time.time()) > exp:
        return False
    return True


if __name__ == "__main__":
    # 直接运行可查看本机机器码（调试用）
    print("本机机器码：", get_machine_code())
    print("本地授权有效：", check_local_valid())
