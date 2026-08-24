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
    （keygen.py）按客户机器码生成"离线备用码"，客户粘贴即可激活。
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
import concurrent.futures

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


def _log(msg):
    """把激活过程写进日志文件 ~/.tfd_license/activation.log，
    客户机器上激活出问题时，可凭此日志精准定位。任何失败都不影响主流程。"""
    try:
        os.makedirs(LICENSE_DIR, exist_ok=True)
        with open(os.path.join(LICENSE_DIR, "activation.log"), "a", encoding="utf-8") as f:
            f.write("[%s] %s\n" % (time.strftime("%Y-%m-%d %H:%M:%S"), msg))
    except Exception:
        pass

# ---------------------------------------------------------------------------
# 离线备用码密钥（与 keygen.py 共用；做了简单分片混淆，提高提取门槛）
# ---------------------------------------------------------------------------
_OFFLINE_KEY = "|".join(("tfd", "kami", "offline", "2026", "sign", "v1")).encode("utf-8")


_MC_CACHE = None  # v1.3.84：进程内缓存——机器码运行期不变，避免每次调用重复跑子进程


def get_machine_code():
    """计算本机指纹（硬件级，尽量稳定）。失败有兜底。

    v1.3.84：结果进程内缓存（wmic/powershell 子进程每次约 0.5~1.5s，
    GUI 启动/激活/徽章/每次修正多处调用，缓存后不再重复开销）。
    """
    global _MC_CACHE
    if _MC_CACHE:
        return _MC_CACHE
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
    _MC_CACHE = hashlib.sha256(raw.encode("utf-8")).hexdigest()[:32].upper()
    return _MC_CACHE


# ---------------------------------------------------------------------------
# 主方案：卡密通在线验证流程
# ---------------------------------------------------------------------------
def _kami_http(card, machine_code, timeout):
    """单次 HTTP 校验，返回原始响应文本；任何异常都向上抛。

    关键点：urllib 的 opener.open(timeout=...) 在部分环境下（开着 VPN/系统代理、
    SSL 握手阶段、DNS 解析）根本不生效，导致线程永久挂在等回包。这里额外用
    socket.setdefaulttimeout 兜底，让本次 socket 也受超时约束。
    """
    params = urllib.parse.urlencode({"card": card, "mac": machine_code, "app": KAMI_APP})
    url = KAMI_CHECK_URL + "?" + params
    prev = socket.getdefaulttimeout()
    socket.setdefaulttimeout(timeout)   # 兜底：覆盖 DNS / SSL 握手阶段的悬挂
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "tfd-desktop/1.0"})
        # 无代理直连（无视环境变量 HTTP(S)_PROXY / 系统代理设置）
        opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
        with opener.open(req, timeout=timeout) as resp:
            return resp.read().decode("utf-8", "ignore").strip()
    finally:
        socket.setdefaulttimeout(prev)


def verify_via_kami(card, machine_code, timeout=30, retries=2):
    """联网校验卡密。返回 (ok: bool, days: int, msg: str)。

    卡密通 check.php 返回约定：以 'ok|' 开头表示通过（其后可能带 天数|分钟）。
    其他内容均视为失败（含 'invalid' / 'expired' / 'bind' 等）。

    设计要点（针对"后台在线、前端卡死"的真实故障）：
    1. 硬性总超时：用独立线程跑 HTTP，并设 overall 截止时间。即便 urllib 的
       socket 超时失效，线程也会被强制判超时，保证界面一定在限定时间内拿到
       结果，绝不永久卡在"联网激活中"。
    2. 一机一码恢复：卡密通是"首次验证成功即绑定本机"。若首次请求已在服务端
       绑定成功（后台显示"在线"）但客户端没读到回包→超时，此时复用同一张卡
       + 同一机器码再查一次，卡密通会返回 ok|，把授权补回来。因此超时后重试
       同一卡密是正确恢复手段，而不是放弃。
    3. error| 是确定性失败（卡密无效/已用/绑定到别的设备），无需重试，直接告知。
    """
    if KAMI_USER in ("你的卡密通用户名",):
        return False, 0, "卡密通尚未配置：请在 tfd_app/license.py 填写 KAMI_USER 与 KAMI_APP"

    overall = timeout + 12   # 硬上限：即便 socket 超时失效，也不让界面永久卡住
    last_err = ""
    _log("验证开始 card=%s mac=%s timeout=%s retries=%s" % (card, machine_code, timeout, retries))
    for attempt in range(max(1, retries)):
        ex = concurrent.futures.ThreadPoolExecutor(max_workers=1)
        try:
            _log("第 %s 次请求发出（overall=%ss）" % (attempt + 1, overall))
            fut = ex.submit(_kami_http, card, machine_code, timeout)
            text = fut.result(timeout=overall)
            _log("第 %s 次请求返回: %r" % (attempt + 1, text[:120]))
        except (concurrent.futures.TimeoutError, TimeoutError, socket.timeout):
            # 超时：可能是卡密通响应慢，也可能是本机 VPN/代理把请求绕路海外。
            # 一机一码下，同卡+同机重查大概率能拿到 ok|，故重试而非直接放弃。
            last_err = "连接/读取超时（卡密通响应慢，或被 VPN/代理绕路）"
            _log("第 %s 次超时" % (attempt + 1))
            if attempt < retries - 1:
                time.sleep(1.5)
                continue
            return False, 0, ("验证超时：请检查网络后重试；若多次失败，请联系客服获取离线激活码。")
        except Exception as e:
            last_err = str(e)
            _log("第 %s 次异常: %s" % (attempt + 1, last_err))
            if attempt < retries - 1:
                time.sleep(1.5)
                continue
            return False, 0, "网络验证失败（请检查网络或稍后重试）：%s" % last_err
        finally:
            ex.shutdown(wait=False)   # 放弃线程池；仍在跑的孤儿线程会自行结束

        # 拿到了服务端响应
        if text.startswith("ok|"):
            _log("验证通过")
            return True, 0, "验证通过"
        # error| 是确定性失败，无需重试
        _log("验证未通过: %r" % text[:120])
        return False, 0, "验证未通过：%s" % text

    _log("最终失败: %s" % last_err)
    return False, 0, "网络验证失败：%s" % last_err


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
