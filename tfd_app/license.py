# -*- coding: utf-8 -*-
"""
授权模块 · 论文格式医生
==================================================================
双轨制：

  主方案（在线）—— 自建授权中台（api.reedskill.com，Cloudflare Workers + D1）
    客户输入激活码 → 软件联网把"激活码 + 本机机器码"发到中台校验
    → 通过则本机写入授权文件 → 之后日常使用完全离线（不联网、不心跳）
    → 中台支持"一机一码"设备绑定；退款后后台点"撤销"，下次联网心跳即锁死
    → 心跳默认 3 天一次，仅"联网 + 服务端标记 revoked"才锁，离线/超时一律不锁

  兜底方案（离线）—— 卖家离线发码
    万一中台不可用，客户联系卖家，卖家用 offlinecode 工具按机器码生成离线码，
    客户粘贴即可激活。离线码用对称签名，安全等级适中，足以防小白共享。

机器码：硬件级（磁盘序列号 + 主板 UUID），重装系统不变，换电脑不同。

本地授权文件：~/.tfd_license/license.json，日常启动只读它，不联网。
"""
import os
import sys
import json
import time
import uuid
import socket
import hashlib
import hmac
import threading
import urllib.parse
import urllib.request
import subprocess

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
# 自建授权中台配置
# ---------------------------------------------------------------------------
API_BASE = "https://api.reedskill.com"     # 自建中台域名（Cloudflare Workers + D1）
DEFAULT_PRODUCT = "tfd-student"            # 学生版；导师版（tfd-mentor）单独打包时改此处
HEARTBEAT_INTERVAL_DAYS = 3                # 心跳间隔（天）；仅联网 + 服务端 revoked 才锁
ACTIVATE_PATH = "/api/activate"
HEARTBEAT_PATH = "/api/heartbeat"

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
# 离线备用码密钥（与 offlinecode.py 共用）。固定签名盐，用于离线激活兜底。
# 注：v1.3.96 起弃用 base64+XOR 混淆存储——该"解密循环"形状会被杀软 ML 引擎
# 误判为恶意载荷解密器（Trojan:Win32/Sabsik.TE.A!ml 误报元凶）。盐值本身非机密
# （客户端内必然可见），改为明文常量以消除误报特征。防逆向靠 PyArmor 加壳策略。
# ---------------------------------------------------------------------------
_OFFLINE_KEY = b"tfd|kami|offline|2026|sign|v1"


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
# 网络底层：直连（绕过系统代理）+ 超时兜底（防线程永久悬挂）
# ---------------------------------------------------------------------------
def _http_json(path, payload, timeout):
    """POST JSON，返回 (resp_dict_or_None, error_or_None)。网络错误返回 error。"""
    data = json.dumps(payload).encode("utf-8")
    url = API_BASE + path
    prev = socket.getdefaulttimeout()
    socket.setdefaulttimeout(timeout)   # 兜底：覆盖 DNS / SSL 握手阶段的悬挂
    try:
        req = urllib.request.Request(
            url, data=data,
            headers={"Content-Type": "application/json", "User-Agent": "tfd-desktop/1.0"},
            method="POST",
        )
        # 无代理直连（无视环境变量 HTTP(S)_PROXY / 系统代理设置）
        opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
        with opener.open(req, timeout=timeout) as resp:
            return json.loads(resp.read().decode("utf-8", "ignore")), None
    except urllib.error.HTTPError as e:
        try:
            return json.loads(e.read().decode("utf-8", "ignore")), None
        except Exception:
            return None, "HTTP %s" % e.code
    except Exception as e:
        return None, str(e)
    finally:
        socket.setdefaulttimeout(prev)


# ---------------------------------------------------------------------------
# 主方案：自建中台在线激活
# ---------------------------------------------------------------------------
def activate_online(card, machine_code, product=DEFAULT_PRODUCT, timeout=30):
    """联网激活：向中台校验激活码，返回 dict：
        {"ok": bool, "type": "lifetime"/"weekly"|None, "expires_at": int|None, "error": str}

    网络错误返回 ok=False（调用方可改走离线码）；激活码无效/已用/已撤/过期给出中文提示。
    """
    card = (card or "").strip()
    if not card:
        return {"ok": False, "error": "激活码为空"}
    resp, err = _http_json(
        ACTIVATE_PATH,
        {"product": product, "code": card, "machine_code": machine_code},
        timeout,
    )
    if err:
        return {"ok": False, "error": "网络验证失败（%s）；若多次失败，请联系客服获取离线激活码" % err}
    if not resp or not resp.get("ok"):
        code_err = (resp or {}).get("error")
        msg = {
            "invalid_code": "激活码无效，请核对后重试",
            "already_used": "该激活码已绑定其他设备，无法在本机使用（一机一码）",
            "revoked": "该激活码已被撤销（可能已退款），无法激活",
            "expired": "该激活码已过期，请联系客服或重新购买",
            "missing_fields": "请求参数缺失，请联系客服",
        }.get(code_err, "激活未通过：%s" % code_err)
        return {"ok": False, "error": msg}
    return {
        "ok": True,
        "type": resp.get("type"),
        "expires_at": resp.get("expires_at"),
        "error": "验证通过",
    }


def verify_via_kami(card, machine_code, timeout=30, retries=1, product=DEFAULT_PRODUCT):
    """兼容旧调用：返回 (ok: bool, days: int, msg: str)。

    days：周卡返回剩余天数，买断版/终身返回 0。
    """
    info = activate_online(card, machine_code, product=product, timeout=timeout)
    days = 0
    if info.get("type") == "weekly" and info.get("expires_at"):
        days = max(0, int((info["expires_at"] - int(time.time() * 1000)) // 86400000))
    return info.get("ok", False), days, (info.get("error") or "验证通过")


# ---------------------------------------------------------------------------
# 心跳：退款锁死检测（3 天一次，仅联网 + 服务端 revoked 才锁）
# ---------------------------------------------------------------------------
_heartbeat_running = False


def start_heartbeat(product=DEFAULT_PRODUCT, on_revoked=None, interval_days=HEARTBEAT_INTERVAL_DAYS):
    """启动后台心跳（守护线程，不阻塞 GUI）。重复调用只起一个。

    on_revoked: 回调（后台线程调用，内部应切回主线程处理 UI），无返回值要求。
    """
    global _heartbeat_running
    if _heartbeat_running:
        return
    lic = load_local_license()
    if not lic or not lic.get("code"):
        return
    _heartbeat_running = True

    def loop():
        try:
            while True:
                time.sleep(interval_days * 86400)
                _do_one_heartbeat(lic, on_revoked)
        except Exception:
            pass

    t = threading.Thread(target=loop, daemon=True)
    t.start()


def _do_one_heartbeat(lic, on_revoked):
    """单次心跳：仅 revoked 才锁；离线/超时/网络错一律忽略（不锁、不计时）。"""
    card = lic.get("code")
    mc = get_machine_code()
    product = lic.get("product") or DEFAULT_PRODUCT
    resp, err = _http_json(
        HEARTBEAT_PATH,
        {"product": product, "code": card, "machine_code": mc},
        10,
    )
    if err:
        return  # 离线/超时 -> 不锁、不计时
    if not resp or not resp.get("ok"):
        return
    st = resp.get("status")
    if st == "revoked":
        _mark_revoked_local()
        if callable(on_revoked):
            try:
                on_revoked()
            except Exception:
                pass
    # active / expired 都更新本地过期时间（周卡到期软提示续费，不硬锁）
    if resp.get("expires_at") is not None:
        _update_expire_local(resp.get("expires_at"))


def _mark_revoked_local():
    """把本地授权标记为已撤销（下次启动 check_local_valid 直接判失效）。"""
    lic = load_local_license()
    if not lic:
        return
    lic["status"] = "revoked"
    try:
        os.makedirs(LICENSE_DIR, exist_ok=True)
        with open(LICENSE_FILE, "w", encoding="utf-8") as f:
            json.dump(lic, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def _update_expire_local(exp):
    lic = load_local_license()
    if not lic:
        return
    lic["expire"] = exp
    try:
        with open(LICENSE_FILE, "w", encoding="utf-8") as f:
            json.dump(lic, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def is_revoked():
    """本地是否已标记撤销（心跳检测到 revoked 后落地）。"""
    lic = load_local_license()
    if not lic:
        return False
    return lic.get("status") == "revoked"


def save_local_license(card, machine_code, permanent=True, lic_type=None,
                       expires_at=None, product=None):
    """激活成功后写本地授权（日常离线用）。"""
    os.makedirs(LICENSE_DIR, exist_ok=True)
    data = {
        "machine_code": machine_code,
        "code": (card or "").strip(),
        "issued": int(time.time()),
        "permanent": permanent,
        "expire": expires_at if expires_at is not None else None,
        "offline": False,
        "type": lic_type,
        "product": product,
        "status": "active",
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
        "type": "lifetime",
        "product": DEFAULT_PRODUCT,
        "status": "active",
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
    """启动时使用：本地授权存在、机器码匹配、未过期、未撤销 → 放行（不联网）。"""
    lic = load_local_license()
    if not lic:
        return False
    if lic.get("status") == "revoked":
        return False
    if lic.get("machine_code") != get_machine_code():
        return False  # 授权文件被拷到别的电脑 → 失效
    if lic.get("permanent"):
        return True
    exp = lic.get("expire")
    if exp and int(time.time() * 1000) > exp:
        return False
    return True


if __name__ == "__main__":
    # 直接运行可查看本机机器码（调试用）
    print("本机机器码：", get_machine_code())
    print("本地授权有效：", check_local_valid())
