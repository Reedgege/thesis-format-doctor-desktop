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
    万一中台不可用，客户联系卖家，卖家用统一发码器 reedcode_unified.py 按机器码生成离线码，
    客户粘贴即可激活。离线码用 Ed25519 非对称签名（私钥只在卖家本机），公开仓库无法伪造。

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
import threading
import urllib.parse
import urllib.request
import subprocess
import base64

# 离线备用码验签：纯标准库 Ed25519（RFC 8032），零第三方加密依赖。
# 公钥 32 字节（base64）编译进客户端，只能验签；私钥只在卖家本机 _signing_keys/，
# 由统一发码器 reedcode_unified.py 持有，绝不入库/不进安装包。详见 ed25519_verify.py。
from .ed25519_verify import verify as _ed25519_verify

_STUDENT_PUBLIC_KEY_B64 = "YjPfUcmn2vuKz0mxHwf6SlBPPXl5wHGCiXi8v8sIcb4="
_STUDENT_PUBLIC_KEY = base64.b64decode(_STUDENT_PUBLIC_KEY_B64)

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
CONSUME_PATH = "/api/consume"

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


# 离线备用码：验签用 Ed25519 公钥（见文件顶部 _STUDENT_PUBLIC_KEY）。
# 签名只由卖家本机统一发码器完成，私钥绝不进客户端/仓库。


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
# 试用额度登记（中台，按机器绑定）—— 与海外版同一套 /api/trial 契约
# ---------------------------------------------------------------------------
TRIAL_PATH = "/api/trial"


def _trial_sync(machine_code, claim=False):
    """试用额度登记/查询（中台）。网络失败返回 None，由调用方走本地兜底。"""
    resp, err = _http_json(TRIAL_PATH, {
        "product": DEFAULT_PRODUCT, "machine_code": machine_code,
        "claim": bool(claim)}, 3)
    if err:
        return None
    return resp


def server_trial_used(machine_code):
    """中台是否记得该机器已用过试用（查不到/离线 → False，绝不阻断本地判定）。

    必须显式要求 ok=True 才信 trial_used —— `_http_json` 对非 2xx 也会解析响应体，
    若错误信封里恰好带 trial_used 字段，会被误判成「已用过」而误锁（Codex 2026-09-12）。
    """
    r = _trial_sync(machine_code, claim=False)
    return bool(r and r.get("ok") and r.get("trial_used"))


def claim_server_trial(machine_code) -> bool:
    """首次试用后向中台登记（幂等）。离线/失败静默忽略，不影响本地计次。

    返回是否**登记成功** —— 调用方据此决定要不要把「中台已知」写进本地缓存：
    否则离线失败也会被当成已登记，把同一会话里的后续试用误拒（Codex 2026-09-12 N2）。
    """
    try:
        r = _trial_sync(machine_code, claim=True)
        return bool(r and r.get("ok"))
    except Exception:
        return False


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
        "uses_total": resp.get("uses_total"),
        "uses_used": resp.get("uses_used"),
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


def _mark_invalid_local(reason="revoked"):
    """把本地授权标记为失效（下次 check_local_valid 直接判 False）。
    reason: revoked（撤销/退款）/ expired（到期）/ uses_exhausted（次卡用尽）。"""
    lic = load_local_license()
    if not lic:
        return
    lic["status"] = "revoked"
    lic["lock_reason"] = reason
    try:
        os.makedirs(LICENSE_DIR, exist_ok=True)
        with open(LICENSE_FILE, "w", encoding="utf-8") as f:
            json.dump(lic, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def _mark_revoked_local():
    """兼容旧调用：标记为已撤销。"""
    _mark_invalid_local("revoked")


def get_lock_reason():
    """返回本地失效原因（revoked/expired/uses_exhausted），无则 None。"""
    lic = load_local_license()
    if not lic:
        return None
    return lic.get("lock_reason")


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


def _classify_expired(resp):
    """区分失效原因：次卡用尽 → uses_exhausted；否则 expired。"""
    if resp.get("uses_total") is not None and (resp.get("uses_used") or 0) >= resp.get("uses_total"):
        return "uses_exhausted"
    return "expired"


def _update_uses_local(total, used):
    """次卡：回写本地剩余次数（供界面显示"剩 N 次"）。"""
    lic = load_local_license()
    if not lic:
        return
    lic["uses_total"] = total
    lic["uses_used"] = used
    try:
        with open(LICENSE_FILE, "w", encoding="utf-8") as f:
            json.dump(lic, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def validate_now(product=None, timeout=4):
    """启动时实时校验（只读 /api/heartbeat，不扣次数、不绑机）。返回：
        'active' / 'revoked' / 'expired' / 'offline'
    revoked/expired 会落地本地失效标记；offline（断网/服务器挂）不锁，按"离线不防"放行。"""
    lic = load_local_license()
    if not lic or not lic.get("code"):
        return "offline"
    card = lic.get("code")
    mc = get_machine_code()
    product = product or lic.get("product") or DEFAULT_PRODUCT
    resp, err = _http_json(
        HEARTBEAT_PATH,
        {"product": product, "code": card, "machine_code": mc},
        timeout,
    )
    if err:
        return "offline"
    if not resp or not resp.get("ok"):
        return "offline"
    st = resp.get("status")
    if st == "revoked":
        _mark_invalid_local("revoked")
        return "revoked"
    if st == "expired":
        _mark_invalid_local(_classify_expired(resp))
        return "expired"
    if resp.get("expires_at") is not None:
        _update_expire_local(resp.get("expires_at"))
    return "active"


def consume_use(product=None, timeout=5):
    """付费操作前：联网实时校验 + 次卡扣 1 次（/api/consume）。返回：
        'active'（放行）/ 'revoked' / 'expired' / 'offline'（离线不拦、不扣）
    次卡用尽返回 'expired' 并落地失效；非次卡只校验不扣次数。"""
    lic = load_local_license()
    if not lic or not lic.get("code"):
        return "offline"
    card = lic.get("code")
    mc = get_machine_code()
    product = product or lic.get("product") or DEFAULT_PRODUCT
    resp, err = _http_json(
        CONSUME_PATH,
        {"product": product, "code": card, "machine_code": mc},
        timeout,
    )
    if err:
        return "offline"
    if not resp or not resp.get("ok"):
        return "offline"
    st = resp.get("status")
    if st == "revoked":
        _mark_invalid_local("revoked")
        return "revoked"
    if st == "expired":
        _mark_invalid_local(_classify_expired(resp))
        return "expired"
    if resp.get("expires_at") is not None:
        _update_expire_local(resp.get("expires_at"))
    if resp.get("uses_total") is not None:
        _update_uses_local(resp.get("uses_total"), resp.get("uses_used"))
    return "active"


def save_local_license(card, machine_code, permanent=True, lic_type=None,
                       expires_at=None, product=None, uses_total=None, uses_used=None):
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
        "uses_total": uses_total,
        "uses_used": uses_used,
    }
    with open(LICENSE_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    return LICENSE_FILE


# ---------------------------------------------------------------------------
# 兜底方案：离线备用码（Ed25519 非对称验签，公钥见文件顶部）
# ---------------------------------------------------------------------------
def verify_offline_code(code, machine_code):
    """校验离线备用码（Ed25519）：① 内嵌公钥验签通过；② 绑定本机机器码。

    对任意畸形 / 伪造输入都返回 False，绝不抛异常（GUI 回调里被调用，
    无控制台下异常会被静默吞掉，用户只会看到「点了没反应」）。
    """
    if not isinstance(code, str) or not isinstance(machine_code, str):
        return False
    code = code.strip()
    if not code or "." not in code:
        return False
    try:
        payload, _, sig_b64 = code.rpartition(".")
        if not payload.startswith(machine_code + "|"):
            return False
        # 宽松解码：容忍缺失的 '=' 填充与标准 base64 的 +/ 字符（紧凑格式常省略 padding），
        # 避免「发码器未补 padding → 整批离线码静默失效」。
        sig_s = sig_b64.strip().replace("+", "-").replace("/", "_")
        sig_s += "=" * (-len(sig_s) % 4)
        sig = base64.urlsafe_b64decode(sig_s.encode("ascii"))
        msg = payload.encode("utf-8")
        return _ed25519_verify(_STUDENT_PUBLIC_KEY, msg, sig)
    except Exception:
        return False


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


def license_summary():
    """本地授权的可读摘要（供 UI 展示，两个版本共用）。

    返回 dict：
      kind   — 卡种中文名：永久版 / 周卡 / 月卡 / 次卡 / 正式版（未知类型时兜底）
      desc   — 状态描述：永久有效 / YYYY-MM-DD 到期（剩 N 天） / 剩余 X / Y 次
      badge  — 右上角徽章短文案，如「正式版 · 周卡（剩 5 天）」
    未授权时返回 None。
    """
    lic = load_local_license()
    if not lic or lic.get("status") == "revoked":
        return None

    type_map = {"lifetime": "永久版", "weekly": "周卡", "monthly": "月卡",
                "times": "次卡"}
    t = (lic.get("type") or "").lower()
    kind = type_map.get(t)
    if kind is None:
        kind = "永久版" if lic.get("permanent") else "正式版"

    desc = "永久有效"
    if t == "times":
        total = lic.get("uses_total")
        used = lic.get("uses_used") or 0
        left = (total - used) if isinstance(total, int) else None
        desc = ("剩余 %d / %d 次" % (left, total)) if left is not None else "按次计费"
    else:
        exp = lic.get("expire") or lic.get("expires_at")
        days_left = None
        if exp:
            try:
                import datetime as _dt
                ms = int(exp)
                if ms > 1e12:  # 秒级时间戳（历史数据）补成毫秒
                    pass
                d = _dt.datetime.fromtimestamp(ms / 1000)
                days_left = max(0, int((ms - time.time() * 1000) // 86400000))
                desc = "%s 到期（剩 %d 天）" % (d.strftime("%Y-%m-%d"), days_left)
            except Exception:
                desc = "有效"

    # v1.3.100：徽章收紧为短文案——v1.3.99 用"正式版 · 周卡（剩 N 天）"在
    # trial_btn 没设 width 的情况下，会撑爆右栏（实测撑到 w=788px/h=403px，
    # 把 btn_row 挤到顶部、时间线挤压错位）。改为 ≤8 字符的紧凑形式。
    # 全角按 2、半角按 1 计，最坏"次卡 3/10"=6 字；"月卡 28天"=6 字。
    badge = "正式版 ✓"
    if t == "times":
        total = lic.get("uses_total")
        used = lic.get("uses_used") or 0
        left = (total - used) if isinstance(total, int) else None
        if left is not None:
            badge = "%s %d/%d ✓" % (kind, left, total)
        else:
            badge = "%s ✓" % kind
    elif t in ("weekly", "monthly") and days_left is not None:
        badge = "%s %d天 ✓" % (kind, days_left)
    elif kind and t not in ("lifetime",) and t != "永久版":
        badge = "%s ✓" % kind
    return {"kind": kind, "desc": desc, "badge": badge, "type": t}


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
