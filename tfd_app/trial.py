# -*- coding: utf-8 -*-
"""
试用模块 · 论文格式医生
==================================================================
未激活（无激活码）时，客户可试用「一键修正」N 次：
  - 次数绑定机器码：计数文件记录 machine_code，换机/重装后计数不认，
    避免"卸载重装无限试用"；
  - 简单防篡改：used 用 机器码+固定密钥 算 HMAC 签名，普通用户改数字无效；
  - 有激活码（license.check_local_valid() 通过）→ 正式版，不走试用。

试用输出：修正后的 docx 带水印（页眉页脚 + 正文穿插），见 watermark.py。
"""
import os
import time
import json
import hmac
import hashlib

from .license import LICENSE_DIR, get_machine_code  # 相对导入：与 gui 同款，打包后可用

# 试用总次数（可调：卖家用）
TRIAL_LIMIT = 2

TRIAL_FILE = os.path.join(LICENSE_DIR, "trial.json")

# 签名盐：与 license._OFFLINE_KEY 同风格，防普通用户直接改 json 里的 used。
# 原始值经混淆存储、运行时还原，反编译只看到乱码。
def _obf(b):
    """反混淆：base64(xor 0x4F)。还原敏感串，挡小白一把梭提取。"""
    import base64 as _b64
    return bytes((c ^ 0x4F) for c in _b64.b64decode(b)).decode("utf-8")


_TRIAL_SALT = _obf("OykrMzs9Ji4jMywgOiE7M31/fXkzOX4=").encode("utf-8")


def _sign(machine, used):
    return hmac.new(_TRIAL_SALT, ("%s|%d" % (machine, used)).encode("utf-8"),
                    hashlib.sha256).hexdigest()[:16]


def _load():
    """读取计数文件。

    返回：
      None          无文件 / 机器码不匹配（换机/重装）→ 重新计；
      "BAD"         签名不符或结构异常（文件被手改）→ 视为已用完（锁定）；
      正常 dict     有效的计数记录。
    """
    if not os.path.isfile(TRIAL_FILE):
        return None
    try:
        with open(TRIAL_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        return "BAD"                     # 文件损坏：不冒险重新计
    machine = get_machine_code()
    if data.get("machine") != machine:
        return None                      # 换机 / 重装：重新计（新机器）
    used = data.get("used", 0)
    if not isinstance(used, int) or used < 0:
        return "BAD"                     # 结构异常：锁定
    if data.get("sig") != _sign(machine, used):
        return "BAD"                     # 文件被手改：锁定（否则可无限重置试用）
    return data


def trials_left():
    """剩余试用次数（0~TRIAL_LIMIT）。"""
    if is_licensed():
        return TRIAL_LIMIT               # 正式版不消耗试用（语义：无限制）
    data = _load()
    if data == "BAD":
        return 0                         # 篡改/损坏：不给试用
    used = (data or {}).get("used", 0)
    return max(0, TRIAL_LIMIT - used)


def consume_trial():
    """扣减一次试用。成功返回 True；已用完/篡改返回 False。"""
    if is_licensed():
        return True
    machine = get_machine_code()
    data = _load()
    if data == "BAD":
        return False
    used = (data or {}).get("used", 0)
    if used >= TRIAL_LIMIT:
        return False
    used += 1
    try:
        os.makedirs(LICENSE_DIR, exist_ok=True)
        with open(TRIAL_FILE, "w", encoding="utf-8") as f:
            json.dump({"machine": machine, "used": used,
                       "sig": _sign(machine, used)}, f, ensure_ascii=False)
    except Exception:
        pass
    return True


def is_licensed():
    """是否正式版（有本地授权且机器码匹配）。"""
    try:
        return bool(check_local_valid())
    except Exception:
        return False


# 延迟导入，避免 trial 与 license 循环依赖
def check_local_valid():
    from . import license as _lic
    return _lic.check_local_valid()


if __name__ == "__main__":
    print("机器码:", get_machine_code())
    print("剩余试用:", trials_left(), "/", TRIAL_LIMIT)
    print("正式版:", is_licensed())
