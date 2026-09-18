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

# 试用总次数（可调：卖家用）。2026-09-12 老板定：2 次免费试用，第 3 次起弹购买引导。
TRIAL_LIMIT = 2

TRIAL_FILE = os.path.join(LICENSE_DIR, "trial.json")

# 签名盐：固定盐值，防普通用户直接改 json 里的 used（盐非机密，仅挡零成本手改）。
# 注：v1.3.96 起弃用 base64+XOR 混淆存储（杀软 ML 误报元凶，见 license.py 注释）。
# 盐值非机密（客户端内可见），改明文常量。
_TRIAL_SALT = b"tfd|trial|count|2026|v1"


def _sign(machine, used):
    return hmac.new(_TRIAL_SALT, ("%s|%d" % (machine, used)).encode("utf-8"),
                    hashlib.sha256).hexdigest()[:16]


# 中台试用登记状态（带 TTL 缓存，避免 UI 刷新频繁联网；离线/异常一律按未用，绝不阻断）。
_SERVER_USED = {"val": False, "ts": 0.0}
_SERVER_TTL = 120.0


def _server_used(network: bool = True):
    """中台是否记得本机已用过试用。带缓存 + 离线兜底。

    作用：堵住「删掉 trial.json 即可重置试用」的白嫖口子——本地可删，中台记录删不掉。

    ``network=False``：只读缓存、**绝不联网**。供 UI 徽标路径使用：``trials_left()`` 在
    App 构造与每次刷新时都会被调用，若在此联网（3s 超时）会冻住主线程。真正裁决在
    ``consume_trial`` 里联网查。
    """
    now = time.time()
    if now - _SERVER_USED["ts"] < _SERVER_TTL:
        return _SERVER_USED["val"]
    if not network:
        return False                     # 未联网过 → 交回本地判定（徽标路径不阻塞 UI）
    try:
        from . import license as _lic
        val = bool(_lic.server_trial_used(get_machine_code()))
    except Exception:
        val = False                      # 离线/接口异常 → 交回本地判定
    _SERVER_USED.update(val=val, ts=now)
    return val


def _claim_server():
    """向中台登记本机已用试用（幂等；失败静默）。

    **只有登记成功**才把缓存标为「中台已知」；离线失败保持原值、等下次联网再试
    （否则失败也写 True，会把同一会话里的后续试用误拒）。
    """
    try:
        from . import license as _lic
        if _lic.claim_server_trial(get_machine_code()):
            _SERVER_USED.update(val=True, ts=time.time())
    except Exception:
        pass


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
    # 手改 trial.json 成 null/[]/"x"/数字 等非 dict 结构时，data.get 会抛 AttributeError
    # 且不被上面的 except 捕获，会一路冒泡到启动时 _update_trial_badge → 程序起不来。
    if not isinstance(data, dict):
        return "BAD"
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
    if _server_used(network=False):
        return 0                         # 中台已登记（读缓存）→ 本地被删/重置也不给
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
    # 中台优先：本机已在中台登记过试用 → 拒绝（防「删 trial.json 重置」白嫖）
    if _server_used(network=True):
        return False
    used = (data or {}).get("used", 0)
    if used >= TRIAL_LIMIT:
        # 本地已用尽：补一次登记（覆盖「首次试用恰在离线时登记失败」的情形）。
        # 注意：这里堵不住「先删掉 trial.json 再联网」—— 那时 _load() 返回 None、used=0，
        # 根本走不到该分支，中台也无记录。属「离线优先」设计的固有残余风险（已在文档标注）。
        _claim_server()
        return False
    if used == 0:
        # 首次试用：**先登记**（在线时立即生效）再扣本地次数 —— 这样用户随后删掉
        # trial.json 也无法重置（中台已记得本机用过）。
        _claim_server()
    used += 1
    try:
        os.makedirs(LICENSE_DIR, exist_ok=True)
        with open(TRIAL_FILE, "w", encoding="utf-8") as f:
            json.dump({"machine": machine, "used": used,
                       "sig": _sign(machine, used)}, f, ensure_ascii=False)
    except Exception:
        # v1.3.121：计次文件写盘失败（权限/磁盘/路径异常）→ 扣减失败，按"已用完"拦截，
        # 与"扣减失败也拦"的计费边界一致（Codex 审查 P2-9）。
        return False
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
