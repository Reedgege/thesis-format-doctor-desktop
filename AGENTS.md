# AGENTS.md — tfd-desktop（论文格式医生 · 桌面版 / 学生版）

> 这是给 Codex / AI 编码助手的**上下文文件**。
> 老板（芦苇）不懂代码 —— 你改完必须交 **diff + 大白话说明**，经「技术专家团」代码把关（fan-zhengma）后才算完成，最终由老板真机验收。

## 一、项目是什么

论文格式医生桌面端（学生版），源自扣子(Coze) skill `thesis-format-doctor`，改造成**本机离线运行的原生桌面程序**。

- **当前版本**：`VERSION` 文件 = `1.3.110`（已 Nuitka 编译，VirusTotal 0 报毒）
- 国际版名：**DocxFmt**；官网 reedskill.com

核心功能：格式体检（通用规范 / 模板驱动）、套标题出目录、参考文献按 GB/T 7714 重排、抽取学校模板画像。

## 二、技术栈与硬约束（⚠️ 红线）

- **Python ≥ 3.10，零第三方依赖**：docx 处理靠 `zipfile` + XML 手写完成。
  🔴 **禁止引入任何第三方库**（pip 装包）——这会摧毁"打包极小、跨平台几乎无依赖坑"的核心优势。需要新能力，用标准库实现。
- **界面**：原生 **Tkinter**，无浏览器、无本地服务、纯本地运行。
- **打包**：**Nuitka**（`build.py`），Windows 用 NSIS（`installer.nsi`）出安装包。
- 🔴 **防杀软红线**：**禁止自写 base64 / XOR 等加密混淆**。历史实锤：v1.3.88–v1.3.95 因 `_obf()` 被 Windows Defender 判 `Sabsik.TE.A!ml`，去混淆后 v1.3.96 解决。真要加固，走**代码签名证书**或 **PyArmor**，不要手搓混淆。
- **版本格式**：`VERSION` 为纯数字（`1.3.110`），改版时需同步 `VERSION`、`pyproject.toml` 的 `version`、UI 显示、官网下载页四处。

## 三、目录与入口

| 路径 | 说明 |
|---|---|
| `main.py` | 程序入口 |
| `tfd_app/gui.py` | Tkinter 界面逻辑 |
| `tfd_app/engine.py` | 引擎编排（体检 / 套标题 / 重排文献 / 画像） |
| `tfd_app/core/` | 核心处理模块（docx 读写、规则、报告） |
| `tfd_app/license.py` `trial.py` `watermark.py` | 激活授权、试用次数控制、水印预览 |
| `tfd_app/assets/` | 内置资源与图标 |
| `tools/make_portable_assets.py` | 生成便携版资源 |

## 四、常用命令

```bash
python main.py                        # 本地运行调试
python tools/make_portable_assets.py  # 生成便携资源
python build.py                       # 本地打包（正式三平台走 GitHub Actions）
```

## 五、不可破坏的对外承诺（源自 MEMORY.md，违反 = L3 拦截）

1. **论文绝对不上传**：全部处理在本机完成。联网**仅**用于「在线激活校验」与「试用额度登记（只送机器码，绝不含论文内容）」，其余时间完全离线。
2. **免费试用 1 次**：机器码绑定 + 中台登记（防「删本地计数文件」重置），用完弹小程序码购买引导；试用输出带水印预览版。
3. **模板驱动**：学校模板画像优先级 **批注说明 > 样式定义 > 通用规范**；无模板时退通用规范并明确提示。
4. **一次付费永久使用**：终身 + 一机一码设备绑定。
5. **每处改动加 Word 批注气泡**，交付物含：改后 docx + 修改报告 + 检查报告。

## 六、改动铁律（每次迭代必守）

1. **小步快跑**：一次只改一件事。
2. **先备份**：改前 `git commit` 或打包留存，翻车可回滚。
3. **先自审**：改完自己先 review 一遍 diff。
4. **白话交底**：diff + 大白话说明（改了啥、动了哪块、影响啥、怎么验）一起交。
5. **先测后上**：先在测试环境验证，再动生产。
6. **真机验收**：老板用**真实论文**走完整流程（加载→检查→修正→看批注气泡），他说 OK 才算上线。

## 七、本仓库驱动 Codex 的标准命令

```bash
export DEEPSEEK_API_KEY="$(cat '/d/AgentSpace/config/deepseek-key.txt' | tr -d '\r\n')" && \
codex exec --skip-git-repo-check "你要它干的事"                        # 只读分析/问代码
codex exec --skip-git-repo-check -s workspace-write "你要它干的事"    # 真改代码：必须加 -s workspace-write
```

⚠️ **沙箱坑**：Codex 默认沙箱为 `read-only`，**不写盘——所有 shell 命令会被策略拦截**（报 `exec_command failed ... rejected: blocked by policy`）。
**要让它真正落盘改代码，必须显式加 `-s workspace-write`**；只做只读问答时用默认即可。取值：`read-only | workspace-write | danger-full-access`（后者不要碰）。

在本目录运行，改动即落本仓库。API key 走本地 keyfile，**不进聊天、不写进 `~/.codex/config.toml`**。

## 八、事实来源（冲突以此为准）

- `D:\AgentSpace\MEMORY.md` —— 全局事实卡：功能 / 版本 / 价格 / 联系方式的唯一来源，**严禁擅自改动对外口径**
- `D:\AgentSpace\技术专家团_接手须知.md` —— 技术专家团 briefing
- `D:\AgentSpace\迭代与需求清单.md` —— 老板的大白话需求入口
