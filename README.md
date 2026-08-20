# 论文格式医生 · 桌面版 (Thesis Format Doctor — Desktop)

把你用**扣子(Coze)** 做的论文格式处理 skill（`thesis-format-doctor`）变成一个**真正能在本机跑的桌面程序**：

- 🪟 **Windows**：双击 `thesis-format-doctor-desktop.exe`
- 🍎 **macOS**：双击 `thesis-format-doctor-desktop`
- 🐧 **Linux**：运行 `thesis-format-doctor-desktop`

**原生图形窗口（Tkinter），不依赖浏览器、完全离线、零第三方依赖、文件不离开你的电脑**。

> 兼容老格式：检测到 `.doc` / WPS `.wps` 时，若本机装有 LibreOffice 会自动转成 `.docx` 再处理；否则提示你先用 Word/WPS「另存为 .docx」。

---

## 它能做什么

源自 skill 的纯 Python 标准库引擎（已重命名为可 import 的模块）：

| 功能 | 说明 |
| --- | --- |
| 🩺 **格式体检** | 对照通用学术红线（标题样式、三线表、参考文献、页边距、全角标点等）逐项诊断，按高危/中危/低危分级，自动保存 Markdown 报告 |
| 🎯 **套标题 & 出目录** | 自动识别标题层级并套用 Heading 样式，让 Word/WPS 可一键出目录；生成逐条修改明细 docx |
| 📚 **重排参考文献** | 按 **GB/T 7714** 范式重排参考文献区 |
| 📐 **抽取模板画像** | 上传学校模板，自动抽取其格式规范（以模板批注为准），用于更精确的"模板驱动诊断/套用" |

> 体检默认是「通用规范体检」；上传学校模板画像(.json)后，会切换为更贴合要求的「模板驱动诊断」。

---

## 使用方法

1. 下载对应系统的安装包（见 [Releases](../../releases)）。
2. 双击运行（Windows 可能弹出"未知发布者"提示，点"仍要运行"即可——这是没花钱买代码签名证书的正常现象）。
3. 程序弹出原生图形窗口：
   - 选择「目标文档 .docx / .doc / .wps」（必选）
   - 可选：选择「格式画像 .json」用于精确诊断 / 套用（由"从模板生成画像"功能产出）
   - 选处理模式，点「开始处理」；进度与报告显示在下方日志区，结果文件自动保存在原文档旁边
4. 关闭窗口即退出。

---

## 技术说明

- 界面：原生 **Tkinter** 窗口，无浏览器、无本地服务、纯本地运行。
- 引擎：直接复用 skill 中的 `docxutils / format_checker / headings_fix / ref_reformat / format_profile / report_docx`，**零第三方依赖**，因此打包极小、跨平台几乎无依赖坑。
- 打包：PyInstaller `--onefile --noconsole`，见 `build.py` 与 `.github/workflows/build.yml`。

### 本地自行构建（可选）

```bash
pip install pyinstaller
python build.py            # 按当前系统构建到 dist/
```

> Linux 构建前需 `sudo apt-get install -y tk tcl`；Windows / macOS 自带 Tk。

### 发版（GitHub Actions 自动构建三平台）

推送一个 `v*` 标签即可触发：

```bash
git tag v1.3.6
git push origin v1.3.6
```

Actions 会在 Windows / macOS / Ubuntu 三个 runner 上分别构建，并把产物发布到 GitHub Release（按平台后缀 `-windows` / `-macos` / `-linux` 区分）。

---

## 目录结构

```
thesis-format-doctor-desktop/
├── main.py                 # 启动入口（检查 Tkinter → 拉起 GUI）
├── build.py                # 跨平台 PyInstaller 构建脚本
├── pyproject.toml
├── tfd_app/
│   ├── gui.py              # 原生 Tkinter 界面
│   ├── engine.py           # 引擎封装层（吃路径、返回文本 + .doc/.wps 自动转换）
│   ├── core/               # skill 引擎（标准库，零依赖）
│   └── assets/
└── .github/workflows/build.yml
```

---

## 许可

MIT License。引擎源自扣子 skill `thesis-format-doctor`，版权归原作者所有。
