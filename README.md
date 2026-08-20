# 论文格式医生 · 桌面版 (Thesis Format Doctor — Desktop)

把你用**扣子(Coze)** 做的论文格式处理 skill（`thesis-format-doctor`）变成一个**真正能在本机跑的桌面程序**：

- 🪟 **Windows**：双击 `thesis-format-doctor-desktop.exe`
- 🍎 **macOS**：双击 `thesis-format-doctor-desktop.app` / 可执行文件
- 🐧 **Linux**：运行 `thesis-format-doctor-desktop`

双击后自动打开浏览器进入操作界面，**完全离线、零第三方依赖、文件不离开你的电脑**。

---

## 它能做什么

源自 skill 的纯 Python 标准库引擎（已重命名为可 import 的模块）：

| 功能 | 说明 |
| --- | --- |
| 🩺 **格式体检** | 对照通用学术红线（标题样式、三线表、参考文献、页边距、全角标点等）逐项诊断，按高危/中危/低危分级，可下载 Markdown 报告 |
| 🎯 **套标题 & 出目录** | 自动识别标题层级并套用 Heading 样式，让 Word/WPS 可一键出目录；生成逐条修改明细 docx |
| 📚 **重排参考文献** | 按 **GB/T 7714** 范式重排参考文献区 |
| 📐 **抽取模板画像** | 上传学校模板，自动抽取其格式规范（以模板批注为准），用于更精确的"模板驱动诊断/套用" |

> 体检默认是「通用规范体检」；上传学校模板并抽取画像后，会切换为更贴合要求的「模板驱动诊断」。

---

## 使用方法

1. 下载对应系统的安装包（见 [Releases](../../releases)）。
2. 双击运行（Windows 可能弹出"未知发布者"提示，点"仍要运行"即可——这是没花钱买代码签名证书的正常现象）。
3. 浏览器自动打开操作界面：
   - 选择「待处理论文 .docx」（必选）
   - 可选：选择「学校模板 .docx」用于精确诊断
   - 点对应按钮执行，结果在右侧展示，可下载处理后的文档
4. 关闭浏览器/程序即可退出。

---

## 技术说明

- 后端：Python 标准库 `http.server` 提供本地服务，前端为纯静态 HTML/CSS/JS。
- 引擎：直接复用 skill 中的 `docxutils / format_checker / headings_fix / ref_reformat / format_profile / report_docx`，**零第三方依赖**，因此打包极小、跨平台几乎无依赖坑。
- 打包：PyInstaller `--onefile`，见 `build.py` 与 `.github/workflows/build.yml`。

### 本地自行构建（可选）

```bash
pip install pyinstaller
python build.py            # 按当前系统构建到 dist/
```

### 发版（GitHub Actions 自动构建三平台）

推送一个 `v*` 标签即可触发：

```bash
git tag v1.3.5
git push origin v1.3.5
```

Actions 会在 Windows / macOS / Ubuntu 三个 runner 上分别构建，并把产物发布到 GitHub Release。

---

## 目录结构

```
thesis-format-doctor-desktop/
├── main.py                 # 启动入口（启动服务 + 打开浏览器）
├── build.py                # 跨平台 PyInstaller 构建脚本
├── pyproject.toml
├── tfd_app/
│   ├── server.py           # 本地 Web 服务 + 引擎封装
│   ├── frontend/           # 前端界面 (index.html / app.js / styles.css)
│   ├── core/               # skill 引擎（标准库，零依赖）
│   └── assets/
└── .github/workflows/build.yml
```

---

## 许可

MIT License。引擎源自扣子 skill `thesis-format-doctor`，版权归原作者所有。
