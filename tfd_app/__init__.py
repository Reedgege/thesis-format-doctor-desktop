"""文档格式先生 · 桌面版 (Mr. Document Formatter — Desktop)"""

# 打包后 python-docx 的自带模板定位修复 —— 必须在任何 python-docx 调用之前跑。
#
# 为什么放在包初始化里：python-docx 读模板时用
# ``os.path.join(os.path.dirname(__file__), "..", "templates", "xxx.xml")``，
# 打包后 ``docx/parts/`` 不是真实目录，POSIX（macOS / Linux）解析这个带 ``..``
# 的路径会直接 ENOENT —— Windows 因为 Win32 先做词法归一化恰好没事。
#
# 本产品的主引擎是纯 XML 操作、不依赖 python-docx，所以检查/修正不受影响；
# 但 selfcheck、批注稿、修改报告这几条路一旦碰页眉/页脚就会在 mac 上崩，
# 而且崩之前不报错（python-docx 在部件懒创建时才抛）。见 docx_boot.py 的说明。
#
# 找不到模板目录时它原地不动（绝不硬塞错误路径），把原因留进 report() 供 selfcheck 输出。
try:
    from . import docx_boot

    docx_boot.apply()
except Exception:       # noqa: BLE001  初始化失败绝不允许拦住产品启动
    pass
