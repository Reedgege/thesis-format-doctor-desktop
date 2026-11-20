# -*- coding: utf-8 -*-
"""校验打出来的 .msix 清单身份/exe/版本是否正确（CI Windows 用）。

用法：
    python tools/check_msix.py release/MDF-Stu.msix mdf-stu

事实源：tools/msix/products.json（产品身份）+ 仓库 VERSION（版本）。
两者都不写死在校验里，版本升级不用改这个脚本。
"""
import json
import os
import re
import sys
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def norm_version(v):
    parts = [p for p in re.split(r"[.\-+]", str(v).strip().lstrip("vV")) if p.isdigit()]
    while len(parts) < 4:
        parts.append("0")
    return ".".join(str(int(p)) for p in parts[:4])


def main():
    if len(sys.argv) < 3:
        raise SystemExit("用法: python tools/check_msix.py <msix路径> <product键>")
    msix, prod = sys.argv[1], sys.argv[2]
    with open(os.path.join(ROOT, "tools", "msix", "products.json"), encoding="utf-8") as fh:
        products = json.load(fh)
    if prod not in products:
        raise SystemExit("products.json 里没有产品 %r" % prod)
    p = products[prod]
    with open(os.path.join(ROOT, "VERSION"), encoding="utf-8") as fh:
        version = norm_version(fh.read())

    with zipfile.ZipFile(msix) as z:
        names = z.namelist()
        assert "AppxManifest.xml" in names, "MSIX 里没有 AppxManifest.xml"
        xml = z.read("AppxManifest.xml").decode("utf-8")

    tokens = [
        p["identity_name"],
        p["publisher"],
        p["display_name"],
        p["executable"],
        version,
        "runFullTrust",
    ]
    for token in tokens:
        assert token in xml, "MSIX 清单缺: %s" % token
    # 清单里不能残留未替换占位符
    left = re.findall(r"@[A-Z_]+@", xml)
    assert not left, "清单残留占位符: %s" % sorted(set(left))
    # 磁贴素材必须齐
    for asset in ("StoreLogo.png", "Square150x150Logo.png", "Square44x44Logo.png",
                  "Square71x71Logo.png", "Square310x310Logo.png", "Wide310x150Logo.png",
                  "SplashScreen.png"):
        assert "Assets/" + asset in names, "MSIX 缺素材: " + asset

    print("OK MSIX 清单正确：%s / %s / v%s" % (p["identity_name"], p["display_name"], version))
    return 0


if __name__ == "__main__":
    sys.exit(main())
