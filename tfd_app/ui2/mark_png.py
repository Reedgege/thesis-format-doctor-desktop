# -*- coding: utf-8 -*-
"""品牌标 PNG 方案（零依赖）：纯标准库解码 PNG -> 按主题色染色 -> 平滑缩放 -> 交给 tkinter。

为什么不用 PhotoImage 直接加载：
  Tk 的图片缩放只有 zoom/subsample 两种，都是「整数倍最近邻」，缩到 40px 会出锯齿。
  而颜色更是死结：位图的颜色烧在像素里，无法跟着主题色变（学生版蓝 / 导师版朱砂红）。

本方案：自己解码 PNG（zlib + struct，标准库），把图当「覆盖度遮罩」用，
按目标尺寸做盒式重采样（自带抗锯齿效果），再染成主题色，
最后用 PhotoImage.put() 建位图。全程零第三方依赖、零图片资源入库（图由老板提供）。
"""
import struct
import zlib


# ---------- 1) 最小 PNG 解码（支持 color type 0/2/3/4/6，8bit，非隔行） ----------
def _paeth(a, b, c):
    p = a + b - c
    pa, pb, pc = abs(p - a), abs(p - b), abs(p - c)
    if pa <= pb and pa <= pc:
        return a
    if pb <= pc:
        return b
    return c


def load_png(path):
    """返回 (w, h, rgba_bytearray)。只处理 8bit 非隔行 PNG。"""
    raw = open(path, "rb").read()
    if raw[:8] != b"\x89PNG\r\n\x1a\n":
        raise ValueError("不是 PNG 文件")
    pos, idat, w = 8, [], 0
    h = bitdepth = ctype = interlace = 0
    plte = None
    trns = None
    while pos < len(raw):
        ln = struct.unpack(">I", raw[pos:pos + 4])[0]
        typ = raw[pos + 4:pos + 8]
        data = raw[pos + 8:pos + 8 + ln]
        if typ == b"IHDR":
            w, h, bitdepth, ctype, _comp, _filt, interlace = struct.unpack(">IIBBBBB", data)
        elif typ == b"PLTE":
            plte = data
        elif typ == b"tRNS":
            trns = data
        elif typ == b"IDAT":
            idat.append(data)
        elif typ == b"IEND":
            break
        pos += 12 + ln
    if bitdepth != 8:
        raise ValueError("只支持 8bit PNG（当前 %d）" % bitdepth)
    if interlace:
        raise ValueError("不支持隔行 PNG，请另存为普通 PNG")
    nch = {0: 1, 2: 3, 3: 1, 4: 2, 6: 4}[ctype]
    buf = zlib.decompress(b"".join(idat))
    stride = w * nch
    out = bytearray(h * stride)
    prev = bytearray(stride)
    p = 0
    for y in range(h):
        f = buf[p]
        p += 1
        line = bytearray(buf[p:p + stride])
        p += stride
        if f == 1:
            for i in range(nch, stride):
                line[i] = (line[i] + line[i - nch]) & 255
        elif f == 2:
            for i in range(stride):
                line[i] = (line[i] + prev[i]) & 255
        elif f == 3:
            for i in range(stride):
                a = line[i - nch] if i >= nch else 0
                line[i] = (line[i] + ((a + prev[i]) >> 1)) & 255
        elif f == 4:
            for i in range(stride):
                a = line[i - nch] if i >= nch else 0
                c = prev[i - nch] if i >= nch else 0
                line[i] = (line[i] + _paeth(a, prev[i], c)) & 255
        out[y * stride:(y + 1) * stride] = line
        prev = line

    # 统一转成 RGBA
    rgba = bytearray(w * h * 4)
    for i in range(w * h):
        s = i * nch
        if ctype == 0:
            v = out[s]
            rgba[i * 4:i * 4 + 4] = bytes((v, v, v, 255))
        elif ctype == 2:
            rgba[i * 4:i * 4 + 3] = out[s:s + 3]
            rgba[i * 4 + 3] = 255
        elif ctype == 3:
            k = out[s] * 3
            rgba[i * 4:i * 4 + 3] = plte[k:k + 3]
            rgba[i * 4 + 3] = trns[out[s]] if (trns and out[s] < len(trns)) else 255
        elif ctype == 4:
            v = out[s]
            rgba[i * 4:i * 4 + 3] = bytes((v, v, v))
            rgba[i * 4 + 3] = out[s + 1]
        else:
            rgba[i * 4:i * 4 + 4] = out[s:s + 4]
    return w, h, rgba


# ---------- 2) 取「覆盖度」遮罩 ----------
def coverage(w, h, rgba):
    """把图变成 0~255 的覆盖度遮罩。

    三种真实来源都能吃：
      · 透明底 + 深色图形（推荐）  → 用 alpha
      · 透明底 + 白色图形（放在色块上用）→ 用 alpha
      · 不透明底（米黄宣纸底 + 深色图形，如 icon.png）→ 自适应估底色，用暗度换算
    判据：透明像素占比够大 → 按「透明图」处理；否则按「不透明底图」处理。
    """
    n = w * h
    transp = 0
    for i in range(n):
        if rgba[i * 4 + 3] < 128:
            transp += 1
    frac = transp / float(n)
    cov = bytearray(n)
    if frac > 0.30:                             # 透明底
        for i in range(n):
            cov[i] = rgba[i * 4 + 3]
        return cov, "透明底"
    # —— 不透明底：估底色亮度 + 估「墨色主调」，用两者之间的跨度归一化 ——
    #   要点：不能用「最暗像素」当基准 —— 图里常有个别近黑的点（描边/阴影），
    #   用它做基准会让整片淡彩羽毛都到不了满覆盖 → 渲染出来发虚。
    #   正确做法：取「明显比底色暗」的像素里出现最多的那一档作为墨色基准。
    luma = bytearray(n)
    hist = [0] * 32
    for i in range(n):
        s = i * 4
        v = (rgba[s] * 299 + rgba[s + 1] * 587 + rgba[s + 2] * 114) // 1000
        luma[i] = v
        hist[v >> 3] += 1
    bgl = (max(range(32), key=lambda k: hist[k]) << 3) + 4

    h2 = [0] * 32
    for i in range(n):
        if rgba[i * 4 + 3] > 128 and luma[i] < bgl - 18:
            h2[luma[i] >> 3] += 1
    il = ((max(range(32), key=lambda k: h2[k]) << 3) + 4) if any(h2) else 0
    span = max(1, bgl - il)
    for i in range(n):
        if rgba[i * 4 + 3] < 128:               # 图自己标了透明的（如圆角外）→ 不算墨
            cov[i] = 0
            continue
        v = bgl - luma[i]
        if v <= 14:                             # 底色噪声 → 归零（米黄底约 7 luma）
            cov[i] = 0
        else:
            cov[i] = 255 if v >= span else (v * 255 // span)
    return cov, "不透明底"


def mark_size(png_path, height, clean=True, boost=None):
    """按目标高度算出裁切后的「紧贴尺寸」(ow, oh) —— 供排版精确定位用。

    为什么要紧贴尺寸：如果把标塞进一个正方形盒子，图形两侧会留下看不见的空白，
    导致「标到标题的间距」比设定的值大出一截、左右留白还不对称 —— 看着就"别扭"。
    """
    w, h, rgba = load_png(png_path)
    cov, _mode = coverage(w, h, rgba)
    if clean:
        cov, _n, _d = keep_main(cov, w, h)
    if boost:
        cov = boost_cov(cov, boost[0], boost[1])
    x0, y0, x1, y1 = ink_bbox(cov, w, h)
    cw, ch = max(1, x1 - x0), max(1, y1 - y0)
    ow = max(1, int(cw * (height / float(ch)) + 0.5))
    return ow, height


def ink_bbox(cov, w, h, thr=26):
    """墨迹包围盒 —— 老板给的图常带大量留白，必须裁掉再适配尺寸。"""
    x0, y0, x1, y1 = w, h, -1, -1
    for y in range(h):
        base = y * w
        for x in range(w):
            if cov[base + x] > thr:
                if x < x0:
                    x0 = x
                if x > x1:
                    x1 = x
                if y < y0:
                    y0 = y
                if y > y1:
                    y1 = y
    if x1 < 0:
        return 0, 0, w, h
    return x0, y0, x1 + 1, y1 + 1


# ---------- 3) 平滑缩放到目标尺寸（盒式平均 = 自带抗锯齿） ----------
def resample(cov, w, h, out_w, out_h):
    """只缩遮罩（染色模式用，够快）。"""
    res = bytearray(out_w * out_h)
    if out_w <= 0 or out_h <= 0:
        return res
    sx, sy = w / float(out_w), h / float(out_h)
    for oy in range(out_h):
        y0, y1 = int(oy * sy), min(h, max(int(oy * sy) + 1, int((oy + 1) * sy)))
        for ox in range(out_w):
            x0, x1 = int(ox * sx), min(w, max(int(ox * sx) + 1, int((ox + 1) * sx)))
            tot = cnt = 0
            for y in range(y0, y1):
                base = y * w
                for x in range(x0, x1):
                    tot += cov[base + x]
                    cnt += 1
            res[oy * out_w + ox] = tot // cnt if cnt else 0
    return res


def resample_rgba(cov, pr, pg, pb, w, h, out_w, out_h):
    """遮罩 + 预乘颜色一起缩（原色模式用）。

    在「预乘空间」里做平均是数学上正确的做法：
      合成公式 out = bg·(1-m̄) + mean(m·c)，其中 mean(m·c) 就是预乘通道的平均值。
    所以不必先解出源色再重合成 —— 直接平均预乘值即可，边缘不会出现黑边/白边。
    """
    om = [0.0] * (out_w * out_h)
    opr = [0.0] * (out_w * out_h)
    opg = [0.0] * (out_w * out_h)
    opb = [0.0] * (out_w * out_h)
    if out_w <= 0 or out_h <= 0:
        return om, opr, opg, opb
    sx, sy = w / float(out_w), h / float(out_h)
    for oy in range(out_h):
        y0, y1 = int(oy * sy), min(h, max(int(oy * sy) + 1, int((oy + 1) * sy)))
        for ox in range(out_w):
            x0, x1 = int(ox * sx), min(w, max(int(ox * sx) + 1, int((ox + 1) * sx)))
            tm = tr = tg = tb = 0
            cnt = 0
            for y in range(y0, y1):
                base = y * w
                for x in range(x0, x1):
                    i = base + x
                    tm += cov[i]
                    tr += pr[i]
                    tg += pg[i]
                    tb += pb[i]
                    cnt += 1
            k = oy * out_w + ox
            if cnt:
                om[k] = tm / float(cnt)
                opr[k] = tr / float(cnt)
                opg[k] = tg / float(cnt)
                opb[k] = tb / float(cnt)
    return om, opr, opg, opb


def premultiplied(rgba, w, h, cov, mode, bg_hex):
    """把源像素折算成「预乘颜色」。

    · 透明底图：RGB 已是真色 → 预乘 = RGB × 覆盖度
    · 不透明底图：像素 = 源色×m + 底色×(1-m) → 反解 预乘 = 像素 - 底色×(1-m)
      （这样底色方块自动消失，且不会在深色羽毛周围留一圈浅色光晕）
    """
    br, bg_, bb = _rgb(bg_hex)
    n = w * h
    pr = bytearray(n)
    pg = bytearray(n)
    pb = bytearray(n)
    if mode == "透明底":
        for i in range(n):
            s = i * 4
            m = cov[i]
            pr[i] = rgba[s] * m // 255
            pg[i] = rgba[s + 1] * m // 255
            pb[i] = rgba[s + 2] * m // 255
    else:
        for i in range(n):
            s = i * 4
            m = cov[i]
            k = (255 - m) / 255.0
            pr[i] = max(0, min(255, int(rgba[s] - br * k + 0.5)))
            pg[i] = max(0, min(255, int(rgba[s + 1] - bg_ * k + 0.5)))
            pb[i] = max(0, min(255, int(rgba[s + 2] - bb * k + 0.5)))
    return pr, pg, pb



# ---------- 4) 染主题色 + 合成到底色 -> 交给 Tk ----------
def _rgb(h):
    h = h.lstrip("#")
    return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)


def ink_color(png_path, cov=None):
    """估原图的「代表墨色」——按面积取主色（不是取最暗的核心像素）。

    只作参考/兜底用；真正保真的原色模式走逐像素合成，不用这个单色。
    """
    w, h, rgba = load_png(png_path)
    if cov is None:
        cov, _ = coverage(w, h, rgba)
    cnt = {}
    for i in range(w * h):
        if cov[i] > 128:
            s = i * 4
            k = (rgba[s] // 12 * 12, rgba[s + 1] // 12 * 12, rgba[s + 2] // 12 * 12)
            cnt[k] = cnt.get(k, 0) + 1
    if not cnt:
        return "#000000"
    r, g, b = max(cnt.items(), key=lambda kv: kv[1])[0]
    return "#%02x%02x%02x" % (r, g, b)


def build_mark(tk_master, png_path, size, color_hex, bg_hex,
               box_w=None, box_h=None, clean=True, boost=None, cache={}):
    """把 PNG 渲染成「指定尺寸、指定颜色、铺在指定底色上」的 PhotoImage。

    color_hex 传值  → 上主题色（学生蓝 / 导师红自动切换）
    color_hex 传 None → 保留老板自己的上色（逐像素保真，含描边/渐变）

    流程：解码 → 取墨迹遮罩 → 自动裁到墨迹包围盒 → 等比缩放进 (box_w, box_h)
          → 上色 → 合成到底色（保留抗锯齿）。

    要点：
      · Tk 的 PhotoImage.put() 不支持半透明，但品牌标永远坐在「信笺抬头带」
        的纯色底上 —— 直接合成到底色，0~255 的过渡边缘完整保留。
      · 自动裁白边：老板给的图常是「大方块 + 中间一小片图形」，
        不裁的话图形会显得比设定尺寸小一圈。
      · 底色方块自动消失：因为拿的是「墨迹遮罩」而不是整张图。
    """
    import tkinter as tk
    bw = box_w or size
    bh = box_h or size
    asis = color_hex is None
    key = (png_path, bw, bh, color_hex or "@asis", bg_hex)
    if key in cache:
        return cache[key]

    w, h, rgba = load_png(png_path)
    cov, mode = coverage(w, h, rgba)
    if clean:
        cov, _n, _d = keep_main(cov, w, h)      # 去水印/杂点，否则包围盒会被撑大
    if boost:
        cov = boost_cov(cov, boost[0], boost[1])
    x0, y0, x1, y1 = ink_bbox(cov, w, h)
    cw, ch = max(1, x1 - x0), max(1, y1 - y0)

    def crop(buf):
        out = bytearray(cw * ch)
        for y in range(ch):
            s = (y + y0) * w + x0
            out[y * cw:(y + 1) * cw] = buf[s:s + cw]
        return out

    k = min(bw / float(cw), bh / float(ch))
    ow, oh = max(1, int(cw * k + 0.5)), max(1, int(ch * k + 0.5))
    br, bg_, bb = _rgb(bg_hex)
    ox, oy = (bw - ow) // 2, (bh - oh) // 2

    img = tk.PhotoImage(master=tk_master, width=bw, height=bh)
    if asis:
        pr, pg, pb = premultiplied(rgba, w, h, cov, mode, bg_hex)
        m, sr, sg, sb = resample_rgba(crop(cov), crop(pr), crop(pg), crop(pb),
                                      cw, ch, ow, oh)
    else:
        cm = resample(crop(cov), cw, ch, ow, oh)
        m = [float(v) for v in cm]
        fr, fg, fb = _rgb(color_hex)
        # 染色模式：预乘 = 颜色 × 覆盖度
        sr = [fr * v / 255.0 for v in m]
        sg = [fg * v / 255.0 for v in m]
        sb = [fb * v / 255.0 for v in m]

    bghex = "#%02x%02x%02x" % (br, bg_, bb)
    rows = []
    for y in range(bh):
        row = []
        my = y - oy
        for x in range(bw):
            mx = x - ox
            if 0 <= my < oh and 0 <= mx < ow:
                i = my * ow + mx
                mv = m[i]
                if mv <= 0.4:
                    row.append(bghex)
                    continue
                # out = bg·(1-m̄) + 预乘值 —— 与源色无关的通用合成
                row.append("#%02x%02x%02x" % (
                    min(255, int(br * (1 - mv / 255.0) + sr[i] + 0.5)),
                    min(255, int(bg_ * (1 - mv / 255.0) + sg[i] + 0.5)),
                    min(255, int(bb * (1 - mv / 255.0) + sb[i] + 0.5))))
            else:
                row.append(bghex)
        rows.append("{" + " ".join(row) + "}")
    img.put(" ".join(rows))
    cache[key] = img
    return img


def write_png_rgba(path, w, h, rgba):
    """写 8bit RGBA PNG（纯标准库）—— 用于离线产出品牌标资源。

    与 load_png 配成一对，闭环可自测（写完再读回来核对）。
    """
    raw = bytearray()
    for y in range(h):
        raw.append(0)                       # 过滤器类型 0（None）
        raw += rgba[y * w * 4:(y + 1) * w * 4]

    def chunk(typ, data):
        return (struct.pack(">I", len(data)) + typ + data
                + struct.pack(">I", zlib.crc32(typ + data) & 0xffffffff))

    png = (b"\x89PNG\r\n\x1a\n"
           + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 6, 0, 0, 0))
           + chunk(b"IDAT", zlib.compress(bytes(raw), 9))
           + chunk(b"IEND", b""))
    with open(path, "wb") as f:
        f.write(png)
    return len(png)


def boost_cov(cov, lo=0, hi=255):
    """墨量加浓：把 [lo, hi] 线性拉满到 [0, 255]，低于 lo 的归零。

    用途：细笔画的标（AI 出的图常见）缩到 24~40px 会发虚，
    把中间调抬成实心就能看清。lo 越小、hi 越小 → 越浓。
    实测豆包细笔画羽毛在 (20, 170) 下「全实像素」由 10% 提到 ~28%。
    """
    if lo <= 0 and hi >= 255:
        return cov
    span = max(1.0, hi - lo)
    lut = bytearray(256)
    for v in range(256):
        t = (v - lo) / span
        lut[v] = 0 if t <= 0 else (255 if t >= 1 else int(t * 255 + 0.5))
    return bytearray(lut[v] for v in cov)


def render_rgba(png_path, out_h, color_hex=None, clean=True, boost=None):
    """把源 PNG 重绘成「透明底 + 指定颜色」的 RGBA 图（离线产出资源用）。

    color_hex 传值 → 平涂该色（边缘保留遮罩的抗锯齿）
    color_hex=None → 保留源图原色（在预乘空间反解，边缘不发灰）
    clean=True     → 先做连通域去杂（去水印/杂点，见 keep_main）
    """
    w, h, rgba = load_png(png_path)
    cov, mode = coverage(w, h, rgba)
    if clean:
        cov, _n, _d = keep_main(cov, w, h)
    if boost:
        cov = boost_cov(cov, boost[0], boost[1])
    x0, y0, x1, y1 = ink_bbox(cov, w, h)
    cw, ch = max(1, x1 - x0), max(1, y1 - y0)

    def crop(buf):
        out = bytearray(cw * ch)
        for y in range(ch):
            s = (y + y0) * w + x0
            out[y * cw:(y + 1) * cw] = buf[s:s + cw]
        return out

    ow = max(1, int(cw * (out_h / float(ch)) + 0.5))
    oh = out_h
    if color_hex is None:
        pr, pg, pb = premultiplied(rgba, w, h, cov, mode, "#FFFFFF")
        m, sr, sg, sb = resample_rgba(crop(cov), crop(pr), crop(pg), crop(pb),
                                      cw, ch, ow, oh)
    else:
        cm = resample(crop(cov), cw, ch, ow, oh)
        m = [float(v) for v in cm]
        r, g, b = _rgb(color_hex)
        sr = [r * v / 255.0 for v in m]
        sg = [g * v / 255.0 for v in m]
        sb = [b * v / 255.0 for v in m]

    out = bytearray(ow * oh * 4)
    for i in range(ow * oh):
        mv = m[i]
        a = 255 if mv > 254.5 else int(mv + 0.5)
        if a <= 0:
            continue
        k = 255.0 / mv
        out[i * 4] = max(0, min(255, int(sr[i] * k + 0.5)))
        out[i * 4 + 1] = max(0, min(255, int(sg[i] * k + 0.5)))
        out[i * 4 + 2] = max(0, min(255, int(sb[i] * k + 0.5)))
        out[i * 4 + 3] = a
    return ow, oh, out


def keep_main(cov, w, h, cell=8, min_cov=22, min_ratio=0.05):
    """只保留最大的一块墨迹（去掉水印、杂点、JPEG 噪块）。

    为什么必须有这一步：AI 生成的图常带右下角水印 + 散布的浅色噪块，
    它们会**把墨迹包围盒撑大**（实测把羽毛 927 宽撑成 1338），
    导致裁出来的图里羽毛被压小、旁边还拖着一块水印。
    做法：格子级占用 → 8 邻接连通域 → 只留最大的一块（并清掉小于它 min_ratio 的块）。
    """
    gw, gh = max(1, w // cell), max(1, h // cell)
    occ = [[0] * gw for _ in range(gh)]
    for gy in range(gh):
        base = gy * cell
        for gx in range(gw):
            n = 0
            for y in range(base, min(base + cell, h)):
                row = y * w
                for x in range(gx * cell, min((gx + 1) * cell, w)):
                    if cov[row + x] >= min_cov:
                        n += 1
            occ[gy][gx] = n

    seen = [[False] * gw for _ in range(gh)]
    comps = []
    for gy in range(gh):
        for gx in range(gw):
            if occ[gy][gx] < 2 or seen[gy][gx]:
                continue
            stack = [(gx, gy)]
            seen[gy][gx] = True
            cells = []
            while stack:
                cx, cy = stack.pop()
                cells.append((cx, cy))
                for dx in (-1, 0, 1):
                    for dy in (-1, 0, 1):
                        nx, ny = cx + dx, cy + dy
                        if (0 <= nx < gw and 0 <= ny < gh and not seen[ny][nx]
                                and occ[ny][nx] >= 2):
                            seen[ny][nx] = True
                            stack.append((nx, ny))
            comps.append((sum(occ[cy][cx] for cx, cy in cells), cells))
    if not comps:
        return cov, 0, 0
    comps.sort(key=lambda t: -t[0])
    keep = set()
    biggest = comps[0][0]
    for pix, cells in comps:
        if pix >= biggest * min_ratio:
            keep.update(cells)
    drop = 0
    for gy in range(gh):
        for gx in range(gw):
            if (gx, gy) in keep:
                continue
            for y in range(gy * cell, min((gy + 1) * cell, h)):
                row = y * w
                for x in range(gx * cell, min((gx + 1) * cell, w)):
                    if cov[row + x]:
                        cov[row + x] = 0
                        drop += 1
    return cov, len(comps), drop


def coverage_stats(png_path):
    """给老板看的口径：这张图被怎么解读、裁掉多少白边。"""
    w, h, rgba = load_png(png_path)
    cov, mode = coverage(w, h, rgba)
    x0, y0, x1, y1 = ink_bbox(cov, w, h)
    n = len(cov)
    solid = sum(1 for v in cov if v > 200)
    return {"size": (w, h), "mode": mode, "ink_ratio": solid / float(n),
            "bbox": (x0, y0, x1, y1),
            "trim": "裁掉 %.0f%% 宽 / %.0f%% 高"
                    % (100.0 * (1 - (x1 - x0) / float(w)),
                       100.0 * (1 - (y1 - y0) / float(h)))}
