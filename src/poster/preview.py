# -*- coding: utf-8 -*-
"""PIL 预览渲染：在没有 PowerPoint / LibreOffice 的机器上校验版式。

要点
----
- 底图 = 模板 pptx 里的 ``ppt/media`` 中最大的一张，首次使用时解出并缓存，
  换机后可复现，不依赖外部目录；画布尺寸同样取自模板的 slide size。
- PIL 没有「虚线圆角矩形」，于是：① 用圆角掩膜贴渐变底；
  ② 手绘 4 条虚线直线段 + 4 段虚线圆弧（按弧长切 dash/gap）。
- 中文墨迹中心比字体 ascent/descent 中点低约 2px，所以胶囊文字按
  ``getbbox`` 的墨迹框中心定位，而不是 ``anchor="mm"``。
"""
from __future__ import annotations

import math
from pathlib import Path

from PIL import Image, ImageDraw

from . import config as C
from . import template as T
from .layout import Layout, cap_line_height, cap_rows_of, cap_size, fig_rows_of
from .typography import font, line_height, unpack, wrap_para

Image.MAX_IMAGE_PIXELS = None


def _script_dy(script: str, size_pt: float, scale: float) -> float:
    """预览里上/下标的纵向位移（px），与 ``config.SCRIPT_BASELINE`` 同口径。

    PowerPoint 的角标只把字形整体位移、不改变字号，预览也必须照做，
    否则交叉审核时看到的与实排不一致。

    ``scale`` 已经是「像素 / 厘米」，所以这里**只乘一次**。曾经多乘了一个
    ``C.PX_PER_CM``：3.65pt 的上标算出来 -1545px，被甩到画布外，预览里
    只剩下 ``Fe`` 后面的一块空白 —— 产物本身是对的，是预览把人骗了。
    """
    pct = C.SCRIPT_BASELINE.get(script)
    if pct is None:
        return 0.0
    return -pct / 100000.0 * size_pt * C.PT2CM * scale


# ---------------------------------------------------------------- 底图
def canvas_size(sc: float = C.PX_PER_CM) -> tuple[int, int]:
    """预览画布像素尺寸（cm -> px）。"""
    return (int(round(C.CANVAS_CM[0] * sc)), int(round(C.CANVAS_CM[1] * sc)))


def template_background(template_path) -> Image.Image:
    """取得模板背景图（尺寸对齐模板 slide size）。"""
    C.ensure_dirs()
    path = T.extract_background(template_path, C.CACHE)
    w, h = canvas_size()
    if path is None:                       # 模板里没有底图 -> 纯白画布
        return Image.new("RGB", (w, h), (255, 255, 255))
    return Image.open(path).convert("RGB").resize((w, h), Image.LANCZOS)


# ---------------------------------------------------------------- 绘制原语
def _vgrad(w_px: int, h_px: int, c1, c2, horiz: bool = False) -> Image.Image:
    w_px, h_px = max(1, w_px), max(1, h_px)
    if horiz:
        g = Image.new("RGB", (max(2, w_px), 1))
        n = g.width
        for j in range(n):
            t = j / (n - 1)
            g.putpixel((j, 0), tuple(int(c1[k] + (c2[k] - c1[k]) * t) for k in range(3)))
    else:
        g = Image.new("RGB", (1, max(2, h_px)))
        n = g.height
        for j in range(n):
            t = j / (n - 1)
            g.putpixel((0, j), tuple(int(c1[k] + (c2[k] - c1[k]) * t) for k in range(3)))
    return g.resize((w_px, h_px), Image.BILINEAR)


def _dline(dr, p0, p1, color, w=1) -> None:
    x0, y0 = p0
    x1, y1 = p1
    length = math.hypot(x1 - x0, y1 - y0)
    if length <= 0:
        return
    ux, uy = (x1 - x0) / length, (y1 - y0) / length
    d = 0.0
    while d < length:
        e = min(d + C.DASH_LEN, length)
        dr.line([x0 + ux * d, y0 + uy * d, x0 + ux * e, y0 + uy * e], fill=color, width=w)
        d = e + C.DASH_GAP


def _darc(dr, cx, cy, r, a0, a1, color, w=1) -> None:
    length = abs(a1 - a0) * r
    if length <= 0:
        return
    d = 0.0
    while d < length:
        e = min(d + C.DASH_LEN, length)
        s0 = a0 + (a1 - a0) * d / length
        s1 = a0 + (a1 - a0) * e / length
        pts = [(cx + r * math.cos(s0 + (s1 - s0) * t / 8),
                cy + r * math.sin(s0 + (s1 - s0) * t / 8)) for t in range(9)]
        dr.line(pts, fill=color, width=w)
        d = e + C.DASH_GAP


def _drrect(t, x, y, w, h, sc) -> None:
    """虚线 + 弧边圆角矩形（预览版）。"""
    x, y, w, h = x * sc, y * sc, w * sc, h * sc
    r = max(1.0, min(C.RADIUS * sc, min(w, h) / 2 - 0.5))
    x0, y0, x1, y1 = x, y, x + w, y + h
    dr = ImageDraw.Draw(t)
    pw, ph = max(1, int(round(w))), max(1, int(round(h)))
    mask = Image.new("L", (pw, ph), 0)
    ImageDraw.Draw(mask).rounded_rectangle([0, 0, pw - 1, ph - 1],
                                           radius=int(round(r)), fill=255)
    t.paste(_vgrad(pw, ph, C.BLOCK_C1, C.BLOCK_C2), (int(round(x0)), int(round(y0))), mask)
    _dline(dr, (x0 + r, y0), (x1 - r, y0), C.BLOCK_LINE)
    _dline(dr, (x1, y0 + r), (x1, y1 - r), C.BLOCK_LINE)
    _dline(dr, (x1 - r, y1), (x0 + r, y1), C.BLOCK_LINE)
    _dline(dr, (x0, y1 - r), (x0, y0 + r), C.BLOCK_LINE)
    _darc(dr, x0 + r, y0 + r, r, math.pi, math.pi * 1.5, C.BLOCK_LINE)
    _darc(dr, x1 - r, y0 + r, r, math.pi * 1.5, math.pi * 2, C.BLOCK_LINE)
    _darc(dr, x1 - r, y1 - r, r, 0.0, math.pi * 0.5, C.BLOCK_LINE)
    _darc(dr, x0 + r, y1 - r, r, math.pi * 0.5, math.pi, C.BLOCK_LINE)


def _dchip(t, x, y, w, h, text, sc) -> None:
    x, y, w, h = x * sc, y * sc, w * sc, h * sc
    pw, ph = max(1, int(round(w))), max(1, int(round(h)))
    mask = Image.new("L", (pw, ph), 0)
    ImageDraw.Draw(mask).rounded_rectangle([0, 0, pw - 1, ph - 1],
                                           radius=min(pw, ph) // 2, fill=255)
    t.paste(_vgrad(pw, ph, C.CHIP_C1, C.CHIP_C2, horiz=True),
            (int(round(x)), int(round(y))), mask)
    f = font(C.TITLE_SZ, True)
    b = f.getbbox(text)
    ImageDraw.Draw(t).text((x + w / 2 - (b[0] + b[2]) / 2, y + h / 2 - (b[1] + b[3]) / 2),
                           text, font=f, fill=(255, 255, 255), anchor="la")


def _pill_w(text: str) -> float:
    return font(C.TITLE_SZ, True).getlength(text) / C.PX_PER_CM + 2 * C.CHIP_PAD_X


# ---------------------------------------------------------------- 主渲染
def render(lay: Layout, task, figs, base: Image.Image | None = None,
           sc: float = C.PX_PER_CM) -> Image.Image:
    t = (base.copy() if base is not None else template_background(task.template))
    dr = ImageDraw.Draw(t)
    dr.rectangle([C.X0 * sc, C.Y0 * sc, (C.X0 + C.W) * sc, (C.Y0 + C.H) * sc], fill=C.PLATE)

    cells = lay.cells
    for c in cells:
        _drrect(t, c.x, c.y, c.bw, c.h, sc)
    for c in cells:
        cy = (c.y + C.TITLE_TOP + C.CHIP_H / 2) * sc
        _dline(dr, ((c.x + C.BP_X) * sc, cy), ((c.x + c.bw - C.BP_X) * sc, cy), C.CHIP_LINK)
    for c in cells:
        pill = min(_pill_w(task.titles[c.r]), c.bw - 2 * C.BP_X)
        _dchip(t, c.x + C.BP_X, c.y + C.TITLE_TOP, pill, C.CHIP_H, task.titles[c.r], sc)

    lh = line_height(lay.body_pt)
    cap_sz = cap_size(lay.body_pt)
    for c in cells:
        row_top = c.y + C.TITLE_TOP + C.CHIP_H + C.TITLE_GAP + c.off
        if c.style == "stack":
            yy = row_top
            iy = row_top + c.th + C.TEXT_IMG_GAP
            ix = c.x + C.BP_X + (c.inner - c.fw) / 2
        elif c.style == "side":
            yy = row_top + max(0.0, (c.ch - c.th) / 2)
            iy = row_top + max(0.0, (c.ch - c.figh) / 2)
            ix = c.x + c.bw - C.BP_X - c.fw
        else:                                              # 纯文本块
            yy = row_top

        body_font = font(lay.body_pt)
        for para in task.paras[c.r]:
            lines = wrap_para(para, lay.body_pt, c.tw * C.MEASURE_SAFETY,
                              policy=C.DRAW_POLICY)
            for k, (ln, lw) in enumerate(lines):
                x = (c.x + C.BP_X) * sc
                last = (k == len(lines) - 1)
                extra = 0.0 if last else (c.tw * sc - lw)
                sp = extra / (len(ln) - 1) if (not last and len(ln) > 1) else 0.0
                for seg in ln:
                    tok, color, script = unpack(seg)
                    dy = _script_dy(script, lay.body_pt, sc)
                    dr.text((x, yy * sc + dy), tok, font=body_font, fill=color)
                    x += body_font.getlength(tok) + sp
                yy += lh

        if not c.has_fig:
            continue

        cf = font(cap_sz)
        rows_bf = fig_rows_of(figs[c.r]) or [figs[c.r]]
        caps_rows = cap_rows_of(task, c.r, figs[c.r])
        y = iy
        for k, rb in enumerate(rows_bf):       # 图网格：行内横排、行间上叠
            ih = c.ihs[k] if k < len(c.ihs) else c.ihs[-1]
            caph = c.caphs[k] if k < len(c.caphs) else c.caphs[-1]
            caps = caps_rows[k] if k < len(caps_rows) else []
            fw_strip = (sum(ih * s for s in rb.spans)
                        + max(0, len(rb.paths) - 1) * C.FIG_GAP_IN)
            x = ix + max(0.0, c.fw - fw_strip) / 2
            for path, span, cap in zip(rb.paths, rb.spans, caps):
                im = Image.open(path)
                fw = ih * span
                px, py = int(round(x * sc)), int(round(y * sc))
                pw, ph = max(1, int(round(fw * sc))), max(1, int(round(ih * sc)))
                t.paste(im.resize((pw, ph), Image.LANCZOS), (px, py))
                dr.rectangle([px, py, px + pw - 1, py + ph - 1], outline=C.IMG_LINE, width=1)
                if cap:                       # 一张图一条注，居中在该图正下方
                    cy = (y + ih + C.CAP_LINEGAP) * sc
                    for ln, lw in wrap_para([(cap, C.GRAY)], cap_sz,
                                            fw * C.MEASURE_SAFETY,
                                            policy=C.DRAW_POLICY):
                        cx = (x + fw / 2) * sc - lw / 2
                        for seg in ln:
                            tok, col, script = unpack(seg)
                            dr.text((cx, cy + _script_dy(script, cap_sz, sc)), tok,
                                    font=cf, fill=col)
                            cx += cf.getlength(tok)
                        cy += cap_line_height(cap_sz) * sc
                x += fw + C.FIG_GAP_IN
            y += ih + C.CAP_LINEGAP + caph + C.FIG_ROW_GAP
    return t


def save_previews(lay: Layout, task, figs, intro_name: str = C.PREVIEW_INTRO,
                  full_name: str = C.PREVIEW_FULL,
                  max_w: int = 900, max_h: int = 2250) -> tuple[Path, Path]:
    """渲染并落盘两张预览：整张 / 内容区局部。"""
    sc = C.PX_PER_CM
    base = template_background(task.template)
    full = render(lay, task, figs, base.copy(), sc)
    thumb = full.copy()
    thumb.thumbnail((max_w, max_h), Image.LANCZOS)
    p_full = task.dist / full_name
    p_full.parent.mkdir(parents=True, exist_ok=True)
    thumb.save(p_full)

    intro = full.crop((0,
                       int((C.Y0 - C.PREVIEW_HEAD_PAD) * sc),
                       full.width,
                       int((C.Y1 + C.PREVIEW_FOOT_PAD) * sc)))
    p_intro = task.dist / intro_name
    intro.save(p_intro)
    return p_intro, p_full
