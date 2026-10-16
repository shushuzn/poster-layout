# -*- coding: utf-8 -*-
"""模板探测：把"某赛事的官方模板"里与排版有关的事实读出来。

原来这些值是写死的（画布 8×20cm、内容区 0.74/8.74/7.30/18.30cm、
背景图就是 ``ppt/media/image1.*``）。换成通用引擎后必须靠探测 —— 换一个赛事、
换一份模板，不需要改代码。

探测到的三件事
--------------
1. **画布尺寸** —— 取 slide size，而不是假设 8×20cm。
2. **内容区** —— 找模板自带的「作品简介」占位文本框，按 :data:`config.AREA_INSET`
   内缩得到真正可排版的矩形。探测不到就退回 :data:`config.FALLBACK_AREA`。
3. **背景图** —— 取 ``ppt/media`` 里面积最大的一张，用于预览底图。
"""
from __future__ import annotations

import zipfile
from pathlib import Path

from pptx import Presentation

from . import config as C

_IMG_EXT = (".jpeg", ".jpg", ".png", ".bmp", ".gif", ".emf", ".wmf")


def cm(emu: int) -> float:
    """EMU -> cm。"""
    return emu / 360000.0


def slide_size_cm(path) -> tuple[float, float]:
    """模板画布尺寸（cm）。"""
    prs = Presentation(str(path))
    return round(cm(prs.slide_width), 3), round(cm(prs.slide_height), 3)


def is_placeholder(shape) -> bool:
    """是否为模板自带的「作品简介」占位文本框。

    模板并非空白：这个框里残留着旧文字，构建时必须清空，否则会和新内容叠印。
    其余模板字段（作品名称 / 项目成员 / 作品编号 / 本科生组）一律不动。

    判定规则（可在任务文档 ``style:`` 里覆盖 ``PH_TOP`` / ``PH_MIN_W``）：
    有文本框 且 顶部 ≈ :data:`config.PH_TOP` 且 宽度 ≥ :data:`config.PH_MIN_W`。
    """
    return (getattr(shape, "has_text_frame", False)
            and abs(cm(shape.top) - C.PH_TOP) < C.PH_TOP_TOL
            and cm(shape.width) >= C.PH_MIN_W)


def placeholder_rect(path):
    """返回占位框的 ``(x, y, w, h)``（cm），没有则 ``None``。"""
    prs = Presentation(str(path))
    for sh in prs.slides[0].shapes:
        if is_placeholder(sh):
            return (cm(sh.left), cm(sh.top), cm(sh.width), cm(sh.height))
    return None


def detect_area(path) -> tuple[float, float, float, float] | None:
    """由占位框推出可排版内容区 ``(x0, y0, x1, y1)``（cm）。

    占位框比真正的白底内区略大，故按 :data:`config.AREA_INSET` 向内收。
    探测不到时返回 ``None``（由调用方退回 ``config.FALLBACK_AREA``）。
    """
    rect = placeholder_rect(path)
    if rect is None:
        return None
    # 模板里形状的坐标带亚 0.01cm 噪声（EMU 换算出来是 0.6943 这种），
    # 先吸附到 0.01cm 再内缩 —— 与当初实测 AREA_INSET 的口径一致。
    x, y, w, h = (round(v, 2) for v in rect)
    il, it, ir, ib = C.AREA_INSET
    return (round(x + il, 3), round(y + it, 3),
            round(x + w - ir, 3), round(y + h - ib, 3))


def extract_background(template_path, cache_dir: Path) -> Path | None:
    """把模板背景图解包到缓存目录（首次才做），返回图片路径。"""
    template_path = Path(template_path)
    dst = cache_dir / ("bg_%s.jpg" % template_path.stem.replace(" ", "_"))
    if dst.exists() and dst.stat().st_mtime >= template_path.stat().st_mtime:
        return dst

    with zipfile.ZipFile(str(template_path)) as z:
        infos = [i for i in z.infolist()
                 if i.filename.startswith("ppt/media/")
                 and i.filename.lower().endswith(_IMG_EXT)]
        if not infos:
            return None
        # 取未压缩体积最大的一张当背景
        info = max(infos, key=lambda i: i.file_size)
        data = z.read(info.filename)

    dst.parent.mkdir(parents=True, exist_ok=True)
    dst.write_bytes(data)
    return dst
