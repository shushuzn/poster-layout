# -*- coding: utf-8 -*-
"""配图预处理：去透明底 + 裁白边，并量出每张图的宽高比。

从文档里抽出来的图多为 RGBA / 调色板模式，直接贴到 PPTX 上透明区会变黑；
四周还带大片白边，会浪费宝贵的版面高度。两步都在**任务独占的缓存目录**里落盘，
避免每次构建重复处理 3000+ px 的大图，也避免不同任务的同名图互相覆盖。

量出的比例交给 :mod:`poster.layout` 求解图条尺寸：单块可配多张图，
它们**等高横排**成一条（见 ``docs/design.md`` 的「图条」一节）。
"""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

import numpy as np
from PIL import Image

from . import config as C

Image.MAX_IMAGE_PIXELS = None

TRIM_THRESHOLD = 246
TRIM_PAD = 6


def _flatten(src: Path, cache: Path) -> Path:
    """透明底 -> 白底，输出到缓存目录。"""
    dst = cache / f"flat_{src.stem}.png"
    if dst.exists() and dst.stat().st_mtime >= src.stat().st_mtime:
        return dst
    im = Image.open(src)
    if im.mode in ("RGBA", "LA", "P"):
        im = im.convert("RGBA")
        bg = Image.new("RGB", im.size, (255, 255, 255))
        bg.paste(im, mask=im.split()[-1])
        im = bg
    else:
        im = im.convert("RGB")
    dst.parent.mkdir(parents=True, exist_ok=True)
    im.save(dst)
    return dst


def trim_white(src: Path, cache: Path | None = None) -> Path:
    """裁掉四周白边（带少量 padding），输出到缓存目录。"""
    cache = cache or C.CACHE
    flat = _flatten(src, cache)
    dst = cache / f"trim_{src.stem}.png"
    if dst.exists() and dst.stat().st_mtime >= flat.stat().st_mtime:
        return dst
    im = Image.open(flat)
    a = np.array(im.convert("L"))
    ys, xs = np.nonzero(a < TRIM_THRESHOLD)
    if len(ys) == 0:
        return flat
    y0, y1 = max(0, ys.min() - TRIM_PAD), min(im.height, ys.max() + TRIM_PAD + 1)
    x0, x1 = max(0, xs.min() - TRIM_PAD), min(im.width, xs.max() + TRIM_PAD + 1)
    im.crop((x0, y0, x1, y1)).save(dst)
    return dst


@dataclass
class BlockFigures:
    """一个分块的配图（0~n 张）。

    ``spans[i]`` 是第 i 张图的 **w/h**，也就是「高 1cm 时该占多宽」；
    图条等高横排时，整条的单位高宽 ``total = Σ spans``，
    给定图高 ``h`` 则图条宽（不含张间间距）= ``h * total``。
    """

    paths: list[Path] = field(default_factory=list)
    spans: list[float] = field(default_factory=list)

    # 图网格：rows[k] = 第 k 条图条行（行内等高并排，行间上下叠）。
    # paths / spans 是 rows 的拍平，只为兼容单行的老逻辑；一律按 rows 求解。
    rows: list["BlockFigures"] = field(default_factory=list)

    @property
    def nrows(self) -> int:
        return max(1, len(self.rows))

    @property
    def empty(self) -> bool:
        return not self.paths

    @property
    def total(self) -> float:
        return sum(self.spans)

    def widths(self, h: float) -> list[float]:
        """图条高为 ``h`` 时，各张图的宽。"""
        return [h * s for s in self.spans]

    def footprint(self, h: float) -> float:
        """图条高为 ``h`` 时，整条占用的总宽（含张间间距）。"""
        if self.empty:
            return 0.0
        return h * self.total + (len(self.paths) - 1) * C.FIG_GAP_IN


def prepare_figures(task) -> list[BlockFigures]:
    """按任务文档预处理各块的配图，返回与 ``task.sections`` 等长的列表。"""
    cache = C.task_cache(task.id)
    out: list[BlockFigures] = []
    for sec in task.sections:
        bf = BlockFigures()
        groups = sec.fig_rows_paths or ([sec.figure_paths] if sec.figure_paths else [])
        for group in groups:
            row = BlockFigures()
            for src in group:
                p = trim_white(src, cache)
                with Image.open(p) as im:
                    row.spans.append(im.width / im.height)
                row.paths.append(p)
            bf.rows.append(row)
            bf.paths.extend(row.paths)
            bf.spans.extend(row.spans)
        if not bf.rows:
            bf.rows.append(bf)          # 无图：唯一一行就是自己（空）
        out.append(bf)
    return out
