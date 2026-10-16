# -*- coding: utf-8 -*-
"""版面求解器：网格 + 图文排法 + 尺寸自动搜索。

约束链（互相耦合，无法独立取最优，故做搜索）
--------------------------------------------
内容区固定（示例为 6.56 x 9.56cm），要同时装下若干段正文、若干个「图条」和若干胶囊标题。
- 字号 ↑ → 正文行数可能不变但行高 ↑ → 总高 ↑
- 图 ↑   → 「文左图右」块的正文列变窄 → 行数 ↑ → 总高 ↑
- 两者都与 `AVAIL_H` 抢空间

图条
----
一个分块可以配 0~n 张图。多张时**等高横排**成一条，宽度按各自 w/h 分配，
于是整条只需一个「单位高宽」 :attr:`poster.imaging.BlockFigures.total` 来描述：

    图条宽（不含张间间距） = 图高 x total

单图只是 n=1 的特例（``total = w/h``），因此下面的公式对 0/1/n 张图统一成立。

排法不是全篇统一的，而是**按图条的形态逐块决定**（任务文档里可用 ``mode:`` 覆盖）：
- 无配图（``mode: text``）      → 纯文本，占满块宽
- 超宽条（w/h > 3.0）        → 文上图下，图铺满块宽作通栏横幅
- 方形（w/h <= 1.2）         → 文左图右，方图居右配文最正
- 其余                      → 通栏块文左图右 / 并排块文上图下
"""
from __future__ import annotations

from dataclasses import dataclass, field

from . import config as C
from .typography import line_count, line_height, lines_count_worst

# ---------------------------------------------------------------- 数据结构


@dataclass
class Cell:
    r: int                 # 正文段序号
    bw: float              # 块宽
    inner: float           # 块内可用宽
    style: str             # 'text' 无图 / 'stack' 文上图下 / 'side' 文左图右
    tw: float              # 正文列宽
    th: float              # 正文高
    nl: int                # 正文行数
    fw: float              # 图条总占地宽（含张间间距；无图为 0）
    ih: float              # 图条高（无图为 0）
    ch: float              # 内容高
    caph: float = 0.0      # 图注区高（= 各张图注行数的最大值 x 行高）
    x: float = 0.0
    y: float = 0.0
    h: float = 0.0
    gh: float = 0.0
    off: float = 0.0

    # 图网格：ihs[k] / caphs[k] = 第 k 条图条行的图高与图注高。
    # 单行时 ihs == [ih]，与老语义完全一致。
    ihs: list = field(default_factory=list)
    caphs: list = field(default_factory=list)
    figh: float = 0.0      # 图区总高（= Σ行(图高 + 图注) + 行间距）

    def __post_init__(self):
        if not self.ihs:
            self.ihs = [self.ih]
        if not self.caphs:
            self.caphs = [self.caph]
        if not self.figh:
            self.figh = self.ih + C.CAP_LINEGAP + self.caph if self.has_fig else 0.0

    @property
    def nrows(self) -> int:
        """图条行数（无图为 0）。"""
        return 0 if not self.has_fig else len(self.ihs)

    @property
    def has_fig(self) -> bool:
        return self.style != "text"


@dataclass
class Row:
    cells: list[Cell]
    bh: float
    y: float = 0.0


@dataclass
class Layout:
    body_pt: float
    ih_cap: float
    rows: list[Row]
    total: float
    gap_v: float
    top0: float

    @property
    def cells(self) -> list[Cell]:
        """所有块（按行、按列展开）。"""
        return [c for row in self.rows for c in row.cells]


# ---------------------------------------------------------------- 图条形态


def strip_ratio(bf) -> float | None:
    """图条的 h/w；无配图返回 ``None``。

    块内有多行图网格时取**最宽那行**（它决定形态）。
    """
    if bf.empty:
        return None
    widest = max(bf.rows or [bf], key=lambda r: r.total)
    return 1.0 / widest.total


def fig_rows_of(bf) -> list:
    """块的图条行（无图返回空表）。"""
    return [] if bf.empty else (bf.rows or [bf])


def cap_rows_of(task, r: int, bf) -> list[list[str]]:
    """第 r 块**逐行**的图注表；单行的老写法退化处理。"""
    rows = fig_rows_of(bf)
    caps_rows = getattr(task.sections[r], "cap_rows", None) or []
    if len(caps_rows) == len(rows):
        return caps_rows
    flat = task.captions[r]
    out, i = [], 0
    for rb in rows:
        out.append(flat[i:i + len(rb.paths)])
        i += len(rb.paths)
    return out


def grid_height(ihs: list[float], caphs: list[float]) -> float:
    """图网格的总高：Σ行(图高 + 图注 + 间距) + 行间缝。"""
    n = len(ihs)
    return (sum(ihs) + sum(caphs) + n * C.CAP_LINEGAP
            + max(0, n - 1) * C.FIG_ROW_GAP)


def fig_gap(bf) -> float:
    """图条内部的张间总间距。"""
    return max(0, len(bf.paths) - 1) * C.FIG_GAP_IN


def cap_line_height(cap_sz: float) -> float:
    """图注单行高（cm）。

    取「真实行高（= 图注字号 x :data:`config.CAP_LINE_FACTOR`）」与下限
    :data:`config.CAP_H` 的较大者 —— 图注字号是跟着正文字号浮动的，
    预留高度也必须跟着走，否则大字号时图注会溢出自己声明的框。
    """
    return max(C.CAP_H, cap_sz * C.PT2CM * C.CAP_LINE_FACTOR)


def caption_block_height(caps, ih: float, spans, cap_sz: float) -> float:
    """图注区高（cm）。

    一张图一条图注，图注**居中在自己那张图的正下方**，宽度就是该图的宽度，
    所以窄图的图注可能折行。取各张图注行数的最大值 x 单行高。
    """
    n = 1
    for cap, span in zip(caps, spans):
        if not cap:
            continue
        w = max(0.25, ih * span)
        n = max(n, line_count([(cap, "gray")], cap_sz, w),
                line_count([(cap, "gray")], cap_sz, w, "hang"))
    return n * cap_line_height(cap_sz)


# ---------------------------------------------------------------- 排法分派


def pick_mode(ratio: float | None) -> str:
    """按图条的 h/w 决定图文排法；``None`` 表示没有配图。"""
    if ratio is None:
        return "text"
    if ratio < 1 / 3.0:            # w/h > 3.0 → 超宽条
        return "stack"
    if ratio >= 1 / 1.20:          # 方图 → 左文右图
        return "side"
    return "auto"                  # 中间形态，由块宽再定


def base_modes(task, figs) -> list[str]:
    """每块的排法：任务文档显式指定优先，否则由图条的宽高比决定。"""
    out = []
    for sec, bf in zip(task.sections, figs):
        if sec.mode != "auto":
            out.append(sec.mode)
        else:
            out.append(pick_mode(strip_ratio(bf)))
    return out


def _row_index(r: int, rows: list[list[int]]) -> int:
    for i, idxs in enumerate(rows):
        if r in idxs:
            return i
    raise ValueError(f"正文段 {r} 不在任何行内")


# ---------------------------------------------------------------- 单块求解


def _cell(r: int, bw: float, body: float, lh: float, ih_cap: float,
          figs, modes: list[str], task) -> Cell | None:
    """算单块的排法与图条尺寸（尚未做行内补白）。"""
    inner = bw - 2 * C.BP_X
    bm = modes[r]
    n_full = len(task.rows[_row_index(r, task.rows)]) == 1
    bf = figs[r]
    grid = (not bf.empty) and len(bf.rows or []) > 1     # 块内多行图网格
    mode = ("text" if bf.empty
            else "stack" if grid and bm == "auto"        # 图网格一律「文上图网格下」
            else ("side" if n_full else "stack") if bm == "auto"
            else bm)

    if mode == "text":
        nl = lines_count_worst(task.paras[r], body, inner)
        th = nl * lh
        return Cell(r, bw, inner, "text", inner, th, nl, 0.0, 0.0, th)

    cap_sz = cap_size(body)
    rows_bf = fig_rows_of(bf)
    caps_rows = cap_rows_of(task, r, bf)
    grid = len(rows_bf) > 1

    if mode == "stack":
        tw = inner
        nl = lines_count_worst(task.paras[r], body, tw)
        th = nl * lh
        ihs, fws, caphs = [], [], []
        for k, rb in enumerate(rows_bf):
            gap_i = fig_gap(rb)
            ih_i = min((inner - gap_i) / rb.total, ih_cap)
            ihs.append(ih_i)
            fws.append(ih_i * rb.total + gap_i)
            caphs.append(caption_block_height(caps_rows[k], ih_i, rb.spans, cap_sz))
        figh = grid_height(ihs, caphs)
        ch = th + C.TEXT_IMG_GAP + figh
        return Cell(r, bw, inner, "stack", tw, th, nl, max(fws), ihs[0], ch,
                    max(caphs), ihs=ihs, caphs=caphs, figh=figh)

    # 以下按单行处理：图网格只支持 stack（文上图网格下）
    bf = rows_bf[0]
    S, gap, caps = bf.total, fig_gap(bf), caps_rows[0]

    best = None
    ih = C.SIDE_MIN
    while ih <= ih_cap + 1e-9:
        fw = ih * S + gap
        tw = inner - fw - C.COLGAP_IN
        if tw >= C.TW_MIN:
            nl = lines_count_worst(task.paras[r], body, tw)
            th = nl * lh
            caph = caption_block_height(caps, ih, bf.spans, cap_sz)
            ch = max(th, ih + C.CAP_LINEGAP + caph)
            if best is None or ch < best.ch - 1e-9:
                best = Cell(r, bw, inner, "side", tw, th, nl, fw, ih, ch, caph)
        ih = round(ih + C.STEPS, 3)
    return best


def solve(body: float, ih_cap: float, figs, modes: list[str], task):
    """给定字号与图高上限，返回 ``(总高, 行列表)``；放不下则返回 ``None``。"""
    lh = line_height(body)
    rows: list[Row] = []
    for ri, idxs in enumerate(task.rows):
        n = len(idxs)
        usable = C.INNER_W - (n - 1) * C.COL_GAP
        cells: list[Cell] = []
        for j, r in enumerate(idxs):
            c = _cell(r, usable * task.row_fracs[ri][j], body, lh, ih_cap,
                      figs, modes, task)
            if c is None:
                return None
            cells.append(c)

        h_avail = max(c.ch for c in cells)          # 行内可用内容高
        cap_sz = cap_size(body)
        for c in cells:                             # 补白：把余量补给图条
            if not c.has_fig:
                continue
            bf = figs[c.r]
            rows_bf = fig_rows_of(bf)
            caps_rows = cap_rows_of(task, c.r, bf)

            if c.style == "stack" and len(rows_bf) > 1:
                # 图网格：各行**轮流长**，谁也别长过块内宽 / 行高预算
                grew = True
                while grew:
                    grew = False
                    for k, rb in enumerate(rows_bf):
                        if c.ihs[k] + C.STEPS > ih_cap + 1e-9:
                            continue
                        trial = list(c.ihs)
                        trial[k] = round(trial[k] + C.STEPS, 3)
                        fws = [trial[i] * rows_bf[i].total + fig_gap(rows_bf[i])
                               for i in range(len(trial))]
                        if max(fws) > c.inner + 1e-9:
                            continue
                        caphs = [caption_block_height(caps_rows[i], trial[i],
                                                      rows_bf[i].spans, cap_sz)
                                 for i in range(len(trial))]
                        if (c.th + C.TEXT_IMG_GAP + grid_height(trial, caphs)
                                > h_avail + 1e-9):
                            continue
                        c.ihs[k] = trial[k]
                        grew = True
                c.figh = grid_height(c.ihs, c.caphs)
                c.fw = max(c.ihs[i] * rows_bf[i].total + fig_gap(rows_bf[i])
                           for i in range(len(rows_bf)))
                c.ih, c.caph = c.ihs[0], max(c.caphs)
                continue

            S, gap = bf.total, fig_gap(bf)
            caps = caps_rows[0]
            ih = c.ih
            while ih <= ih_cap + 1e-9:
                fw = ih * S + gap
                caph = caption_block_height(caps, ih, bf.spans, cap_sz)
                if c.style == "stack":
                    if fw > c.inner + 1e-9:
                        break
                    if c.th + C.TEXT_IMG_GAP + ih + C.CAP_LINEGAP + caph > h_avail + 1e-9:
                        break
                    c.ih, c.fw, c.caph = ih, fw, caph
                else:
                    tw = c.inner - fw - C.COLGAP_IN
                    if tw < C.TW_MIN:
                        break
                    nl = lines_count_worst(task.paras[c.r], body, tw)
                    th = nl * lh
                    if max(th, ih + C.CAP_LINEGAP + caph) > h_avail + 1e-9:
                        break
                    c.ih, c.fw, c.tw, c.th, c.nl, c.caph = ih, fw, tw, th, nl, caph
                ih = round(ih + C.STEPS, 3)
            c.ihs, c.caphs = [c.ih], [c.caph]
            c.figh = c.ih + C.CAP_LINEGAP + c.caph
        rows.append(Row(cells, C.BLOCK_OH + h_avail))
    total = sum(x.bh for x in rows) + (len(rows) - 1) * C.BGAP
    return total, rows


# ---------------------------------------------------------------- 尺寸搜索


def search(figs, modes: list[str], task, verbose: bool = True):
    """先保配图尺寸、再取最大字号。

    先扫一遍「图高上限 → 该上限下能用的最大字号」得到权衡表，
    再在字号 >= :data:`config.BODY_TARGET` 的前提下挑最大的图高上限。
    """
    bodies = [round(C.BODY_MAX - C.STEPS * i, 3)
              for i in range(int((C.BODY_MAX - C.BODY_MIN) / C.STEPS) + 1)]
    caps = [round(C.IH_MAX - C.STEPS * i, 3)
            for i in range(int((C.IH_MAX - C.IH_MIN) / C.STEPS) + 1)]

    table = []
    for cap in caps:
        fit = next((b for b in bodies
                    if (s := solve(b, cap, figs, modes, task)) and s[0] <= C.AVAIL_H),
                   None)
        table.append((cap, fit))

    if verbose:
        print("--- 图高上限 x 可行最大字号 ---")
        for cap, fit in table:
            print("  cap=%.2f -> 字号 %s" % (cap, ("%.2f" % fit) if fit else "N/A"))

    ok = [t for t in table if t[1] is not None and t[1] >= C.BODY_TARGET]
    if ok:
        ih_cap, body = ok[0]
    else:
        ok = [t for t in table if t[1] is not None]
        if not ok:
            raise SystemExit("无可行解：内容装不进内容区")
        ih_cap, body = sorted(ok, key=lambda t: (-t[1], -t[0]))[0]
    return body, ih_cap


# ---------------------------------------------------------------- 网格落位


def build_layout(task, figs, verbose: bool = True) -> Layout:
    """完整求解：字号 / 图高上限 / 每块几何 / 网格落位。"""
    modes = base_modes(task, figs)
    body, ih_cap = search(figs, modes, task, verbose=verbose)
    if verbose:
        print("=> 采用 字号 %.2fpt / 图高上限 %.2fcm" % (body, ih_cap))

    total, rows = solve(body, ih_cap, figs, modes, task)
    slack = C.AVAIL_H - total
    gap_v = C.BGAP + min(0.22, slack / max(1, len(rows) - 1))
    leftover = C.AVAIL_H - (total + (gap_v - C.BGAP) * (len(rows) - 1))
    top0 = C.Y0 + C.PAD_T + max(0.0, leftover) * 0.5

    # 横向落位
    for row in rows:
        x = C.X0 + C.PAD_X
        for c in row.cells:
            c.x = x
            x += c.bw + C.COL_GAP

    # 纵向落位
    y = top0
    for row in rows:
        row.y = y
        for c in row.cells:
            c.y = y
            c.h = row.bh
            if c.style == "stack":
                c.gh = c.th + C.TEXT_IMG_GAP + c.figh
                c.off = max(0.0, (c.ch - c.gh) / 2)
            else:
                c.gh = c.ch
                c.off = 0.0
        y += row.bh + gap_v

    return Layout(body, ih_cap, rows, total, gap_v, top0)


def describe(lay: Layout, task, figs=None) -> str:
    """构建日志用的版面摘要。"""
    n_lines = sum(c.nl for row in lay.rows for c in row.cells)
    out = [
        "正文 %.2fpt | 行距 %.3fcm | 图注 %.2fpt | 图高上限 %.2fcm | 共 %d 行 / %d 字"
        % (lay.body_pt, line_height(lay.body_pt), cap_size(lay.body_pt),
           lay.ih_cap, n_lines, task.char_count()),
        "行高: %s（合计 %.2f，预算 %.2f，行距 %.3f）"
        % (["%.2f" % r.bh for r in lay.rows], lay.total, C.AVAIL_H, lay.gap_v),
    ]
    for row in lay.rows:
        for c in row.cells:
            if not c.has_fig:
                fig = "无图"
            elif len(c.ihs) > 1:
                fig = "网格 %d 行 %s" % (len(c.ihs),
                    "/".join("%.2f" % h for h in c.ihs))
            else:
                n = len(figs[c.r].paths) if figs else 0
                fig = "图条 %dx %.2f x %.2f" % (n, c.fw, c.ih)
            out.append("  #%d %-5s 块宽 %.2f 正文宽 %.2f -> %2d 行(%.2fcm) %-17s 块高 %.2f"
                       % (c.r, c.style, c.bw, c.tw, c.nl, c.th, fig, row.bh))
    return "\n".join(out)


def cap_size(body: float) -> float:
    """图注字号。"""
    return max(2.6, round(body * 0.90, 1))
