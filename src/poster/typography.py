# -*- coding: utf-8 -*-
"""排版量算：字体缓存、分词、断行。

用 Pillow 的 ``getlength`` 做像素级宽度量算，而不是靠字符数估算 ——
这样才能处理中英混排、数字/单位不被拆开、标点避头尾。

⚠️ 断行策略必须**跟渲染器一致**，否则会少算出行数、末行被下面的图盖住
   （2026-10-14 实际发生的版面事故，见 docs/rendering-notes.md）。
   PowerPoint / WPS 遇到行尾放不下的标点时，是把标点**推到下一行**（``push``），
   不是把标点挤在行尾（``hang``）。所以默认策略是 ``push``。
"""
from __future__ import annotations

import re

from PIL import ImageFont

from .config import (
    BLACK,
    FONT,
    FONT_BOLD,
    LINE_FACTOR,
    LINE_REF_FACTOR,
    MEASURE_SAFETY,
    PT2CM,
    PX_PER_CM,
)

# 不可置于行首的标点（避头尾）
NO_START = set("，。、；：？！）》」』】…·,.;:?!)%")

# 数字 / 单位 / 化学式视为整体，不拆开
_TOKEN_RE = re.compile(r"[\-+/~≈°]?[0-9][0-9A-Za-z.%\-+/°~]*|[A-Za-z][A-Za-z0-9+\-]*")

_FONT_CACHE: dict[tuple[float, bool], ImageFont.FreeTypeFont] = {}

# 标点落在行尾、放不下时的两套做法（见模块 docstring）
POLICIES = ("push", "hang")


def font(size: float, bold: bool = False) -> ImageFont.FreeTypeFont:
    """按 pt 字号取（带缓存的）字体对象。"""
    key = (round(size, 3), bold)
    if key not in _FONT_CACHE:
        px = max(1, int(round(size * PT2CM * PX_PER_CM)))
        path = FONT_BOLD if bold else FONT
        try:
            _FONT_CACHE[key] = ImageFont.truetype(str(path), px)
        except OSError:                     # 字体缺失时退化为 Pillow 内置字体
            _FONT_CACHE[key] = ImageFont.load_default()
    return _FONT_CACHE[key]


def unpack(seg) -> tuple[str, tuple[int, int, int], str]:
    """把运行段统一成 ``(文本, RGB, 角标)``。

    运行段允许写成二元组 ``(文本, RGB)``（普通文字）或三元组
    ``(文本, RGB, 角标)``；``角标`` 取 ``""`` / ``"sup"`` / ``"sub"``。
    二元组是历史写法，一律按普通文字处理，便于渐进改造。
    """
    if len(seg) == 2:
        return seg[0], seg[1], ""
    return seg[0], seg[1], seg[2]


def tokenize(segments) -> list[tuple[str, tuple[int, int, int], str]]:
    """把 ``[(文本, 颜色[, 角标]), ...]`` 切成不可再分的排版单元。

    角标**不改变字宽**：上/下标只是把同一字号的字形整体抬升或下沉
    （PowerPoint 的 ``baseline`` 只做位移，不缩放），所以求解器量宽度时
    无需区分角标，断行结果与渲染器一致。
    """
    out: list[tuple[str, tuple[int, int, int], str]] = []
    for seg in segments:
        text, color, script = unpack(seg)
        i, n = 0, len(text)
        while i < n:
            m = _TOKEN_RE.match(text, i)
            if m:
                out.append((m.group(0), color, script))
                i = m.end()
            else:
                out.append((text[i], color, script))
                i += 1
    return out


def wrap_para(segments, size: float, width_cm: float, policy: str = "push"):
    """按 ``width_cm`` 断行，返回 ``[(行内单元列表, 行宽px), ...]``。

    ``policy`` 决定行尾放不下的标点怎么办：

    * ``"push"``（默认）—— 推到下一行。这就是 PowerPoint / WPS 的做法，
      因此引擎的**行数估算**必须用它，否则会少算一行、末行被图盖住。
    * ``"hang"`` —— 挤在本行末尾（行尾悬挂）。版式上更好看，但除非在
      ``bodyPr`` 上真的开了悬挂标点，渲染器不会这么做，只可用于预览对比。
    """
    if policy not in POLICIES:
        raise ValueError("未知的断行策略 %r（可选 %s）" % (policy, "、".join(POLICIES)))
    maxw = width_cm * PX_PER_CM
    f = font(size)
    lines: list[tuple[list, float]] = []
    cur: list = []
    curw = 0.0

    for tok, color, script in tokenize(segments):
        w = f.getlength(tok)
        if cur and curw + w > maxw:
            if policy == "hang" and tok and tok[0] in NO_START:
                cur.append((tok, color, script))    # 标点不落行首：挤在本行末尾
                curw += w
                lines.append((cur, curw))
                cur, curw = [], 0.0
                continue
            lines.append((cur, curw))           # 推标点：本行到这里为止
            cur, curw = [], 0.0
        if w > maxw:                            # 单个超宽单元：逐字拆
            for ch in tok:
                cw = f.getlength(ch)
                if curw + cw > maxw and cur:
                    lines.append((cur, curw))
                    cur, curw = [], 0.0
                cur.append((ch, color, script))
                curw += cw
        else:
            cur.append((tok, color, script))
            curw += w
    if cur:
        lines.append((cur, curw))

    fixed = []
    for ln, lw in lines:                        # 去掉行尾空白
        while ln and ln[-1][0].strip() == "":
            lw -= f.getlength(ln[-1][0])
            ln = ln[:-1]
        fixed.append((ln if ln else [(" ", BLACK, "")], lw))
    return fixed


def line_count(segments, size: float, width_cm: float, policy: str = "push") -> int:
    """断行后的行数；量算宽度打 :data:`MEASURE_SAFETY` 折以对齐渲染器。"""
    return len(wrap_para(segments, size, width_cm * MEASURE_SAFETY, policy))


def lines_count(paras, size: float, width_cm: float, policy: str = "push") -> int:
    """一个分块内**多个段落**的合计行数。

    段间距为 0（见 writer.add_text），所以块高 = 合计行数 x 行高。
    """
    return sum(line_count(p, size, width_cm, policy) for p in paras)


def lines_count_worst(paras, size: float, width_cm: float) -> int:
    """两种标点策略里更悲观的那个行数 —— 供合规自检用它兜住渲染器差异。"""
    return max(lines_count(paras, size, width_cm, p) for p in POLICIES)


def line_height(size: float) -> float:
    """单行高（cm）—— 就是 :func:`line_spacing_pct` 落进 pptx 后的真实行距。"""
    return size * PT2CM * LINE_FACTOR


def line_spacing_pct(factor: float = LINE_FACTOR) -> int:
    """把「行距倍数」换算成 pptx 要写的 ``spcPct`` 千分数。

    为什么不用 ``spcPts``
    --------------------
    PowerPoint 会把**绝对点值行距四舍五入到整点**：写 ``spcPts 552``（5.52pt）
    实际按 6.0pt 排 —— 一行多 8.7% 高度，整块多出一行，末行就被下面的图盖住了
    （2026-10-14 的版面事故）。而 ``spcPct`` 是精确的：
    实测 100% = 字号 x 1.20（PowerPoint 的「单倍行距」），133% = 5.505pt，
    与 3.45 x 1.20 x 1.33 = 5.506pt 完全吻合。
    """
    return int(round(factor / LINE_REF_FACTOR * 100000))
