# -*- coding: utf-8 -*-
"""排版量算：字体缓存、分词、断行。

用 Pillow 的 ``getlength`` 做像素级宽度量算，而不是靠字符数估算 ——
这样才能处理中英混排、数字/单位不被拆开、标点避头尾。

⚠️ 断行策略必须**跟渲染器一致**，否则会少算出行数、末行被下面的图盖住
   （2026-10-14 实际发生的版面事故，见 docs/rendering-notes.md）。

   行尾标点放不下时，渲染器可能有**三种**做法，本模块全部实现：

   * ``push`` —— 标点单独推到下一行（下一行会以标点开头）；
   * ``hang`` —— 标点挤出边界、悬挂在行外（``bodyPr/@hangingPunct="1"`` 时的行为）；
   * ``pull`` —— **行首禁则**：标点不能落在行首，于是把它连同**前一个单元**
     一起移到下一行。这是 PowerPoint 在"悬挂标点挤不下"时的真实兜底做法，
     行数可能比前两者都多。

   **求解尺寸一律取三者中最悲观的行数**（见 :func:`lines_count_worst`）。
   只看 ``push``/``hang`` 会漏掉 ``pull`` 那一行 —— 表现为"估算 10 行、
   实排 11 行"，末行压到图上（2026-10-16 在传感块复发过一次）。
"""
from __future__ import annotations

import re

from PIL import ImageFont

from . import config as _cfg
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

# 不可置于**行尾**的字符：左括号不能孤零零挂在行尾（"…（" 换行读不通）
NO_END = set("([{（【《「『")

# 数字 / 单位 / 化学式视为整体，不拆开
_TOKEN_RE = re.compile(r"[\-+/~≈°]?[0-9][0-9A-Za-z.%\-+/°~]*|[A-Za-z][A-Za-z0-9+\-]*")

_FONT_CACHE: dict[tuple[float, bool], ImageFont.FreeTypeFont] = {}

# 标点落在行尾、放不下时的几种做法（见模块 docstring）
# 「求解尺寸」取三者中最悲观的那一个，见 lines_count_worst
POLICIES = ("push", "hang", "pull")


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


def _spc_px() -> float:
    """量宽用的**每字符字距**（px）—— 实写字距放大 :data:`CHAR_SPC_MEASURE_SCALE` 倍。

    量宽时按悲观估计（放大）而不是照抄实排值：字距是绝对量、不随字号缩放，
    再叠加 Pillow 与 PowerPoint 之间那点固有偏差，1x 计字距仍会偶尔少算一行
    （2026-10-16 实测：窄栏块实排 8 行、模型 7 行，末行压到图上）。
    写进 pptx 的仍是 :data:`CHAR_SPC_PT` 本身（见 writer.add_text）。
    """
    return (_cfg.CHAR_SPC_PT * _cfg.CHAR_SPC_MEASURE_SCALE
            * _cfg.PT2CM * _cfg.PX_PER_CM)


def wrap_para(segments, size: float, width_cm: float, policy: str = "push"):
    """按 ``width_cm`` 断行，返回 ``[(行内单元列表, 行宽px), ...]``。

    ``policy`` 决定行尾放不下的标点怎么办：

    * ``"push"`` —— 推到下一行（下一行会以标点开头）；
    * ``"hang"`` —— 挤在本行末尾、溢出边界（行尾悬挂）；渲染器上要真开
      ``bodyPr/@hangingPunct`` 才会这么做，只可用于预览对比；
    * ``"pull"``（默认）—— **行首禁则**：标点不落行首，连同前一个单元一起
      移到下一行。这是 PowerPoint 在悬挂挤不下时的真实兜底，也是三者中
      行数最多的一个，**求解尺寸必须按它兜底**。
    """
    if policy not in POLICIES:
        raise ValueError("未知的断行策略 %r（可选 %s）" % (policy, "、".join(POLICIES)))
    maxw = width_cm * PX_PER_CM
    f = font(size)
    spc = _spc_px()                     # 每字符额外字距：量宽必须与渲染器同口径
    lines: list[tuple[list, float]] = []
    cur: list = []
    curw = 0.0

    def _tw(toks) -> float:
        """一组排版单元的实排宽度（含字符间距）。"""
        return sum(f.getlength(t[0]) + len(t[0]) * spc for t in toks)

    for tok, color, script in tokenize(segments):
        w = f.getlength(tok) + len(tok) * spc
        # ``not script``：角标（``^{}``/``_{}``）必须紧跟它的基字 —— 不许在
        # 基字与角标之间断行，否则会排出"…0.5M A / 2+在三者间"这种断法
        if cur and curw + w > maxw and not script:
            # 行尾禁则：左括号之类不能收尾 —— 连同它后面的内容一起挪到下一行
            tail: list = []
            while len(cur) >= 2 and cur[-1][0] and cur[-1][0][-1] in NO_END:
                last = cur.pop()
                curw -= f.getlength(last[0]) + len(last[0]) * spc
                tail.insert(0, last)
            nxt: list = list(tail)
            if policy == "hang" and tok and tok[0] in NO_START:
                cur.append((tok, color, script))    # 标点不落行首：挤在本行末尾
                curw += w
                lines.append((cur, curw))
                cur, curw = nxt, _tw(nxt)
                continue                            # 本 token 已处理，别再追加一次
            if policy == "pull" and tok and tok[0] in NO_START and len(cur) >= 2:
                # 行首禁则：标点连同**前一个单元**一起移到下一行。
                # 本行为此少放一个单元，总行数可能比 push / hang 多一行。
                last = cur.pop()
                lines.append((cur, curw - f.getlength(last[0]) - len(last[0]) * spc))
                nxt = [last] + nxt
            else:
                lines.append((cur, curw))           # 推标点：本行到这里为止
            cur, curw = nxt, _tw(nxt)
        if w > maxw:                            # 单个超宽单元：逐字拆
            for ch in tok:
                cw = f.getlength(ch) + spc
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
            lw -= f.getlength(ln[-1][0]) + len(ln[-1][0]) * spc
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
