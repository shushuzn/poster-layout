# -*- coding: utf-8 -*-
"""合规自检：把「这份墙报算不算合格」变成一份可执行、可复现的清单。

为什么要有它
------------
版式与内容的要求散落在模板说明、样张、评审通知和历次返工里，靠人记忆一定会漏
—— 本仓库就漏过两次：一次图文错配（一张表征图挂到了不是讨论它的那个块下），
一次多张图共用一条图注。两次都不是能力问题，是**没有清单**。

所以每条要求都落成一个带编号的 :class:`Item`：

* 能自动判的 → 挂一个 ``check`` 函数，构建完直接跑；
* 判不了的（例如"这张图在讲什么、该配哪段正文"）→ 标 ``manual``，报告里列出来等人认，
  并把**判读依据**写在条目里，免得下次又认反。

清单正文写在 ``docs/compliance.md``。``tests/test_compliance.py`` 会校验
"文档里的编号" 与 "代码里的编号" 完全一致 —— 加了检查不写文档，
或写了文档不实现，测试都会红，两边不会偷偷脱节。

用法::

    python build.py --check          # 构建 + 自检 + 写报告到 dist/
    python build.py --check -q       # 只要结论与退出码（红了会打印问题）
"""
from __future__ import annotations

import hashlib
import re
from dataclasses import dataclass, field
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn

from . import config as C
from .layout import cap_size
from .template import is_placeholder
from .typography import POLICIES, line_count

# 禁用字形：微软雅黑缺这些上下标 / 单位符号的字形，上版后会变成空心方框（豆腐块）。
# 正文一律用 ASCII 等价写法：°C / m2 / at% / ID/IG / Fe3+。
BAD_GLYPHS = set("℃₀₁₂₃₄₅₆₇₈₉ₓ⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼₊₋")
# 未完成的痕迹
STUB_WORDS = ("TODO", "todo", "待补", "待填", "待定", "xxx", "XXX", "（无配图）", "正文……")
TOL = 0.03                      # 几何比对容差（cm）
_CAP_RE = re.compile(r"图\s*(\d+)")
_HAN_RE = re.compile(r"[\u4e00-\u9fff]")        # 字数口径：汉字（不含数字/英文/标点）


def _sha(p: Path) -> str:
    return hashlib.sha256(Path(p).read_bytes()).hexdigest()


# ---------------------------------------------------------------- 从 pptx 读回
def _cm(v) -> float:
    return round(v / 360000.0, 4)


def _grad(sh):
    """返回 ``(起点色, 终点色, 角度)``；不是渐变则 ``(None, None, None)``。"""
    spPr = sh._element.spPr                                    # noqa: SLF001
    gf = spPr.find(qn("a:gradFill"))
    if gf is None:
        return None, None, None
    stops = []
    gs_lst = gf.find(qn("a:gsLst"))
    for gs in (gs_lst.findall(qn("a:gs")) if gs_lst is not None else []):
        clr = gs.find(qn("a:srgbClr"))
        stops.append(clr.get("val").upper() if clr is not None else None)
    lin = gf.find(qn("a:lin"))
    ang = int(lin.get("ang")) if lin is not None else None
    c1 = stops[0] if stops else None
    c2 = stops[-1] if stops else None
    return c1, c2, ang


def _dash(sh):
    """返回线的虚线样式（如 ``'dash'``）；实线返回 ``None``。"""
    ln = sh._element.spPr.find(qn("a:ln"))                     # noqa: SLF001
    if ln is None:
        return None
    pd = ln.find(qn("a:prstDash"))
    return None if pd is None else pd.get("val")


def _rgb(t: tuple[int, int, int]) -> str:
    return "%02X%02X%02X" % t


@dataclass
class Ctx:
    """一次自检要用到的全部事实（构建产物 + 任务文档 + 求解结果）。"""

    task: object
    figs: list
    lay: object
    out: Path
    template_texts: list[str] = field(default_factory=list)
    template_pics: int = 0
    n_tpl: int = 0                                  # 模板原有形状数；前这么多个不归引擎管
    blocks: list = field(default_factory=list)      # (x,y,w,h,sh,idx) 虚线块底
    autos: list = field(default_factory=list)       # 所有自选图形
    links: list = field(default_factory=list)       # (x,y,w,sh,idx) 连接线
    pics: list = field(default_factory=list)        # (x,y,w,h,sh,idx) 图片
    texts: list = field(default_factory=list)       # (x,y,w,h,text,idx) 文本框

    @property
    def cap_sz(self) -> float:
        return cap_size(self.lay.body_pt)

    def block_of(self, c):
        """按几何找到某块对应的虚线框形状。"""
        for x, y, w, h, sh, _i in self.blocks:
            if (abs(x - c.x) < TOL and abs(y - c.y) < TOL
                    and abs(w - c.bw) < 0.05 and abs(h - c.h) < 0.05):
                return sh
        return None

    def chip_rec_of(self, title: str):
        """取胶囊小标题那条记录 ``(x, y, w, sh, idx)``（要拿 idx 判叠序）。"""
        for rec in self.autos:
            if rec[4].has_text_frame and rec[4].text_frame.text.strip() == title:
                return rec
        return None

    def chip_of(self, title: str):
        rec = self.chip_rec_of(title)
        return None if rec is None else rec[4]

    def pics_in(self, c):
        return [p for p in self.pics
                if c.x - 0.02 <= p[0] <= c.x + c.bw + 0.02
                and c.y - 0.02 <= p[1] <= c.y + c.h + 0.02]

    def texts_in(self, c):
        return [t for t in self.texts
                if c.x - 0.02 <= t[0] <= c.x + c.bw + 0.02
                and c.y - 0.02 <= t[1] <= c.y + c.h + 0.02]


def load(task, figs, lay, out_path) -> Ctx:
    """把产物 pptx 读回成可校验的事实集合。"""
    ctx = Ctx(task=task, figs=figs, lay=lay, out=Path(out_path))

    tpl = Presentation(str(task.template))
    ctx.n_tpl = len(tpl.slides[0].shapes)
    ctx.template_pics = sum(1 for sh in tpl.slides[0].shapes
                            if sh.shape_type == MSO_SHAPE_TYPE.PICTURE)
    ctx.template_texts = [sh.text_frame.text.strip() for sh in tpl.slides[0].shapes
                          if sh.has_text_frame and sh.text_frame.text.strip()
                          and not is_placeholder(sh)]

    for i, sh in enumerate(Presentation(str(out_path)).slides[0].shapes):
        st = sh.shape_type
        if st == MSO_SHAPE_TYPE.PICTURE:
            ctx.pics.append((_cm(sh.left), _cm(sh.top), _cm(sh.width), _cm(sh.height), sh, i))
        elif st == MSO_SHAPE_TYPE.LINE:
            ctx.links.append((_cm(sh.left), _cm(sh.top), _cm(sh.width), sh, i))
        elif st == MSO_SHAPE_TYPE.TEXT_BOX:
            ctx.texts.append((_cm(sh.left), _cm(sh.top), _cm(sh.width), _cm(sh.height),
                              sh.text_frame.text, i))
        elif st == MSO_SHAPE_TYPE.AUTO_SHAPE:
            rec = (_cm(sh.left), _cm(sh.top), _cm(sh.width), _cm(sh.height), sh, i)
            ctx.autos.append(rec)
            if _dash(sh) == "dash" and sh.width > 0 and sh.height > 0:
                ctx.blocks.append(rec)
    return ctx


# ---------------------------------------------------------------- 条目定义
@dataclass
class Item:
    id: str
    group: str
    title: str
    source: str                 # 出处：模板 / 样张 / 评审 / 工程约定
    level: str                  # hard（硬性） / soft（提示） / manual（人工）
    func: object = None         # callable(ctx) -> (ok, detail)


@dataclass
class Result:
    item: Item
    ok: bool | None             # None = 人工项
    detail: str = ""


def _norm_color(s) -> str:
    return s.upper() if isinstance(s, str) else "--"


# ---- 范围 --------------------------------------------------------
def _a1(ctx):
    joined = "\n".join(t[4] for t in ctx.texts)
    missing = [t for t in ctx.template_texts if t.replace("\n", "") not in joined.replace("\n", "")]
    return (not missing,
            "模板 %d 个既有字段原样保留" % len(ctx.template_texts) if not missing
            else "模板字段被改动/删除: %s" % missing)


def _a2(ctx):
    want = ctx.template_pics + sum(len(f.paths) for f in ctx.figs)
    got = len(ctx.pics)
    return got == want, "产物图片 %d 张 = 模板背景 %d + 配图 %d" % (
        got, ctx.template_pics, sum(len(f.paths) for f in ctx.figs)) if got == want \
        else "图片数不符：期望 %d（背景 %d + 配图 %d），实得 %d" % (
            want, ctx.template_pics, sum(len(f.paths) for f in ctx.figs), got)


# ---- 尺寸 --------------------------------------------------------
def _b1(ctx):
    a = Presentation(str(ctx.task.template)).slide_width, Presentation(
        str(ctx.task.template)).slide_height
    b = Presentation(str(ctx.out)).slide_width, Presentation(str(ctx.out)).slide_height
    cm = lambda v: round(v / 360000.0, 2)                              # noqa: E731
    same = a == b
    return same, "画布 %.2f x %.2f cm（1:10 缩放稿，实印 %.0f x %.0f cm）" % (
        cm(b[0]), cm(b[1]), cm(b[0]) * 10, cm(b[1]) * 10) if same \
        else "画布被改：模板 %.2fx%.2f -> 产物 %.2fx%.2f" % (
            cm(a[0]), cm(a[1]), cm(b[0]), cm(b[1]))


def _b2(ctx):
    sz = ctx.lay.body_pt
    ok = sz >= C.BODY_MIN - 1e-9
    return ok, "正文 %.2fpt（实印 %.1fpt）≥ 下限 %.2fpt，图注 %.2fpt" % (
        sz, sz * 10, C.BODY_MIN, ctx.cap_sz)


def _b3(ctx):
    ok = ctx.lay.total <= C.AVAIL_H + 1e-6
    return ok, "总高 %.2f ≤ 预算 %.2fcm（余 %.2f）" % (
        ctx.lay.total, C.AVAIL_H, C.AVAIL_H - ctx.lay.total)


def _b4(ctx):
    bad = ["#%d" % c.r for c in ctx.lay.cells
           if (c.x < C.X0 - 1e-6 or c.x + c.bw > C.X1 + 1e-6
               or c.y < C.Y0 - 1e-6 or c.y + c.h > C.Y1 + 1e-6)]
    area = "(%.2f, %.2f) -> (%.2f, %.2f)" % (C.X0, C.Y0, C.X1, C.Y1)
    return (not bad, "%d 个分块都在内容区 %s 内" % (len(ctx.lay.cells), area) if not bad
            else "越界: %s" % "、".join(bad))


def _b5(ctx):
    bad = ["#%d %.2fcm" % (c.r, c.tw) for c in ctx.lay.cells if c.tw < C.TW_MIN - 1e-6]
    return (not bad, "各块正文列宽 ≥ %.2fcm（最窄 %.2fcm）" % (
        C.TW_MIN, min(c.tw for c in ctx.lay.cells)) if not bad
        else "正文被挤成细长条: %s" % "、".join(bad))


def _b6(ctx):
    """图高下限。

    注意与 ``config.IH_MIN``（求解器搜索下界 0.90cm）区分：那条是"引擎愿意试多小"，
    本条是"多小算不合格"。历史事故是 v15 把 5.46:1 的流程条塞进半栏、压成 0.46cm
    的一条细缝 —— 实印 4.6cm、坐标轴完全读不清，被用户打回。0.70cm（实印 7cm）
    是"还能看清坐标轴"的经验线。
    """
    lim = C.IH_MIN_READABLE
    # 块内可能是多行图网格，每一行都有自己的高，**逐行**判
    bad = ["#%d 行%d %.2fcm" % (c.r, k + 1, h) for c in ctx.lay.cells if c.has_fig
           for k, h in enumerate(c.ihs) if h < lim - 1e-9]
    figs = [(h, c) for c in ctx.lay.cells if c.has_fig for h in c.ihs]
    if not figs:
        return True, "本任务无配图"
    lo_h, lo_c = min(figs, key=lambda kv: kv[0])
    detail = "最小图高 %.2fcm（#%d，实印 %.1fcm）≥ 可读下限 %.2fcm" % (
        lo_h, lo_c.r, lo_h * 10, lim)
    return (not bad, detail if not bad
            else "配图被压成细缝（实印 <%.0fcm，读不清坐标轴）: %s" % (lim * 10, "、".join(bad)))


# ---- 版式 --------------------------------------------------------
def _c1(ctx):
    bad = []
    for c in ctx.lay.cells:
        sh = ctx.block_of(c)
        if sh is None or _dash(sh) != "dash":
            bad.append("#%d" % c.r)
    return (not bad, "%d 个分块底色均为虚线圆角框" % len(ctx.lay.cells) if not bad
            else "缺少虚线框: %s" % "、".join(bad))


def _c2(ctx):
    want = (_rgb(C.BLOCK_C1), _rgb(C.BLOCK_C2))
    bad = []
    for c in ctx.lay.cells:
        sh = ctx.block_of(c)
        c1, c2, _ = _grad(sh) if sh is not None else (None, None, None)
        if (c1, c2) != want:
            bad.append("#%d %s->%s" % (c.r, _norm_color(c1), _norm_color(c2)))
    return (not bad, "块底渐变 #%s -> #%s" % want if not bad
            else "块底渐变不符: %s" % "、".join(bad))


def _c3(ctx):
    bad = []
    for c in ctx.lay.cells:
        title = ctx.task.titles[c.r]
        sh = ctx.chip_of(title)
        if sh is None:
            bad.append("#%d 缺胶囊" % c.r)
            continue
        c1, c2, _ = _grad(sh)
        if (c1, c2) != (_rgb(C.CHIP_C1), _rgb(C.CHIP_C2)):
            bad.append("#%d 胶囊配色不符" % c.r)
            continue
        runs = sh.text_frame.paragraphs[0].runs
        r = runs[0] if runs else None
        if r is None or not r.font.bold:
            bad.append("#%d 标题未加粗" % c.r)
        elif r.font.color is None or str(r.font.color.rgb) != "FFFFFF":
            bad.append("#%d 标题非白字" % c.r)
    return (not bad, "%d 个胶囊小标题（全圆角 + 蓝渐变 + 白字粗体）" % len(ctx.lay.cells)
            if not bad else "胶囊问题: %s" % "、".join(bad))


def _c4(ctx):
    bad = []
    for c in ctx.lay.cells:
        y_want = c.y + C.TITLE_TOP + C.CHIP_H / 2
        hit = [lk for lk in ctx.links
               if abs(lk[1] - y_want) < 0.02
               and abs(lk[0] - (c.x + C.BP_X)) < 0.03
               and abs((lk[0] + lk[2]) - (c.x + c.bw - C.BP_X)) < 0.03]
        if not hit:
            bad.append("#%d 无骑线虚线" % c.r)
            continue
        chip = ctx.chip_rec_of(ctx.task.titles[c.r])
        if chip is not None and not (hit[0][4] < chip[5]):
            bad.append("#%d 虚线与胶囊叠序反了" % c.r)
    return (not bad, "虚线整宽横铺、胶囊骑其上（%d 条）" % len(ctx.lay.cells) if not bad
            else "骑线问题: %s" % "、".join(bad))


def _c5(ctx):
    """版面有节奏，不是千行一面。

    两种形态都算成立：混排网格（既有通栏行也有并排行），或
    全通栏图墙但**各行配图张数不同**（示例同款：每行一条满宽图条，
    节奏由「文+双图 / 文+三方图 / 文+三图 / 纯文本」交替给出）。
    """
    ns = [len(r) for r in ctx.task.rows]
    figs_row = [sum(len(ctx.figs[r].paths) for r in row) for row in ctx.task.rows]
    mixed = max(ns) >= 2 and min(ns) == 1
    rhythm = len(set(figs_row)) >= 2
    if mixed or rhythm:
        how = "混排" if mixed else "各行配图 %s" % figs_row
        return True, "网格 %s：版面有节奏（%s）" % (ns, how)
    return False, "网格 %s 千行一面：既无混排，各行配图数也相同" % ns


def _c6(ctx):
    bad = []
    for c in ctx.lay.cells:
        if not c.has_fig:
            continue
        sec = ctx.task.sections[c.r]
        body = "".join(t for para in sec.paras for t, _, _ in para)
        if not [t for t in ctx.texts_in(c) if body[:16] in t[4].replace("\n", "")]:
            bad.append("#%d「%s」块内找不到与图同框的正文" % (c.r, sec.title))
    return (not bad, "有图的 %d 个分块内都同时有正文（图文同框）" % sum(
        1 for c in ctx.lay.cells if c.has_fig) if not bad else "图文不同框: %s" % "、".join(bad))


def _c7(ctx):
    bad = []
    for c in ctx.lay.cells:
        ps = ctx.pics_in(c)
        for x, y, w, h, text, _i in ctx.texts_in(c):
            for px, py, pw, ph, _sh, _j in ps:
                ox = min(x + w, px + pw) - max(x, px)
                oy = min(y + h, py + ph) - max(y, py)
                if ox > 0.02 and oy > 0.02:
                    bad.append("#%d 正文压到图" % c.r)
                    break
    return (not bad, "正文与配图无重叠" if not bad else "重叠: %s" % "、".join(sorted(set(bad))))


# ---- 渲染器一致性 -------------------------------------------------
# C7 只看**声明的框**有没有和图标几何重叠 —— 如果引擎把文字需要的高度算少了
# （少算一行），声明框看着没问题，可渲染器会把多出来的那行画到图上。
# 下面 C8~C11 就是补这个洞：把"渲染器会怎么排"重新算一遍，跟写进 pptx 的值对账。

def _autofit_tags(txBody) -> list[str]:
    bp = txBody.find(qn("a:bodyPr"))
    if bp is None:
        return []
    return [t for t in ("a:noAutofit", "a:normAutofit", "a:spAutoFit")
            if bp.find(qn(t)) is not None]


def _hanging(bp) -> bool:
    return bp is not None and bp.get("hangingPunct") in ("1", "true")


def text_is_ours(ctx, idx: int) -> bool:
    """第 ``idx`` 个形状是不是本引擎排进去的（前 ``ctx.n_tpl`` 个来自模板）。"""
    return idx >= ctx.n_tpl


def _text_shapes(ctx, ours_only: bool = True, textbox_only: bool = False):
    """产物里所有"带文字"的形状（正文文本框 + 胶囊小标题）。

    ``ours_only`` 只留本引擎新加的：写 pptx 是**只追加**的，模板原有形状（作品编号 /
    本科生组 / 作品名称 / 项目成员）永远排在前面、且要求一字不动 ——
    它们用不用 spcPct、有没有悬挂标点不归我们管，别去报它们。
    ``textbox_only`` 再窄一层到"由引擎断行的文本框"（胶囊是单行装饰，不参与断行）。
    """
    out = []
    for i, sh in enumerate(Presentation(str(ctx.out)).slides[0].shapes):
        if ours_only and not text_is_ours(ctx, i):
            continue
        if textbox_only and sh.shape_type != MSO_SHAPE_TYPE.TEXT_BOX:
            continue
        if sh.has_text_frame and sh.text_frame.text.strip():
            out.append(sh)
    return out


def _size_and_lnspc(sh):
    """读回：最大字号(pt)、(spcPct 或 None)、(spcPts 或 None)。

    行距只认**第一段**：本引擎每个文本框要么单段，要么所有段一致。
    """
    txBody = sh.text_frame._txBody                             # noqa: SLF001
    size = 0.0
    for rPr in txBody.iter(qn("a:rPr")):
        sz = rPr.get("sz")
        if sz:
            size = max(size, float(sz) / 100.0)
    pct = pts = None
    p = txBody.find(qn("a:p"))
    pPr = None if p is None else p.find(qn("a:pPr"))
    ln = None if pPr is None else pPr.find(qn("a:lnSpc"))
    if ln is not None:
        e = ln.find(qn("a:spcPct"))
        if e is not None:
            pct = int(e.get("val"))
        e = ln.find(qn("a:spcPts"))
        if e is not None:
            pts = int(e.get("val"))
    return size, pct, pts


def _pitch_cm(size: float, pct, pts) -> float:
    """该文本框的真实单行高（cm）—— 完全照渲染器的口径算。

    spcPct 精确：行距 = 字号 x 1.2 x 百分比；
    spcPts 会被渲染器**四舍五入到整点**（写 5.52pt 实际按 6.0pt 排），
    这里照四舍五入后的值算，宁可高估不可低估。
    """
    if pct is not None:
        return size * C.PT2CM * C.LINE_REF_FACTOR * pct / 100000.0
    if pts is not None:
        return round(pts / 100.0) * C.PT2CM
    return size * C.PT2CM * C.LINE_REF_FACTOR      # 无 lnSpc = 单倍行距


def _needed_height(text: str, size: float, width_cm: float, pitch: float) -> float:
    """这个文本框里的文字，渲染器一共要占多高（cm）。

    两种标点策略都算一遍、取更悲观的那个；宽度**直接传进去** ——
    :func:`typography.line_count` 内部已经乘过 :data:`config.MEASURE_SAFETY`，
    这里再乘一次就等于打了两次折（0.985² ≈ 0.970），
    会凭空多出整整一行，把求解器明明排得下的版面判成"溢出"。
    """
    n = 0
    for para in text.split("\n"):
        para = para.strip()
        if not para:
            continue
        n += max(line_count([(para, "x")], size, width_cm, pol)
                 for pol in POLICIES)
    return max(1, n) * pitch


def _c8(ctx):
    """文字需要的实际高度 ≤ 自己声明的框高。

    这是 2026-10-14 版面事故（末行整行被配图盖住）的直接守卫：
    当时引擎低估了一行，声明框自己符合几何，可渲染器多画了一行。

    只判**引擎排进去**的文本框 —— 模板自带字段（作品名称 / 项目成员 …）的框是
    模板作者画的，字号与框高都不由我们定，报它们只会制造假警报
    （与 C9 / C11 / F5 同一口径）。
    """
    bad, n = [], 0
    shapes = Presentation(str(ctx.out)).slides[0].shapes
    for x, y, w, h, text, i in ctx.texts:
        t = text.strip()
        if not t or not text_is_ours(ctx, i):
            continue
        n += 1
        sh = shapes[i]
        size, pct, pts = _size_and_lnspc(sh)
        pitch = _pitch_cm(size, pct, pts)
        need = _needed_height(t, size, w, pitch)
        if need > h + TOL:
            bad.append("y=%.2f 「%s…」需要 %.3fcm > 声明 %.3fcm"
                       % (y, t[:10], need, h))
    return (not bad, "%d 个文本框的声明高度都够（按最悲观断行重算）" % n if not bad
            else "文字会溢出自己的框（渲染器会把它画到框外）: %s" % "；".join(bad))


def _c9(ctx):
    """行距一律写 spcPct —— 绝对点值会被渲染器四舍五入到整点。"""
    bad = []
    for sh in _text_shapes(ctx):
        size, pct, pts = _size_and_lnspc(sh)
        if pts is not None:
            bad.append("「%s」用了 spcPts %s（会被取整成 %dpt）"
                       % (sh.text_frame.text.strip()[:8], pts, round(pts / 100)))
        elif pct is None:
            bad.append("「%s」没写行距" % sh.text_frame.text.strip()[:8])
    return (not bad, "%d 个文本框都写了 spcPct 行距（不受渲染器取整影响）"
            % len(_text_shapes(ctx)) if not bad else "行距写法有风险: %s" % "；".join(bad))


def _c10(ctx):
    """bodyPr 里自动调整是三选一：挂两个会让 PowerPoint 拒绝打开整份文件。"""
    bad = []
    for sh in Presentation(str(ctx.out)).slides[0].shapes:
        if not sh.has_text_frame:
            continue
        tags = _autofit_tags(sh.text_frame._txBody)            # noqa: SLF001
        if len(tags) > 1:
            bad.append("「%s」同时挂 %s" % (sh.text_frame.text.strip()[:8] or sh.shape_type,
                                          "+".join(t[2:] for t in tags)))
    return (not bad, "每个 bodyPr 至多一个自动调整子元素（文件能被 PowerPoint 打开）"
            if not bad else "自动调整冲突，文件可能打不开: %s" % "；".join(bad))


def _c11(ctx):
    """悬挂标点开关与断行策略必须一致。"""
    on = C.HANG_PUNCT
    shapes = _text_shapes(ctx, textbox_only=True)
    bad = []
    for sh in shapes:
        bp = sh.text_frame._txBody.find(qn("a:bodyPr"))        # noqa: SLF001
        if _hanging(bp) != on:
            bad.append("「%s」%s hangingPunct" % (sh.text_frame.text.strip()[:8],
                                                "缺" if on else "多了"))
    return (not bad, "%d 个断行文本框的悬挂标点与策略一致（HANG_PUNCT=%s，行首不会出现「，。、」）"
            % (len(shapes), on) if not bad else "悬挂标点设置不一致: %s" % "；".join(bad))


# ---- 图文 --------------------------------------------------------
def _d1(ctx):
    bad = []
    for c in ctx.lay.cells:
        want = len(ctx.figs[c.r].paths)
        got = len(ctx.pics_in(c))
        if got != want:
            bad.append("#%d「%s」声明 %d 张、实得 %d 张" % (c.r, ctx.task.titles[c.r], want, got))
    n = len(ctx.pics) - ctx.template_pics
    return (not bad, "每张图都落在声明它的分块内（共 %d 张）" % n
            if not bad else "图文错配: %s" % "；".join(bad))


def _caption_pairs(ctx):
    """把「图 -> 图注文本框」配对：图注在图的**正下方**、水平居中对齐。"""
    pairs, orphan = [], []
    for px, py, pw, ph, _sh, _i in ctx.pics:
        if pw > 7.9:                                   # 模板整页背景图
            continue
        cands = [(x, y, w, h, t, i) for x, y, w, h, t, i in ctx.texts
                 if py + ph - 0.02 <= y <= py + ph + C.CAP_LINEGAP + C.CAP_H + 0.06
                 and abs((x + w / 2) - (px + pw / 2)) < 0.08
                 and t.strip()]
        if not cands:
            orphan.append("(%.2f,%.2f)" % (px, py))
            continue
        pairs.append(((px, py, pw, ph), cands[0]))
    return pairs, orphan


def _d2(ctx):
    pairs, orphan = _caption_pairs(ctx)
    n_fig = sum(len(f.paths) for f in ctx.figs)
    n_cap = len([1 for x, y, w, h, t, i in ctx.texts if _CAP_RE.match(t.strip())])
    ok = not orphan and len(pairs) == n_fig and n_cap == n_fig
    return ok, "%d 张图各配 1 条图注（共 %d 条）" % (n_fig, n_cap) if ok \
        else "图注配对异常：图 %d 张 / 配对 %d / 图注 %d，缺注的图 %s" % (
            n_fig, len(pairs), n_cap, orphan)


def _d3(ctx):
    pairs, _orphan = _caption_pairs(ctx)
    bad = []
    for (px, py, pw, ph), (x, y, w, h, t, _i) in pairs:
        if not (y >= py + ph - 0.03):
            bad.append("「%s」跑到图上方" % t.strip()[:12])
        if abs((x + w / 2) - (px + pw / 2)) > 0.05:
            bad.append("「%s」未居中" % t.strip()[:12])
    return (not bad, "%d 条图注均在该图正下方且居中" % len(pairs) if not bad
            else "图注位置: %s" % "；".join(bad))


def _d4(ctx):
    pairs, _orphan = _caption_pairs(ctx)
    bad = ["「%s」%.2f>%.2f" % (t.strip()[:10], w, pw)
           for (px, py, pw, ph), (x, y, w, h, t, _i) in pairs if w > pw + 0.01]
    return (not bad, "图注宽度都不超过所注图宽" if not bad
            else "图注压到邻图: %s" % "、".join(bad))


def _d5(ctx):
    """图号要么不用，要么从 图1 起连续不重复（不允许 图1/图3/图4 这种跳号）。"""
    nums = []
    for sec in ctx.task.sections:
        for cap in sec.captions:
            m = _CAP_RE.match(cap.strip())
            if m:
                nums.append(int(m.group(1)))
    n_cap = sum(len(s.captions) for s in ctx.task.sections)
    if not nums:
        return True, "本任务图注未使用图号"
    ok = len(nums) == n_cap and sorted(nums) == list(range(1, len(nums) + 1))
    label = "、".join("图%d" % i for i in sorted(nums))
    return ok, "图号 %s 连续不重复" % label if ok \
        else "图号异常：%s（图注共 %d 条，应从 图1 连续编到 图%d）" % (nums, n_cap, n_cap)


# ---- 文字 --------------------------------------------------------
def _content_text(ctx):
    t = "".join(s.text() for s in ctx.task.sections)
    t += "".join(ctx.task.titles)
    t += "".join(c for s in ctx.task.sections for c in s.captions)
    return t


def _e1(ctx):
    hit = sorted(set(_content_text(ctx)) & BAD_GLYPHS)
    return (not hit, "无字体缺失字形（℃ ₂ ⁺ 等已写成 ASCII 等价写法）" if not hit
            else "含字体缺失字形: %s" % " ".join(hit))


def _e2(ctx):
    bad = []
    for s in ctx.task.sections:
        if not s.text().strip():
            bad.append("「%s」无正文" % s.title)
        for w in STUB_WORDS:
            if w in s.text():
                bad.append("「%s」含未完成痕迹 %r" % (s.title, w))
    return (not bad, "%d 个分块都有正文，无占位/未完成痕迹" % len(ctx.task.sections) if not bad
            else "；".join(bad))


def _e3(ctx):
    ts = ctx.task.titles
    dup = sorted({t for t in ts if ts.count(t) > 1})
    return (not dup, "%d 个小标题互不重复" % len(ts) if not dup else "小标题重复: %s" % dup)


def _e4(ctx):
    red = black = 0
    for s in ctx.task.sections:
        for para in s.paras:
            for t, tag, _ in para:
                if tag == "red":
                    red += len(t)
                else:
                    black += len(t)
    ratio = red / max(1, red + black)
    return ratio <= 0.40, "标红 %d 字，占正文 %.1f%%（≤40%%）" % (red, ratio * 100)


def _e5(ctx):
    empty = [s.title for s in ctx.task.sections if s.figures and not all(c.strip() for c in s.captions)]
    return (not empty, "每条图注都非空" if not empty else "空图注: %s" % empty)


def _e6(ctx):
    """正文容量：字越多 → 求解器把字号压得越小 → 上了墙谁也看不清。

    上限不是拍的，是从示例（``assets/ref`` 的样张）量出的图文比来的：示例全篇
    约 250~300 汉字、最重的文本块约 130 字，其余板块只有图 + 短标签。见
    ``config.TEXT_CHAR_BLOCK_MAX / TEXT_CHAR_TOTAL_MAX``（可用任务文档 ``style:`` 覆盖）。
    """
    per = [(s.title, len(_HAN_RE.findall(s.text()))) for s in ctx.task.sections]
    tot = sum(n for _, n in per)
    worst = max(per, key=lambda kv: kv[1])
    over = ["「%s」%d 字" % (t, n) for t, n in per if n > C.TEXT_CHAR_BLOCK_MAX]
    ok = not over and tot <= C.TEXT_CHAR_TOTAL_MAX
    msg = "全篇 %d ≤ %d 字；最重块「%s」%d ≤ %d 字" % (
        tot, C.TEXT_CHAR_TOTAL_MAX, worst[0], worst[1], C.TEXT_CHAR_BLOCK_MAX)
    if over:
        msg += "；单块超限: " + "、".join(over)
    if tot > C.TEXT_CHAR_TOTAL_MAX:
        msg += "；全篇超限 %d 字" % (tot - C.TEXT_CHAR_TOTAL_MAX)
    return ok, msg


# 化学式 / 单位的 ASCII 直写形态：必须以角标（``^{}`` / ``_{}``）书写。
# 只列**确定是排版错误**的形态，避免误伤普通文本。
_ASCII_CHEM_RE = re.compile(
    r"(?:Fe|Cr|Co|Cu|Ni|Mn|Zn|Ti|V|Al)\d+[+-]"        # Fe3+ / Cr3+ / Co2+ …
    r"|(?:Fe|Cr|Co|Cu|Zn|Ti)\d+[A-Z]"                  # Fe3O4 / Cr2O3 …
    r"|\b(?:CO2|H2O|H2S|NO2|SO2|O2|N2)\b"              # CO2 / H2O …
    r"|\b(?:c?m|mm|km|dm)2\b"                          # m2 / cm2 …
    r"|I[_\s]?D\s*/\s*I[_\s]?G"                        # ID/IG（应为 I_D/I_G）
)


def _e7(ctx):
    """化学式与单位必须用真角标：``Fe^{3+}``、``元/m^{2}``、``I_{D}/I_{G}``。

    直接写 ``Fe3+`` 会被按普通字符排出来（3 和 + 与字母同高），上墙即判排版错误。
    角标只做位移、不改字号，所以不影响求解器的断行估算。
    """
    bad = []
    for s in ctx.task.sections:
        # 只看**非角标**的运行段：写成角标的电荷数/下标本来就不该被拦
        plain = "".join(t for para in s.paras for t, _, sc in para if not sc)
        hits = sorted({m.group(0) for m in _ASCII_CHEM_RE.finditer(plain)})
        if hits:
            bad.append("「%s」%s" % (s.title, "、".join(hits)))
    n_sup = sum(1 for s in ctx.task.sections for para in s.paras
                for _, _, sc in para if sc in ("sup", "sub"))
    return (not bad,
            "角标运行段 %d 个，无 ASCII 直写的化学式/单位" % n_sup if not bad
            else "需改成角标: " + "；".join(bad))


# ---- 交付 --------------------------------------------------------
def _f1(ctx):
    ok = ctx.out.is_file() and ctx.out.name == ctx.task.output
    return ok, "产物 %s（%.1f MB）" % (ctx.out.name, ctx.out.stat().st_size / 1e6) if ok \
        else "产物名与任务文档不一致或缺失: %s" % ctx.out


def _f2(ctx):
    try:
        prs = Presentation(str(ctx.out))
        n = len(prs.slides[0].shapes)
        return True, "可正常重新打开，%d 个形状 / %d 张幻灯片" % (n, len(prs.slides))
    except Exception as e:                                     # noqa: BLE001
        return False, "产物打不开: %s" % e


def _f3(ctx):
    from .cli import build

    first = _sha(ctx.out)
    build(task_ref=ctx.task.source, with_preview=False, verbose=False)
    second = _sha(ctx.out)
    return first == second, "两次构建字节一致（sha256 %s…）" % first[:12] if first == second \
        else "构建不可复现：%s… -> %s…" % (first[:12], second[:12])


def _f4(ctx):
    a = ctx.task.dist / C.PREVIEW_INTRO
    b = ctx.task.dist / C.PREVIEW_FULL
    ok = a.is_file() and b.is_file()
    return ok, "预览图已生成（%s / %s）" % (a.name, b.name) if ok else "缺预览图"


def _f5(ctx):
    """让真渲染器（PowerPoint）把产物打开、把文字墨迹框报回来对账。

    前面 C7~C10 都是"按模型自洽"的检查；只有这一条能发现
    "模型以为对、渲染器画出来不对"的问题。本机没有 PowerPoint 时记为跳过。

    "溢出声明框"只判**引擎排进去**的文字（前 ``n_tpl`` 个形状来自模板，
    它们的框高是模板作者定的，不归我们管）；"文字压到图"对全部文字生效。
    """
    from . import renderer_check as RC

    if not RC.available():
        return None, "本机没有 PowerPoint，跳过（装了就会自动复核）"
    try:
        n_tpl = RC.template_shape_count(ctx.task.template)
        problems = RC.verify(ctx.out, n_tpl=n_tpl)
    except Exception as e:                                     # noqa: BLE001
        return False, "PowerPoint 复核失败（文件可能打不开）: %s" % e
    return (not problems, "PowerPoint 实排复核通过：无文字压图、无文字溢出声明框"
            if not problems else "渲染器视角有问题: %s" % "；".join(problems[:3]))


# ---------------------------------------------------------------- 清单
REQUIREMENTS: list[Item] = [
    # 范围
    Item("A1", "范围", "只动「作品简介」区，模板既有字段一字未改", "模板/评审", "hard", _a1),
    Item("A2", "范围", "模板背景图保留，配图数与文档声明一致", "模板", "hard", _a2),
    # 尺寸
    Item("B1", "尺寸", "画布不被改动（8x20cm 缩放稿 = 实印 80x200cm）", "模板", "hard", _b1),
    Item("B2", "尺寸", "正文字号不低于下限（实印后仍可读）", "评审", "hard", _b2),
    Item("B3", "尺寸", "内容总高不溢出内容区", "模板", "hard", _b3),
    Item("B4", "尺寸", "每个分块都落在「作品简介」内容区内", "模板", "hard", _b4),
    Item("B5", "尺寸", "正文列宽有下限，不出现细长条", "样张", "hard", _b5),
    Item("B6", "尺寸", "配图高度有下限，不出现细缝", "样张", "hard", _b6),
    # 版式
    Item("C1", "版式", "每个分块是虚线圆角框", "样张", "hard", _c1),
    Item("C2", "版式", "块底为白 → 浅蓝渐变", "样张", "hard", _c2),
    Item("C3", "版式", "小标题是蓝色胶囊（全圆角 + 渐变 + 白字粗体）", "样张", "hard", _c3),
    Item("C4", "版式", "胶囊骑在一条整宽虚线上", "样张", "hard", _c4),
    Item("C5", "版式", "版面有节奏：混排网格，或各行配图数不同（不是千行一面）", "样张", "hard", _c5),
    Item("C6", "版式", "图文同框：有图的分块内同时有正文", "样张", "hard", _c6),
    Item("C7", "版式", "正文与配图不相互重叠", "工程", "hard", _c7),
    Item("C8", "版式", "文字实际需要的高度不超过自己声明的框高", "工程", "hard", _c8),
    Item("C9", "版式", "行距写 spcPct（spcPts 会被渲染器取整到整点）", "工程", "hard", _c9),
    Item("C10", "版式", "bodyPr 的自动调整子元素至多一个", "工程", "hard", _c10),
    Item("C11", "版式", "悬挂标点开关与断行策略一致", "工程", "hard", _c11),
    # 图文
    Item("D1", "图文", "每张图都落在声明它的那个分块内（防止错配）", "评审", "hard", _d1),
    Item("D2", "图文", "一张图一条图注，不得几图共用一注", "评审", "hard", _d2),
    Item("D3", "图文", "图注在该图正下方且水平居中于该图", "样张", "hard", _d3),
    Item("D4", "图文", "图注宽度不超过所注图宽", "工程", "hard", _d4),
    Item("D5", "图文", "图号从 图1 起连续、不重不漏", "样张", "hard", _d5),
    Item("D6", "图文", "素材归属按坐标轴判读并人工复核", "评审", "manual", None),
    # 文字
    Item("E1", "文字", "无字体缺失字形（℃ ₂ ⁺ 等改 ASCII 等价写法）", "工程", "hard", _e1),
    Item("E2", "文字", "每个分块都有正文，且无占位/未完成痕迹", "评审", "hard", _e2),
    Item("E3", "文字", "小标题互不重复", "工程", "hard", _e3),
    Item("E4", "文字", "标红只用于关键数据（占比 ≤ 40%）", "样张", "soft", _e4),
    Item("E5", "文字", "每条图注都非空", "工程", "hard", _e5),
    Item("E6", "文字", "正文字数有上限（单块 / 全篇），别写成小字报", "样张", "hard", _e6),
    Item("E7", "文字", "化学式 / 单位用真角标（`Fe^{3+}`、`I_{D}/I_{G}`、`m^{2}`）",
         "工程", "hard", _e7),
    # 交付
    Item("F1", "交付", "产物名与任务文档一致且存在", "工程", "hard", _f1),
    Item("F2", "交付", "产物可被重新打开（zip/xml 合法）", "工程", "hard", _f2),
    Item("F3", "交付", "构建可复现（同输入同字节）", "工程", "hard", _f3),
    Item("F4", "交付", "预览图已生成", "工程", "soft", _f4),
    Item("F5", "交付", "PowerPoint 实排复核：无文字压图 / 无溢出", "工程", "hard", _f5),
    # 流程
    Item("G1", "流程", "每次改动后跑 python build.py --check 且硬性项全绿", "工程", "manual", None),
    Item("G2", "流程", "版式变更记入 CHANGELOG.md（含动因与教训）", "工程", "manual", None),
]

BY_ID = {it.id: it for it in REQUIREMENTS}
GROUPS = ("范围", "尺寸", "版式", "图文", "文字", "交付", "流程")
MARK = {True: "PASS", False: "FAIL", None: "跳过"}


def run(ctx: Ctx, skip: tuple[str, ...] = ()) -> list[Result]:
    """跑一遍清单。``skip`` 里的编号跳过（例如不想花时间重跑 F3）。

    ``check`` 返回 ``ok=None`` 表示**本机跑不了**（例如没装 PowerPoint），
    记成「跳过」而不是失败 —— 别让环境缺件伪装成不合规。
    """
    out = []
    for it in REQUIREMENTS:
        if it.id in skip:
            continue
        if it.func is None:
            out.append(Result(it, None, "需人工确认（依据见 docs/compliance.md）"))
            continue
        try:
            ok, detail = it.func(ctx)
        except Exception as e:                                 # noqa: BLE001
            ok, detail = False, "校验自身出错: %s: %s" % (type(e).__name__, e)
        out.append(Result(it, None if ok is None else bool(ok), detail))
    return out


def failures(results: list[Result], only_hard: bool = True) -> list[Result]:
    return [r for r in results
            if r.ok is False and (not only_hard or r.item.level == "hard")]


def report(results: list[Result], task, ctx: Ctx | None = None) -> str:
    """把自检结果排成一份 Markdown 报告。"""
    hard = [r for r in results if r.item.level == "hard"]
    soft = [r for r in results if r.item.level == "soft"]
    man = [r for r in results if r.item.level == "manual"]
    n_fail = len([r for r in hard if r.ok is False])

    L = ["# 合规自检报告 · %s" % task.id, "",
         "任务：%s  " % task.title,
         "产物：`%s`" % (ctx.out.name if ctx else task.output),
         "清单：%d 项（硬性 %d / 提示 %d / 人工 %d）" % (
             len(results), len(hard), len(soft), len(man)), ""]
    if ctx is not None:
        n_blocks = len(ctx.lay.cells)
        n_fig = sum(len(f.paths) for f in ctx.figs)
        n_fig_blocks = sum(1 for f in ctx.figs if not f.empty)
        L += ["关键数据：正文 %.2fpt（实印 %.1fpt）｜图注 %.2fpt｜总高 %.2f / %.2fcm｜"
              "%d 个分块（%d 块有图 / 共 %d 张配图）"
              % (ctx.lay.body_pt, ctx.lay.body_pt * 10, ctx.cap_sz, ctx.lay.total, C.AVAIL_H,
                 n_blocks, n_fig_blocks, n_fig), ""]

    L += ["## 结论", "",
          "- 硬性项：**%d / %d 通过**%s" % (len(hard) - n_fail, len(hard),
                                       "（全部通过 ✅）" if not n_fail else "  ← 有未通过项 ❌"),
          "- 提示项：%d / %d 通过" % (len([r for r in soft if r.ok]), len(soft)),
          "- 人工项：%d 项待确认（见文末）" % len(man), ""]
    n_skip = len([r for r in results if r.ok is None and r.item.func is not None])
    if n_skip:
        L += ["- 跳过项：%d 项本机跑不了（装了 PowerPoint 就会自动复核）" % n_skip, ""]

    for g in GROUPS:
        rs = [r for r in results if r.item.group == g]
        if not rs:
            continue
        L += ["## %s" % g, "", "| 编号 | 要求 | 结果 | 说明 |", "|---|---|---|---|"]
        for r in rs:
            icon = {True: "✅", False: "❌", None: "⚠️"}[r.ok]
            L.append("| %s | %s | %s %s | %s |" % (
                r.item.id, r.item.title, icon, MARK[r.ok], r.detail))
        L.append("")

    if man:
        L += ["## 待人工确认（判读依据）", ""]
        for r in man:
            src = {
                "D6": "看坐标轴、别看图名：轴标签说明这张图是什么表征或数据，据此确认它是否"
                      "落在讨论它的那个分块里。认反过一次，务必看图不看名。",
                "G1": "改动文案 / 图 / 版式后，重跑 `python build.py --check`，硬性项须全绿再提交。",
                "G2": "版式 / 引擎变更在 `CHANGELOG.md` 对应版本条目下写清**动因与教训**，不只记结果。",
            }.get(r.item.id, "")
            L.append("- **%s** %s —— %s" % (r.item.id, r.item.title, src))
        L.append("")
    return "\n".join(L)
