# -*- coding: utf-8 -*-
"""全局配置：路径解析 + 版面几何 + 配色 + 按任务覆盖。

设计原则
--------
1. **不硬编码绝对路径**。模板、素材、产物一律从仓库根目录相对定位，
   并允许环境变量覆盖（``POSTER_TEMPLATE`` / ``POSTER_DIST``），便于换机、换目录。
2. **版面常量集中**。所有 cm 尺寸、字号、颜色只在本文件出现一次，
   调版式时不需要翻求解器和渲染器。
3. **这里的值是"默认值"，不是"唯一值"**。每个墙报任务的差异都写在它自己的
   任务文档（``tasks/<id>/task.md``）里，由 :func:`configure` 应用到本模块。
   引擎里的模块一律直接读 ``config.X``，因此任务覆盖必须在开工前完成。

⚠️ 由于覆盖是写给模块级变量的，**一个进程只应构建一个任务**。
   CLI 正是这么用的；测试里跨任务请调用 :func:`reset`。
"""
from __future__ import annotations

import os
from pathlib import Path

# ============================================================ 路径
REPO = Path(__file__).resolve().parents[2]        # src/poster/config.py -> 仓库根
TASKS = REPO / "tasks"                            # 每个墙报任务一个目录
ASSETS = REPO / "assets"                          # 跨任务共享素材
TEMPLATE_DIR = ASSETS / "template"                # 默认模板目录
DIST = Path(os.environ.get("POSTER_DIST", REPO / "dist"))
CACHE = REPO / ".cache"                           # 可再生缓存（模板背景 / 配图去底）

# 交付物 / 中间产物文件名（产物名默认取自任务文档，这里是兜底）
OUT_NAME = "poster.pptx"
PREVIEW_INTRO = "preview_intro.png"
PREVIEW_FULL = "preview_full.png"

# 系统字体（Windows）。缺失时 Pillow 会退化为内置位图字体，仅影响预览。
FONT = Path(r"C:/Windows/Fonts/msyh.ttc")
FONT_BOLD = Path(r"C:/Windows/Fonts/msyhbd.ttc")
FONT_FAMILY = "微软雅黑"

# 化学式的上/下标：OOXML 的 ``a:rPr/@baseline``（千分之一百分比）。
# 正值上标、负值下标；字号不变（PowerPoint 的角标只做位移、不缩放），
# 所以量宽度时无需区分角标 —— 求解器的断行结果与渲染器一致。
# 用法：正文写 ``Fe^{3+}``（上标）、``I_{D}/I_{G}``（下标）。
SCRIPT_BASELINE = {"sup": 30000, "sub": -25000}

# ============================================================ 模板探测
# 模板自带的「作品简介」占位文本框判定规则（top≈8.52cm、宽>5.5cm）。
# 换成别的赛事模板时，若探测不准，请改用任务文档里的显式 ``area:``。
PH_TOP, PH_TOP_TOL, PH_MIN_W = 8.52, 0.6, 5.5

# 模板里「作品编号」框的文字前缀。任务文档写了 ``number:`` 才按
# ``<前缀>：<编号>`` 覆写这个框；没写就一字不动（与作品名称 / 项目成员 /
# 本科生组三个模板字段同等待遇）。
NUMBER_PREFIX = "作品编号"

# 占位框 -> 实际可排版内区的内缩量（left, top, right, bottom，实测值）。
# 占位框比真正的白底内区略大：左 0.05、上 0.22、下 0.42cm，右边恰好齐平。
AREA_INSET = (0.05, 0.22, 0.00, 0.42)
FALLBACK_AREA = (0.74, 8.74, 7.30, 18.30)        # 探测失败时的兜底内容区
FALLBACK_CANVAS = (8.0, 20.0)                    # 探测失败时的兜底画布（cm）

# ============================================================ 版面（cm）
# 「作品简介」区在模板中的可用矩形；由 :func:`set_area` 设定
X0, Y0, X1, Y1 = FALLBACK_AREA
W, H = X1 - X0, Y1 - Y0

PAD_X, PAD_T, PAD_B = 0.15, 0.14, 0.14
INNER_W = W - 2 * PAD_X                 # 通栏块宽
AVAIL_H = H - PAD_T - PAD_B             # 纵向可用高

BP_X, BP_Y = 0.10, 0.07                 # 块内左右 / 上下内边距
BGAP = 0.10                             # 块间距
COL_GAP = 0.12                          # 并排块之间的列距
COLGAP_IN = 0.13                        # 块内「文 | 图」之间的间距
FIG_GAP_IN = 0.10                       # 同一图条内多张配图并排时的间距
FIG_ROW_GAP = 0.10                      # 同一块内上下两「图条行」之间的间距（图网格的行距）
TEXT_IMG_GAP = 0.08                     # 「文上图下」时文与图之间距
CAP_LINEGAP, CAP_H = 0.03, 0.13         # 图注与图之间距 / 图注行高的下限（cm）
CAP_LINE_FACTOR = 1.12                  # 图注行距倍数（真实行高 = 图注字号 x 本值）

RADIUS = 0.14                           # 块圆角半径
BORDER_W = 0.8                          # 块边框线宽（pt）
DASH_LEN, DASH_GAP = 7, 5               # 预览用虚线段 / 间隙（px）

TITLE_SZ = 5.0                          # 胶囊标题字号（pt）
TITLE_TOP = 0.07                        # 胶囊距块顶
TITLE_GAP = 0.07                        # 胶囊与正文之间距
CHIP_PAD_X = 0.11                       # 胶囊内左右留白

LINE_FACTOR = 1.60                      # 行距倍数（行高 = 字号 x 本值）
LINE_REF_FACTOR = 1.20                  # PowerPoint「单倍行距」= 字号 x 本值，换算 spcPct 用
MEASURE_SAFETY = 0.97                   # 量算宽度打折，抵掉渲染器的断行差异
#
# 取值依据（2026-10-16 实测）：Pillow 用同一份字体文件测出的行宽，比 PowerPoint
# 实排**窄约 1.5%** —— 于是偶发"估算 10 行、实排 11 行"，末行压到下面的图上
# （传感块在 2.87cm 分栏宽度下复现）。0.985 兜不住这一行，0.97 才够。
# 注意这是**全局**系数：一旦真出现估算/实排不一致，先怀疑这里，
# 但**不要把 F5 实排复核当成可选项**——它是这类差异的最后一道闸。

# 正文**字符间距**（pt/字），写进 pptx 的 run 属性 ``a:rPr/@spc``（单位 1/100 pt）。
# 0.1pt 让密排的汉字之间透一口气；量算宽度时**必须**把它算进去（见 typography._spc_px），
# 否则求解器会以为每行比实际窄，行数估少、末行压到图上。
CHAR_SPC_PT = 0.1

# 量宽时把字距**放大**这么多倍再算（只影响求解，不影响写进 pptx 的值）。
# 为什么：字距是**绝对量、不随字号缩放**，而 Pillow 量出的字宽本就比 PowerPoint
# 实排窄一点；两条误差叠加后，1x 计字距仍会偶尔少算一行（2026-10-16 实测，
# 窄栏块实排 8 行、模型 7 行，末行压到图上）。取 2x 作悲观估计。
CHAR_SPC_MEASURE_SCALE = 2.0

# 行尾标点要不要「悬挂」在行外（中西文混排里叫悬挂标点）
#   True  -> pptx 里写 bodyPr/@hangingPunct="1"，渲染器让行尾标点溢出边界，
#            于是行首不会出现「，。、」；本引擎按 hang 策略画，与渲染器一致。
#   False -> 用渲染器的默认行为（把标点推到下一行），行首可能出现标点。
# 无论取哪个值，**求解尺寸一律用三种策略里最悲观的那个**（见 lines_count_worst），
# 所以换成别的渲染器也不会溢出。
HANG_PUNCT = True
DRAW_POLICY = "hang" if HANG_PUNCT else "push"     # 由 _recompute() 保持同步

# ---- 求解器搜索区间 ----
BODY_MAX, BODY_MIN = 4.20, 2.60         # 正文字号搜索范围（pt）
IH_MAX, IH_MIN = 2.20, 0.90             # 配图高搜索范围（cm）
IH_FLOOR = 1.00                         # 可接受的最小图高
IH_MIN_READABLE = 0.70                  # 合规下限：图高再低就成「细缝」（实印 <7cm 读不清坐标轴）
SIDE_MIN = 1.20                         # 左文右图块的图高下限
TW_MIN = 1.30                           # 左文右图块的正文列宽下限
BODY_TARGET = 2.75                      # 字号目标线：≥ 此值才肯为「图更大」让路
STEPS = 0.05                            # 搜索步长

# ---- 正文容量 ----
# 海报是**站着远看、看图为主**的媒介：字越多，求解器就得把字号压得越小，
# 最后谁也看不清（"字写满了"和"没人读"往往是同一张海报）。
# 上限来自示例（assets/ref 里的样张）量出的图文比：示例全篇约 250~300 汉字，
# 最重的文本块约 130 字，其余板块基本只有图 + 短标签。
#
# 2026-10-16 按一份需要把机制讲清楚的**详版文案**放宽（原先 120 / 550 卡死了它）：
# 该版单块最多 201 汉字、全篇 752 汉字，正文 3.30pt（实印 33pt，仍远高于
# B2 的 2.60pt 下限）—— 这种密度是"讲清楚"，不是"小字报"，用示例的图文比
# 一刀切反而把该说的内容砍掉了。**放宽的正当性只来自字号仍够大，不来自字写多了。**
# 计数口径 = **汉字**（不含数字 / 英文 / 标点）。
TEXT_CHAR_BLOCK_MAX = 220               # 单个分块的正文汉字数上限
TEXT_CHAR_TOTAL_MAX = 800               # 全篇分块正文的汉字数上限

# ============================================================ 配色（RGB）
BLACK = (0x1A, 0x1A, 0x1A)
RED = (0xC0, 0x00, 0x00)                # 关键数据标红（文档里写 **文字**）
GRAY = (0x45, 0x45, 0x45)
PLATE = (0xFF, 0xFF, 0xFF)              # 底板
BLOCK_C1, BLOCK_C2 = (0xFF, 0xFF, 0xFF), (0xE8, 0xF4, 0xFD)   # 块底渐变
BLOCK_LINE = (0x7F, 0xB5, 0xD9)         # 块虚线框
IMG_LINE = (0xBF, 0xD8, 0xEA)           # 配图细边
CHIP_C1, CHIP_C2 = (0x1F, 0x7F, 0xC0), (0x5A, 0xB4, 0xE6)     # 胶囊渐变
CHIP_LINK = (0x9F, 0xC8, 0xE4)          # 胶囊后的虚线

# ============================================================ 预览
PX_PER_CM = 200.0                       # 预览渲染分辨率（px/cm）
PT2CM = 2.54 / 72.0                     # pt -> cm
CANVAS_CM = FALLBACK_CANVAS             # 预览画布（cm），由模板 slide size 设定
PREVIEW_HEAD_PAD = 0.84                 # 局部预览上扩（把「作品简介：」标签也框进来）
PREVIEW_FOOT_PAD = 0.50                 # 局部预览下扩

CHIP_H = round(TITLE_SZ * PT2CM + 0.10, 4)      # 胶囊高度
BLOCK_OH = TITLE_TOP + CHIP_H + TITLE_GAP + BP_Y  # 块顶部固定开销

# ============================================================ 派生量重算


def _recompute() -> None:
    """重算所有派生常量。改动任何几何常量后必须调用。"""
    g = globals()
    g["W"] = g["X1"] - g["X0"]
    g["H"] = g["Y1"] - g["Y0"]
    g["INNER_W"] = g["W"] - 2 * g["PAD_X"]
    g["AVAIL_H"] = g["H"] - g["PAD_T"] - g["PAD_B"]
    g["CHIP_H"] = round(g["TITLE_SZ"] * g["PT2CM"] + 0.10, 4)
    g["BLOCK_OH"] = g["TITLE_TOP"] + g["CHIP_H"] + g["TITLE_GAP"] + g["BP_Y"]
    # 画出来用哪个断行策略：开了悬挂标点就跟渲染器说好一起悬挂
    g["DRAW_POLICY"] = "hang" if g["HANG_PUNCT"] else "push"


def set_area(x0: float, y0: float, x1: float, y1: float) -> None:
    """设定「作品简介」可排版内容区（cm，左上原点）。"""
    if not (x1 > x0 and y1 > y0):
        raise ValueError("内容区矩形非法: (%s, %s) -> (%s, %s)" % (x0, y0, x1, y1))
    globals().update({"X0": x0, "Y0": y0, "X1": x1, "Y1": y1})
    _recompute()


# 任务文档 ``style:`` 段允许覆盖的键。未列入的一律拒绝，防止拼错键名后静默失效。
OVERRIDABLE = {
    # 几何
    "PAD_X", "PAD_T", "PAD_B", "BP_X", "BP_Y", "BGAP", "COL_GAP", "COLGAP_IN",
    "FIG_GAP_IN", "TEXT_IMG_GAP", "CAP_LINEGAP", "CAP_H", "CAP_LINE_FACTOR",
    "RADIUS", "BORDER_W",
    "TITLE_SZ", "TITLE_TOP", "TITLE_GAP", "CHIP_PAD_X", "LINE_FACTOR",
    "LINE_REF_FACTOR", "MEASURE_SAFETY", "CHAR_SPC_PT", "CHAR_SPC_MEASURE_SCALE", "HANG_PUNCT",
    # 求解器
    "BODY_MAX", "BODY_MIN", "IH_MAX", "IH_MIN", "IH_FLOOR", "IH_MIN_READABLE",
    "SIDE_MIN", "TW_MIN", "BODY_TARGET", "STEPS",
    # 正文容量
    "TEXT_CHAR_BLOCK_MAX", "TEXT_CHAR_TOTAL_MAX",
    # 配色
    "BLACK", "RED", "GRAY", "PLATE", "BLOCK_C1", "BLOCK_C2", "BLOCK_LINE",
    "IMG_LINE", "CHIP_C1", "CHIP_C2", "CHIP_LINK",
    # 字体
    "FONT", "FONT_BOLD", "FONT_FAMILY",
    # 预览
    "PX_PER_CM", "DASH_LEN", "DASH_GAP", "PREVIEW_HEAD_PAD", "PREVIEW_FOOT_PAD",
}
_COLOR_KEYS = {"BLACK", "RED", "GRAY", "PLATE", "BLOCK_C1", "BLOCK_C2",
               "BLOCK_LINE", "IMG_LINE", "CHIP_C1", "CHIP_C2", "CHIP_LINK"}
_BOOL_KEYS = {"HANG_PUNCT"}


def _coerce(key: str, val):
    """把 YAML 读进来的值收敛成引擎里该有的类型。"""
    if key in _COLOR_KEYS:
        v = tuple(int(x) for x in val)
        if len(v) != 3 or not all(0 <= x <= 255 for x in v):
            raise ValueError("颜色 %s 必须是 3 个 0-255 整数，得到 %r" % (key, val))
        return v
    if key in ("FONT", "FONT_BOLD"):
        return Path(val)
    if key in _BOOL_KEYS:
        if not isinstance(val, bool):
            raise ValueError("配置 %s 只接受 true / false，得到 %r" % (key, val))
        return val
    if isinstance(val, bool):
        raise ValueError("配置 %s 不接受布尔值" % key)
    if isinstance(val, (int, float)):
        return float(val)
    return val


def apply_style(style: dict | None) -> None:
    """应用任务文档 ``style:`` 段的覆盖。键名必须在 :data:`OVERRIDABLE` 内。"""
    if not style:
        return
    unknown = sorted(set(style) - OVERRIDABLE)
    if unknown:
        raise KeyError("任务文档 style 段含未支持的键: %s（可用键见 config.OVERRIDABLE）"
                       % ", ".join(unknown))
    globals().update({k: _coerce(k, v) for k, v in style.items()})
    _recompute()


# 用于 :func:`reset` 的出厂快照（在模块末尾采集）
_DEFAULTS: dict = {}


def reset() -> None:
    """恢复到出厂默认（测试跨任务时用）。"""
    globals().update(_DEFAULTS)


def set_canvas(w: float, h: float) -> None:
    """设定预览画布尺寸（cm），取自模板 slide size。"""
    globals()["CANVAS_CM"] = (float(w), float(h))


def configure(task) -> None:
    """按 :class:`poster.spec.Task` 配置本进程：内容区 + 风格覆盖。"""
    reset()
    if task.area is not None:
        set_area(*task.area)
    apply_style(task.style)


def ensure_dirs() -> None:
    """确保输出与缓存目录存在。"""
    DIST.mkdir(parents=True, exist_ok=True)
    CACHE.mkdir(parents=True, exist_ok=True)


def task_cache(task_id: str) -> Path:
    """任务独占的缓存目录（配图去底 / 裁边的中间产物）。"""
    p = CACHE / task_id
    p.mkdir(parents=True, exist_ok=True)
    return p


def template_path() -> Path:
    """定位模板 pptx：环境变量 > assets/template 下唯一 pptx。

    注意：任务文档里的 ``template:`` 优先级更高，由 :mod:`poster.spec` 解析并
    传给引擎；本函数只是"没指定模板时"的兜底。
    """
    env = os.environ.get("POSTER_TEMPLATE")
    if env:
        p = Path(env)
        if p.is_file():
            return p
        raise FileNotFoundError(f"POSTER_TEMPLATE 指向的文件不存在: {p}")
    cands = sorted(TEMPLATE_DIR.glob("*.pptx"))
    if cands:
        return cands[0]
    raise FileNotFoundError(
        f"未找到模板：请把模板 pptx 放到 {TEMPLATE_DIR}，或设置 POSTER_TEMPLATE，"
        "或在任务文档里写 template:"
    )


_DEFAULTS.update({
    k: globals()[k] for k in (
        # 内容区与派生量
        "X0", "Y0", "X1", "Y1", "W", "H", "INNER_W", "AVAIL_H", "CHIP_H", "BLOCK_OH",
        # 版面
        "PAD_X", "PAD_T", "PAD_B", "BP_X", "BP_Y", "BGAP", "COL_GAP", "COLGAP_IN",
        "FIG_GAP_IN", "TEXT_IMG_GAP", "CAP_LINEGAP", "CAP_H", "CAP_LINE_FACTOR",
    "RADIUS", "BORDER_W",
        "TITLE_SZ", "TITLE_TOP", "TITLE_GAP", "CHIP_PAD_X", "LINE_FACTOR",
        "MEASURE_SAFETY", "PT2CM", "LINE_REF_FACTOR", "CHAR_SPC_PT",
        "CHAR_SPC_MEASURE_SCALE", "HANG_PUNCT", "DRAW_POLICY",
        # 求解器
        "BODY_MAX", "BODY_MIN", "IH_MAX", "IH_MIN", "IH_FLOOR", "IH_MIN_READABLE",
        "SIDE_MIN", "TW_MIN", "BODY_TARGET", "STEPS",
        # 配色
        "BLACK", "RED", "GRAY", "PLATE", "BLOCK_C1", "BLOCK_C2", "BLOCK_LINE",
        "IMG_LINE", "CHIP_C1", "CHIP_C2", "CHIP_LINK",
        # 字体
        "FONT", "FONT_BOLD", "FONT_FAMILY",
        # 预览
        "PX_PER_CM", "DASH_LEN", "DASH_GAP", "CANVAS_CM",
        "PREVIEW_HEAD_PAD", "PREVIEW_FOOT_PAD",
        # 模板探测
        "PH_TOP", "PH_TOP_TOL", "PH_MIN_W", "AREA_INSET",
        "FALLBACK_AREA", "FALLBACK_CANVAS",
    )
})
