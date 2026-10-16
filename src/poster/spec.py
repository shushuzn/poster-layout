# -*- coding: utf-8 -*-
"""任务文档解析：把 ``tasks/<id>/task.md`` 变成一份可执行的墙报任务。

一个墙报任务 = 一份 Markdown 文档。文档由两半组成：

* **YAML front-matter** —— 任务级设置：模板、产物名、内容区、网格、风格覆盖。
* **Markdown 正文** —— 每个 ``## 小标题`` 就是版面上的一个分块；
  ``<!-- figure: 文件 | 图注 ; 文件 | 图注 -->`` 给该块配图。
  多张用**分号**分隔、版面上等高并排成一「图条」，**每张图必须各写自己的图注**
  （“一张图一个图注”）；``**文字**`` 标红。

格式细节与可直接抄的范例见 ``docs/task-format.md`` 与 ``tasks/_template.md``。
"""
from __future__ import annotations

import re
import shutil
from dataclasses import dataclass, field
from pathlib import Path

import yaml

from . import config as C

# ---------------------------------------------------------------- 词法
_FRONT_RE = re.compile(r"\A\ufeff?---[ \t]*\r?\n(.*?)\r?\n---[ \t]*\r?\n?", re.S)
_H2_RE = re.compile(r"^##[ \t]+(.+?)[ \t]*$")
_DIRECTIVE_RE = re.compile(r"^[ \t]*<!--[ \t]*([A-Za-z_]+)[ \t]*:[ \t]*(.*?)[ \t]*-->[ \t]*$")
_COMMENT_RE = re.compile(r"<!--.*?-->", re.S)
_EMPH_RE = re.compile(r"\*\*(.+?)\*\*", re.S)
# 角标：^{...} 上标 / _{...} 下标（前面的反斜杠表示字面量）
_SCRIPT_RE = re.compile(r"(?<!\\)([\^_])\{([^{}]*)\}")

MODES = ("auto", "stack", "side", "text")
NO_FIGURE = {"", "-", "none", "null", "无"}
KNOWN_DIRECTIVES = ("figure", "caption", "mode")
MAX_FIGURES = 3                     # 单条图条里的配图张数上限（再多就挤成细缝了）
MAX_FIG_ROWS = 3                    # 单块图条行数上限（图网格最多三行）


def tag_colors() -> dict[str, tuple[int, int, int]]:
    """颜色标签 -> RGB。**必须在使用时调用**，因为任务覆盖是后生效的。"""
    return {"black": C.BLACK, "red": C.RED, "gray": C.GRAY}


# ---------------------------------------------------------------- 数据结构
@dataclass
class Section:
    """版面上的一个分块。"""

    title: str
    # 段落 -> 运行段 (文本, 颜色标签, 角标)；角标 ""/"sup"/"sub"（见 rich_text）
    paras: list[list[tuple[str, str, str]]] = field(default_factory=list)
    figures: list[str] = field(default_factory=list)      # 配图文件名（可 1~n 张，按行拍平）
    captions: list[str] = field(default_factory=list)     # 每张图各自的图注，与 figures 一一对应
    mode: str = "auto"

    # 图网格：fig_rows[k] = 第 k 条图条行的文件名。单行时 = [figures]。
    # 示例（样张）的并排块内部都是「两行图网格」，只有支持多行才能既分栏又放大图。
    fig_rows: list[list[str]] = field(default_factory=list)
    cap_rows: list[list[str]] = field(default_factory=list)

    # 解析后填入（绝对路径；figure_paths 是 fig_rows_paths 的拍平）
    figure_paths: list[Path] = field(default_factory=list)
    fig_rows_paths: list[list[Path]] = field(default_factory=list)

    @property
    def has_figures(self) -> bool:
        return bool(self.figure_paths)

    def rich(self) -> list[list[tuple[str, tuple[int, int, int], str]]]:
        """把颜色标签换成当前生效的 RGB，保留角标标记。"""
        pal = tag_colors()
        return [[(t, pal.get(c, pal["black"]), s) for t, c, s in para]
                for para in self.paras]

    def text(self) -> str:
        return "".join(t for para in self.paras for t, _, _ in para)


@dataclass
class Task:
    """一份任务文档的完整内容。"""

    id: str
    title: str
    source: Path                       # task.md 路径
    root: Path                         # 任务目录
    template: Path
    output: str                        # 产物文件名
    dist: Path
    figures_dir: Path
    sections: list[Section]
    rows: list[list[int]]
    row_fracs: list[list[float]]
    area: tuple[float, float, float, float] | None      # None = 自动探测
    style: dict = field(default_factory=dict)
    meta: dict = field(default_factory=dict)
    _paras_cache: list = field(default_factory=list, repr=False, compare=False)
    _paras_key: tuple = field(default=(), repr=False, compare=False)

    # ---------------------------------------------------------- 便捷视图
    @property
    def titles(self) -> list[str]:
        return [s.title for s in self.sections]

    @property
    def captions(self) -> list[list[str]]:
        """``captions[i][j]`` = 第 i 块第 j 张图的图注（一张图一条）。"""
        return [s.captions for s in self.sections]

    @property
    def figures(self) -> list[list[Path]]:
        """``figures[i]`` = 第 i 块的全部配图路径（无图则是空列表）。"""
        return [s.figure_paths for s in self.sections]

    @property
    def paras(self):
        """``paras[i]`` = 第 i 块的段落列表，段落内是 ``(文本, RGB)``。

        求解器会在搜索循环里反复取它，所以按当前配色缓存；
        任务若改了配色标签对应的 RGB，缓存会自动失效。
        """
        pal = tag_colors()
        key = (pal["black"], pal["red"], pal["gray"])
        if self._paras_key != key:
            self._paras_cache = [s.rich() for s in self.sections]
            self._paras_key = key
        return self._paras_cache

    def char_count(self) -> int:
        """正文总字数（用于构建日志）。"""
        return sum(len(s.text()) for s in self.sections)

    def out_path(self) -> Path:
        return self.dist / self.output

    # ---------------------------------------------------------- 展示
    def summary(self) -> str:
        lines = ["任务 %s（%s）" % (self.id, self.title),
                 "  文档 %s" % self.source,
                 "  模板 %s" % self.template,
                 "  产物 %s" % self.out_path(),
                 "  内容区 %s" % ("自动探测" if self.area is None else
                                  "%.2f,%.2f -> %.2f,%.2f" % self.area),
                 "  网格 %s" % self.rows]
        for i, s in enumerate(self.sections):
            if s.figures:
                fig = "；".join("%s[%s]" % (n, c) for n, c in zip(s.figures, s.captions))
            else:
                fig = "（无配图）"
            lines.append("  #%d %-5s %s" % (i, s.mode, s.title))
            lines.append("       %s" % fig)
        return "\n".join(lines)


# ---------------------------------------------------------------- 富文本
def rich_text(text: str, base: str = "black", accent: str = "red") \
        -> list[tuple[str, str, str]]:
    """把 ``**强调**``、``^{上标}``、``_{下标}`` 解析成 ``[(文本, 颜色标签, 角标)]``。

    角标取 ``""`` / ``"sup"`` / ``"sub"``。化学式与单位必须写成角标形式，
    例如 ``Fe^{3+}``、``元/m^{2}``、``I_{D}/I_{G}`` —— 直接写 ``Fe3+`` 属于
    排版错误（合规项 ``E7`` 会拦）。想输出字面量 ``^{``，写成 ``\\^{``。
    """
    out: list[tuple[str, str, str]] = []

    def push(text_: str, color: str, script: str = "") -> None:
        # 逐个片段再切一遍强调，保证 ``**`` 与角标可以嵌套
        pos = 0
        for m in _EMPH_RE.finditer(text_):
            if m.start() > pos:
                out.append((text_[pos:m.start()], color, script))
            out.append((m.group(1), accent, script))
            pos = m.end()
        if pos < len(text_):
            out.append((text_[pos:], color, script))

    pos = 0
    for m in _SCRIPT_RE.finditer(text):
        if m.start() > pos:
            push(text[pos:m.start()], base)
        body = m.group(2).replace("\\^{", "^{").replace("\\_{", "_{")
        push(body, base, "sup" if m.group(1) == "^" else "sub")
        pos = m.end()
    if pos < len(text):
        push(text[pos:], base)
    return [(t, c, s) for t, c, s in out if t]


# ---------------------------------------------------------------- 正文解析
def _split_front(text: str) -> tuple[dict, str]:
    m = _FRONT_RE.match(text)
    if not m:
        return {}, text
    meta = yaml.safe_load(m.group(1)) or {}
    if not isinstance(meta, dict):
        raise ValueError("任务文档 front-matter 必须是键值对")
    return meta, text[m.end():]


# 张与张之间用分号；图条行之间用双竖线；一条 ``figure`` 里就是一个块的图网格
_ITEM_SPLIT_RE = re.compile(r"[;；]")
_COMMA_RE = re.compile(r"[,、，]")
_ROW_SPLIT_RE = re.compile(r"\|\|")


def _split_figures(val: str, where: str, title: str) \
        -> tuple[list[list[str]], list[list[str]]]:
    """把 ``figure:`` 的值切成 **按行分组** 的 ``(文件名, 图注)``。

    语法::

        <!-- figure: 图a | 图1 甲 ; 图b | 图2 乙 -->            一条图条（两张并排）
        <!-- figure: 图a | 图1 甲 ; 图b | 图2 乙  ||  图c | 图3 丙 -->   两行图网格

    双竖线 ``||`` 分「图条行」：行内分号分张、竖线分文件名与图注。
    块内多行是**分栏版式下把图放大**的唯一手段（样张的并排块内部都是图网格），
    单块最多 MAX_FIG_ROWS 行、每行最多 MAX_FIGURES 张。
    """
    rows = [x.strip() for x in _ROW_SPLIT_RE.split(val) if x.strip()]
    if not rows:
        return [], []
    names_rows: list[list[str]] = []
    caps_rows: list[list[str]] = []
    for r in rows:
        n, c = _split_one_row(r, where, title)
        if n:
            names_rows.append(n)
            caps_rows.append(c)
    if sum(len(r) for r in names_rows) > MAX_FIGURES * MAX_FIG_ROWS:
        raise ValueError("%s: 小节「%s」配了 %d 张图，单块最多 %d 行 x %d 张"
                         % (where, title, sum(len(r) for r in names_rows),
                            MAX_FIG_ROWS, MAX_FIGURES))
    return names_rows, caps_rows


def _split_one_row(val: str, where: str, title: str) -> tuple[list[str], list[str]]:
    """切**一条图条行**里的若干个「文件 | 图注」项。

    一张图一条图注 —— 多张图**必须逐张写注**，不允许几张贴一条注
    （版面上它们是等高的独立图，各有各的图号与说明）。
    整条写成 ``none`` 表示该块不配图。
    """
    items = [x.strip() for x in _ITEM_SPLIT_RE.split(val) if x.strip()]
    if not items:
        return [], []
    if len(items) == 1 and items[0].partition("|")[0].strip().lower() in NO_FIGURE:
        return [], []

    names: list[str] = []
    caps: list[str] = []
    for it in items:
        name, _, cap = it.partition("|")
        name, cap = name.strip(), cap.strip()
        if not name:
            raise ValueError("%s: 小节「%s」的 figure 里 %r 缺少文件名" % (where, title, it))
        if name.lower() in NO_FIGURE:
            raise ValueError("%s: 小节「%s」的 figure 里混进了 %r；"
                             "不要配图请整行写 figure: none" % (where, title, name))
        if _COMMA_RE.search(name):
            raise ValueError(
                "%s: 小节「%s」的 figure 用逗号列了多个文件名（%r）。\n"
                "  多张配图要用分号分隔，且**每张图各写自己的图注**：\n"
                "      <!-- figure: 图a.png | 图1 流程示意 ; 图b.png | 图2 样品照片 -->"
                % (where, title, name))
        names.append(name)
        caps.append(cap)

    if len(names) > MAX_FIGURES:
        raise ValueError("%s: 小节「%s」配了 %d 张图，单块最多 %d 张"
                         % (where, title, len(names), MAX_FIGURES))
    return names, caps


def _parse_section(title: str, lines: list[str], where: str) -> Section:
    """解析单个小节的指令行与正文。"""
    sec = Section(title=title)
    rest: list[str] = []
    seen: set[str] = set()
    for ln in lines:
        m = _DIRECTIVE_RE.match(ln)
        if not m:
            rest.append(ln)
            continue
        key, val = m.group(1).lower(), m.group(2).strip()
        if key not in KNOWN_DIRECTIVES:
            raise ValueError("%s: 未知指令 %r（可用：%s）"
                             % (where, key, ", ".join(KNOWN_DIRECTIVES)))
        if key in seen:
            raise ValueError("%s: 小节「%s」的 %s 指令写了多次；"
                             "多张配图请写在同一条 figure 里，用分号分隔；"
                             "要排成图网格就用双竖线 || 分行"
                             % (where, title, key))
        seen.add(key)
        if key == "figure":
            sec.fig_rows, sec.cap_rows = _split_figures(val, where, title)
            sec.figures = [n for row in sec.fig_rows for n in row]
            sec.captions = [c for row in sec.cap_rows for c in row]
        elif key == "caption":
            # 单独指定图注：多张图就用分号写成多条，与图片一一对应
            sec.captions = [x.strip() for x in _ITEM_SPLIT_RE.split(val) if x.strip()]
            sec.cap_rows = [[c] for c in sec.captions]
        else:
            if val.lower() not in MODES:
                raise ValueError("%s: mode 只能是 %s，得到 %r"
                                 % (where, "/".join(MODES), val))
            sec.mode = val.lower()

    # 指令摘掉之后再清注释（注释可能跨多行，所以整段处理）
    body = _COMMENT_RE.sub("", "\n".join(rest))

    # 空行分段；段落内换行无意义（中文排版），直接拼接
    para: list[str] = []
    for ln in body.splitlines():
        if ln.strip():
            para.append(ln.strip())
        elif para:
            sec.paras.append(rich_text("".join(para)))
            para = []

    if para:
        sec.paras.append(rich_text("".join(para)))
    if not sec.paras:
        raise ValueError("%s: 小节「%s」没有正文" % (where, title))

    if not sec.figures and sec.mode not in ("auto", "text"):
        raise ValueError("%s: 小节「%s」没有配图，mode 不能是 %s" % (where, title, sec.mode))
    if not sec.figures:
        sec.mode = "text"
        sec.captions = []
    else:
        if len(sec.captions) != len(sec.figures):
            raise ValueError(
                "%s: 小节「%s」有 %d 张图，图注却有 %d 条 —— 必须一张图一条图注"
                % (where, title, len(sec.figures), len(sec.captions)))
        # 图注校验放在所有指令读完之后：既支持 figure 里逐张写注，
        # 也支持用 caption: 指令逐条补（两种写法都是"一张图一条"）。
        missing = [n for n, c in zip(sec.figures, sec.captions) if not c.strip()]
        if len(sec.figures) > 1 and missing:
            raise ValueError(
                "%s: 小节「%s」有 %d 张图，但 %s 没有图注。\n"
                "  一张图必须配一条自己的图注（不能几张贴一条）：\n"
                "      <!-- figure: 图a.png | 图1 流程示意 ; 图b.png | 图2 样品照片 -->"
                % (where, title, len(sec.figures), "、".join(missing)))
        # 单图没写图注时退回小节标题（多图在上一段已强制逐张写注）
        sec.captions = [c.strip() or sec.title for c in sec.captions]
    return sec


def parse_sections(body: str, where: str = "task.md") -> list[Section]:
    """把 Markdown 正文切成若干小节。"""
    sections: list[Section] = []
    title: str | None = None
    buf: list[str] = []
    for ln in body.splitlines():
        m = _H2_RE.match(ln)
        if m:
            if title is not None:
                sections.append(_parse_section(title, buf, where))
            title, buf = m.group(1).strip(), []
        elif title is not None:
            buf.append(ln)
    if title is not None:
        sections.append(_parse_section(title, buf, where))
    if not sections:
        raise ValueError("%s: 没找到任何小节（用 '## 小标题' 起一个小节）" % where)
    return sections


# ---------------------------------------------------------------- 网格
def auto_rows(n: int) -> list[list[int]]:
    """默认网格：首节通栏，其余两两并排（示例墙报的节奏）。"""
    rows = [[0]]
    i = 1
    while i < n:
        rows.append([i] + ([i + 1] if i + 1 < n else []))
        i += 2
    return rows


def _norm_rows(raw, n: int, where: str) -> list[list[int]]:
    if raw is None:
        return auto_rows(n)
    rows = [[int(x) for x in r] for r in raw]
    flat = [x for r in rows for x in r]
    if sorted(flat) != list(range(n)):
        raise ValueError("%s: grid 必须不重不漏地覆盖全部 %d 个小节，得到 %s"
                         % (where, n, flat))
    return rows


def _norm_fracs(raw, rows: list[list[int]], where: str) -> list[list[float]]:
    if raw is None:
        return [[round(1.0 / len(r), 6)] * len(r) for r in rows]
    fracs = [[float(x) for x in r] for r in raw]
    if len(fracs) != len(rows):
        raise ValueError("%s: row_fracs 有 %d 行，grid 有 %d 行"
                         % (where, len(fracs), len(rows)))
    for i, (f, r) in enumerate(zip(fracs, rows)):
        if len(f) != len(r):
            raise ValueError("%s: row_fracs 第 %d 行有 %d 个数，grid 该行有 %d 块"
                             % (where, i, len(f), len(r)))
        if abs(sum(f) - 1.0) > 1e-6:
            raise ValueError("%s: row_fracs 第 %d 行合计 %.4f，应为 1.0"
                             % (where, i, sum(f)))
    return fracs


# ---------------------------------------------------------------- 解析主入口
def _resolve(base: Path, val) -> Path:
    p = Path(str(val)).expanduser()
    return p if p.is_absolute() else (base / p).resolve()


def parse_task(path) -> Task:
    """解析一份任务文档。

    :param path: ``task.md`` 路径、任务目录，或任务 id
    """
    src = find_task(path)
    where = str(src)
    meta, body = _split_front(src.read_text(encoding="utf-8-sig"))
    root = src.parent

    sections = parse_sections(body, where)
    rows = _norm_rows(meta.get("grid"), len(sections), where)
    fracs = _norm_fracs(meta.get("row_fracs"), rows, where)

    # 模板：任务文档 > 环境变量 > assets/template 下唯一 pptx
    tpl = meta.get("template")
    if tpl:
        template = _resolve(root, tpl)
    else:
        template = C.template_path()
    if not template.is_file():
        raise FileNotFoundError("%s: 模板不存在 %s" % (where, template))

    # 配图目录：默认 task.md 同级的 assets/
    fdir = _resolve(root, meta.get("figures_dir") or "assets")
    for s in sections:
        if not s.fig_rows:                      # 只有 caption 没写 figure 的兼容路径
            s.fig_rows = [s.figures] if s.figures else []
        for row in s.fig_rows:
            paths: list[Path] = []
            for name in row:
                p = (fdir / name).resolve()
                if not p.is_file():
                    raise FileNotFoundError(
                        "%s: 小节「%s」的配图不存在: %s\n  配图目录 = %s"
                        % (where, s.title, name, fdir))
                paths.append(p)
            s.fig_rows_paths.append(paths)
            s.figure_paths.extend(paths)

    area = meta.get("area")
    if area in (None, "auto", "Auto", "AUTO"):
        area = None                       # 交给 template.detect_area 探测
    else:
        if not isinstance(area, (list, tuple)) or len(area) != 4:
            raise ValueError("%s: area 必须是 'auto' 或 [x0, y0, x1, y1]，得到 %r"
                             % (where, area))
        area = tuple(float(x) for x in area)

    output = str(meta.get("output") or C.OUT_NAME)
    if not output.lower().endswith(".pptx"):
        output += ".pptx"

    return Task(
        id=str(meta.get("id") or root.name),
        title=str(meta.get("title") or meta.get("name") or root.name),
        source=src,
        root=root,
        template=template,
        output=output,
        dist=_resolve(root, meta.get("dist") or C.DIST),
        figures_dir=fdir,
        sections=sections,
        rows=rows,
        row_fracs=fracs,
        area=area,
        style=meta.get("style") or {},
        meta=meta,
    )


# ---------------------------------------------------------------- 任务定位
def _task_docs() -> list[Path]:
    """tasks/ 下的全部任务文档（下划线开头的是骨架/草稿，不算任务）。"""
    return sorted(p for p in C.TASKS.glob("*/task.md")
                  if p.is_file() and not p.parent.name.startswith("_"))


def find_task(ref) -> Path:
    """把「任务 id / 目录 / task.md 路径」统一成 ``task.md`` 路径。"""
    if ref is None:
        cands = _task_docs()
        if len(cands) == 1:
            return cands[0]
        if not cands:
            raise FileNotFoundError("tasks/ 下没有任何任务（每个任务是 tasks/<id>/task.md）")
        raise ValueError("tasks/ 下有多个任务，请指定其一：%s"
                         % ", ".join(p.parent.name for p in cands))
    p = Path(str(ref)).expanduser()
    if p.is_dir():
        p = p / "task.md"
    elif not p.suffix and not p.exists():
        p = C.TASKS / str(ref) / "task.md"
    if not p.is_file():
        raise FileNotFoundError("找不到任务文档: %s" % p)
    return p.resolve()


def list_tasks() -> list[tuple[str, str]]:
    """列出 ``(id, 标题)``。标题解析失败也不影响列出。"""
    out = []
    for p in _task_docs():
        try:
            t = parse_task(p)
            out.append((t.id, t.title))
        except Exception as e:                                  # noqa: BLE001
            out.append((p.parent.name, "<解析失败: %s>" % e))
    return out


def new_task(task_id: str, root: Path | None = None) -> Path:
    """从 ``tasks/_template.md`` 复制出一份新任务文档。"""
    root = root or C.TASKS
    tpl = root / "_template.md"
    if not tpl.is_file():
        raise FileNotFoundError("缺少任务模板: %s" % tpl)
    dst_dir = root / task_id
    if dst_dir.exists():
        raise FileExistsError("任务已存在: %s" % dst_dir)
    (dst_dir / "assets").mkdir(parents=True)
    shutil.copyfile(tpl, dst_dir / "task.md")
    return dst_dir / "task.md"
