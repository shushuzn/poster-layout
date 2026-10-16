# 通用墙报排版工程

把一堆图文按示例版式**自动排版**并填进给定模板 PPTX 的工具。**一个墙报 = 一份任务文档**，
引擎不认识任何具体赛事 —— 换比赛、换模板，写一份新的 `task.md` 就行，不动代码。

给它一份任务文档（几段正文 + 几张配图 + 一行网格），它按示例的
「图主文辅图墙 + 虚线弧边分块 + 胶囊小标题」语言，自动求解字号与图片尺寸、
逐块分派图文排法，产出成品 pptx 与预览图。

> 正文字数是这张海报的**第一自变量**：字越多 → 字号被压得越小 → 上了墙谁也看不清。
> 引擎因此对正文字数设了硬上限（单块 220 / 全篇 800 汉字，见合规项 `E6`，
> 可用任务的 `style:` 覆盖）。**超限时删铺垫句，不要放宽阈值** —— 海报是看图为主、
> 站着远看的媒介，不是能凑近读的小字报。

> 仓库里随附一份**示例任务** [`tasks/demo`](tasks/demo/task.md)：它的模板与配图
> 都由 [`tools/make_demo.py`](tools/make_demo.py) 合成，文案是示意内容，
> 用来演示任务文档的写法与各种排法。构建它：
>
> ```bash
> python build.py demo            # -> dist/demo-poster.pptx
> ```
>
> 本仓库**只发布引擎**：具体墙报的文案、素材、模板与成品都不入库（见
> [仓库收录范围](#仓库收录范围)）。

---

## 快速开始

```bash
python -m venv .venv
.venv/Scripts/python.exe -m pip install -r requirements.txt    # Windows
# source .venv/bin/activate && pip install -r requirements.txt # macOS/Linux

python build.py --list          # 看有哪些任务
python build.py demo            # 构建指定任务
python build.py                 # tasks/ 下只有一个任务时可省略
```

装了 `make` 的话：

```bash
make               # 构建（TASK=<id> 可指定任务）
make check         # 构建 + 跑合规自检
make test          # 跑测试
make list          # 列出任务
make new ID=my-poster
```

产物（写在任务的 `dist:` 目录下）：

| 文件 | 说明 |
|---|---|
| `<任务 output>` | 交付物（模板填充后的成品 pptx） |
| `preview_intro.png` | 内容区局部预览 |
| `preview_full.png` | 整张墙报预览 |
| `compliance-<任务id>.md` | 合规自检报告（`--check` 生成） |

---

## 一个墙报任务长什么样

```
tasks/my-poster/
├── task.md        ← 唯一要写的东西
└── assets/        ← 这个任务的配图（+ 模板，若不想放共享目录）
```

```markdown
---
id: my-poster
title: 我的墙报
template: "assets/template.pptx"
output: my-poster.pptx
dist: ../../dist
figures_dir: assets
area: auto
grid:
  - [0]
  - [1]
  - [2, 3]
---

## 背景
没有对应素材的小节就不配图。

## 方法
<!-- figure: fig_a.png | 图1  流程 ; fig_b.png | 图2  样品 -->
正文……关键数据用 **两个星号** 包起来就标红。

## 结果
<!-- figure: fig_c.png | 图3  形貌 -->
正文……
```

- `## 小标题` → 版面上一个分块，顺序即展示顺序
- `<!-- figure: 文件 | 图注 ; 文件 | 图注 -->` → 给这一块配 1~3 张图，多张**等高并排成一「图条」**，
  **一张图一条图注**（分号分隔，不许几图共用一注）
- `**文字**` → 标红
- 想控制排法就加 `<!-- mode: stack|side|text -->`，不写就由图条的宽高比自动决定

**完整格式说明：[`docs/task-format.md`](docs/task-format.md)　·　合规要求：[`docs/compliance.md`](docs/compliance.md)**

---

## 目录结构

```
.
├── build.py                  # 入口
├── skills/                   # 随仓库发布的 Skill（AI 助手的操作手册，见下）
│   └── poster-layout/SKILL.md
├── tasks/                    # ★ 所有墙报任务
│   ├── _template.md          # 新任务骨架
│   ├── README.md
│   └── demo/                 # 示例任务（合成模板 + 占位配图 + 示意文案）
├── src/poster/               # 排版引擎
│   ├── spec.py               # 任务文档解析（Markdown + YAML front-matter）
│   ├── template.py           # 模板探测（画布 / 内容区 / 背景图 / 占位框）
│   ├── config.py             # 默认路径 + 版面几何 + 配色（可被任务覆盖）
│   ├── typography.py         # 字体缓存 + 分词 + 断行 / 行距量算
│   ├── imaging.py            # 配图去透明底 + 裁白边 + 量宽高比（图条）
│   ├── layout.py             # 网格求解 + 排法分派 + 尺寸搜索
│   ├── writer.py             # 写 PPTX
│   ├── compliance.py         # 合规自检（清单见 docs/compliance.md）
│   ├── renderer_check.py     # 真渲染器复核（PowerPoint COM 读文字墨迹框）
│   ├── repro.py              # 可复现构建（归一化 zip 时间戳）
│   ├── preview.py            # PIL 预览渲染
│   └── cli.py                # 命令行
├── dist/                     # 产物 + 预览 + 合规报告（不入库）
├── tests/                    # 测试（runner.py 无需 pytest）
├── tools/                    # 独立小工具
│   ├── make_demo.py          # 生成示例任务的模板与占位配图
│   ├── ppt_verify.py         # 命令行跑 PowerPoint 实排复核
│   ├── ppt_dump.ps1          # COM 导出形状几何到 TSV
│   └── probe_metrics.py + read_probe.ps1   # 量渲染器真行距 / 真字宽
├── docs/                     # task-format.md / compliance.md / design.md
│                             # / rendering-notes.md
└── legacy/                   # 历史版本脚本与文案草稿（本地归档）
```

---

## 工作流

```
tasks/<id>/task.md
      │  ①解析：front-matter → 任务设置；## 小节 → 分块；figure → 配图（可多张）
      ▼
tasks/<id>/assets/*.png
      │  ②去透明底 + 裁白边（缓存到 .cache/<id>/），量出每张图的 w/h
      ▼
  图条形态（Σ w/h）+ 文档里的 mode → 逐块分派排法
      │
      ▼
  layout 求解 ── 字号 / 图高上限 / 每块几何（先保配图，再取字号）
      │
      ├──③ writer.py  → dist/<output>.pptx
      ├──④ preview.py → dist/preview_*.png
      └──⑤ compliance.py → 逐条对照合规清单，报告写 dist/compliance-<id>.md
```

---

## 合规自检（每份墙报发布前必跑）

```bash
python build.py demo --check      # 或 make check
```

构建完立刻逐条对照 [`docs/compliance.md`](docs/compliance.md) 的清单（**40 项**：
硬性 35 / 提示 2 / 人工 3），把报告写到 `dist/compliance-<任务id>.md`。
**硬性项全部通过则退出码 0，否则 2** —— 可以直接串进脚本，红了就别提交。

清单覆盖：范围（模板字段一字不动）、尺寸（画布 / 字号 / 不溢出 / 图不压成细缝）、
版式（虚线圆角框 / 渐变 / 胶囊骑线 / 版面节奏 / 图文同框）、
渲染器一致性（**行距写 `spcPct`** / 文字需要高 ≤ 声明框高 / 自动调整至多一个 /
悬挂标点与断行策略一致 / 正文不用两端对齐）、
图文对位（**一张图一条图注**、图必须落在声明它的块内、图号连续、图注居中在图下）、
文字（无缺失字形 / 无未完成痕迹 / **正文字数有上限**）、交付（产物可打开 / 构建可复现 /
**PowerPoint 实排复核无压图**）。

> `F5` 会用 PowerPoint COM 打开产物，读**渲染器自己算出的文字墨迹框**，判「文字压图 / 溢出」。
> 本机没装 PowerPoint 时该项显示为**跳过**，不算不合格；也可手跑 `python tools/ppt_verify.py`。
> 为什么非要真渲染器：`spcPts` 会被 PowerPoint 取整到整点，**模型自洽证明不了渲染结果**
> —— 细节见 [`docs/rendering-notes.md`](docs/rendering-notes.md)。

一条新要求怎么加：在 `src/poster/compliance.py::REQUIREMENTS` 里加 `Item`，
同时在 `docs/compliance.md` 补一行 —— 两边不一致会**直接测试失败**。

---

## Skill：给 AI 助手的操作手册

[`skills/poster-layout/SKILL.md`](skills/poster-layout/SKILL.md) 把「怎么用这个引擎」写成一份
可被 AI 助手直接加载的 Skill：什么时候该用它、动手前必须先读的不变式、从备料到交付的七步流程、
`task.md` 速查表、交付前检查单，以及会复发的坑。

**它和 README 是两种读者**：README 给人看（先讲设计意图与全貌），SKILL 给助手看
（先讲"别踩什么"和"从哪一步开始"）。改引擎的公开行为时两边都要同步。

加载方式（任选其一）：

```bash
cp -r skills/poster-layout .workbuddy/skills/     # 项目级：在本仓库开发时生效
cp -r skills/poster-layout ~/.workbuddy/skills/   # 用户级：所有项目都能用
```

版本号三处一致（`pyproject.toml` / `src/poster/__init__.py` / `CHANGELOG.md` 首条）由
`tests/test_layout.py::test_release_version_is_in_sync` 守着，`python build.py --version` 可查。

---

## 仓库收录范围

本仓库发布的是**引擎**，不是某一次比赛的成品。`.gitignore` 明确排除：

| 排除项 | 为什么 |
|---|---|
| `/tasks/*`（只放行 `tasks/demo`、`tasks/_template.md`、`tasks/README.md`） | 具体墙报的文案与素材是**作品内容**，不该随引擎公开 |
| `/assets/` | 官方模板、参考样张、素材文档逐页渲染 —— 赛事与第三方材料 |
| `/dist/` | 成品 pptx 含作品内容；报告与预览都能重建 |
| `/legacy/` | 历史脚本、文案草稿、旧交付物 |
| `/.workbuddy/` | 内部工作记录 |

被排除的文件**留在本机磁盘上**，只是不进版本库；想换成公开，删掉 `.gitignore`
里对应几行再 `git add` 即可。想让引擎自己带一份能跑的样例，
就加一个 `tasks/<id>/` 并确保它只含**自己合成**的素材（照 `tasks/demo` 的样子）。

---

## 环境变量

| 变量 | 作用 |
|---|---|
| `POSTER_TEMPLATE` | 覆盖模板 pptx 路径（优先于任务文档里的 `template:`） |
| `POSTER_DIST` | 覆盖产物输出目录 |
| `SOURCE_DATE_EPOCH` | 钉住构建时间戳（见下） |

---

## 可复现构建

**同样的输入必然产出同样的字节。** 构建结束时会调用 `repro.freeze_zip()`
把所有 zip 条目的时间戳归一化到 1980-01-01（zip 纪元）。

这不是洁癖：pptx 动辄十几 MB，若每次构建都产出不同字节，每跑一次就在 git 里
多留一份全量副本。本仓库的 `.git` 曾因此从 71MB 涨回 88MB
（详见 [`CHANGELOG.md`](CHANGELOG.md) 的 2.0.0 一节）。

```bash
python build.py demo --no-preview && md5sum dist/demo-poster.pptx
python build.py demo --no-preview && md5sum dist/demo-poster.pptx   # 两次必须一致
```

`tests/test_layout.py::test_build_is_reproducible` 会自动守护这条性质。

---

## 环境要求

- Python **3.10+**
- 依赖：`pillow`、`python-pptx`、`numpy`、`PyYAML`
- 字体：PPTX 里写死「微软雅黑」；预览渲染用 `C:/Windows/Fonts/msyh*.ttc`，
  缺失时退化为 Pillow 内置字体，**只影响预览**，不影响 PPTX 成品。
- 真渲染器复核（合规 `F5`）需要 **Windows + PowerPoint + PowerShell**；缺了就自动跳过。

---

## 模板约束（重要）

模板**通常不是空白**：内容区里往往有一个残留旧文字的占位文本框，构建时会主动清空。
模板上的其他字段（作品名称 / 项目成员 / 本科生组 / 作品编号 之类）**一字不动**
—— 由 `tests/test_layout.py` 自动守护。

内容区默认由该占位框**自动推导**（`area: auto`，判定规则见 `config.PH_TOP` 等）。
换模板后请瞟一眼构建日志里打印的内容区；不对就在任务文档里写死 `area: [x0, y0, x1, y1]`。

---

## 改东西该动哪个文件

| 想做什么 | 改哪里 |
|---|---|
| 改某份墙报的文案 / 标题 / 图注 / 配图 | 该任务的 `tasks/<id>/task.md` |
| 调某份墙报的网格、排法、配色 | 同上（`grid` / `mode` / `style`） |
| 改**所有**墙报的默认版式（几何、配色、字体、搜索区间） | `src/poster/config.py` |
| 改排法分派规则 / 求解策略 | `src/poster/layout.py` |
| 支持新的文档指令、新的 front-matter 字段 | `src/poster/spec.py` |
| 换背景图提取方式 / 内容区探测规则 | `src/poster/template.py` |
| 加 / 改一条合规要求 | `src/poster/compliance.py` + `docs/compliance.md`（两边必须一致） |
| 改「文字压图 / 溢出」的判定口径 | `src/poster/renderer_check.py` |
| 改 AI 助手看到的操作手册 | `skills/poster-layout/SKILL.md`（改完重新复制到 `.workbuddy/skills/`） |

---

## 已知陷阱（踩过，记录备查）

- **`℃` / `ₓ` / `⁺` 微软雅黑缺字形** → 正文一律用 ASCII 等价写法（`°C` / `Fe3+` / `m2`）。
  测试会拦住这些字符。
- **`python-pptx` 的 `gradient_angle` setter 是反的** → 写 90 会落成 `ang=16200000`。
  直接改 `a:lin/@ang`（0=左→右，5400000=上→下）。
- **PIL `anchor="mm"` 让中文偏低约 2px** → 用 `getbbox` 的墨迹框中心定位。
- **RGBA 图直接贴白底会变黑** → 先按 alpha 蒙版合到白底。
- **段落缓冲区忘了清空会让正文翻倍** → `spec._parse_section` 里每段追加后必须重置。
- **`area: auto` 推出来的值会带上亚 0.01cm 的噪声** → 探测时先吸附到 0.01cm。
- **配图挂错块**（图跟着"叙事顺序"走、却和它支撑的正文不在同一块）→ 现在
  `tests/test_layout.py::test_every_figure_sits_in_its_own_block` 用几何包含关系守着：
  每张图必须落在声明它的那个块的矩形里。**图上是什么，看坐标轴/轴标签认，
  不要凭文件名或感觉猜**（认反过一次）。
- **行距写 `spcPts` 会被 PowerPoint 取整到整点**（写 552 = 5.52pt，实排按 6.0pt）→
  一行多 8.7% 高，整块多出一行，**末行就钻进配图底下**。行距一律写 `a:spcPct`
  （`typography.line_spacing_pct`，基准 1.20 = PowerPoint 的"单倍行距"），
  并由合规 `C9` 守着。`sz` 同样会被抬档（3.45 与 3.50 量到同一字宽）。
- **自查全绿 ≠ 渲染器不压字**：自查脚本若用**同一套错误模型**重算，只会自证清白。
  必须让真渲染器说话 —— `tools/ppt_verify.py`（PowerShell COM），或用合规 `F5`。
- **`a:bodyPr` 的自动调整三选一**（`noAutofit` / `normAutofit` / `spAutoFit`）：
  同时挂两个会让 PowerPoint 判定文件损坏、**拒绝打开**（HRESULT:0x80070570）。
- **Windows PowerShell 5.1 按 ANSI 读 `.ps1`** → 脚本里写中文注释会导致解析失败，
  `tools/*.ps1` 必须全 ASCII；取回它输出时别加 `text=True`（stdout 是本地代码页）。
- **合规检查要区分"谁排的"**：写 pptx 是只追加的，模板原有形状排在前面。
  它们用不用 `spcPct`、框高够不够，不归引擎管 —— 报它们只会制造假警报
  （`ctx.n_tpl` 划出这条界线）。

版式层面的取舍与试错记录见 [`docs/design.md`](docs/design.md)，
渲染器行为与量测手法见 [`docs/rendering-notes.md`](docs/rendering-notes.md)。
