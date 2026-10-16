# 渲染器行为笔记

> 这份文档回答一个问题：**为什么"按模型算对了"还不够，非得让 PowerPoint 亲自确认？**
>
> 起因是一次真实的版面事故：产物在 PowerPoint 里打开时，某个分块的末行整行**钻进配图底下**
> （只露出上半截）；而引擎自带的几何自查报"发现 0 个问题"。根因不是几何算错，
> 而是**引擎与渲染器对同一份 XML 的解读不一致**。

工程侧的四层守卫见 [`compliance.md`](compliance.md) 的「渲染器一致性」一节
（`C8` / `C9` / `C10` / `C11` / `F5`）。本文记**怎么量出来的**。

---

## 1. 已确认的量化规则

`1 pt = 1/72 inch`，但 PowerPoint **不是照 XML 原样画的**，它会把一部分值量化。
下面全部是用 COM 读 `TextRange.Bound*` 实测出来的（量法见第 2 节），
字号 3.45pt 的样本：

| 写法 | 实测行距 | 结论 |
|---|---|---|
| `a:spcPts val="400"` | 4.0000 pt | — |
| `a:spcPts val="500"` | 5.0000 pt | — |
| `a:spcPts val="540"` | 5.0000 pt | 540 → 5.0（5.4 取整到 5） |
| `a:spcPts val="550"` | **6.0000 pt** | 550 → 6.0（5.5 进位到 6） |
| `a:spcPts val="552"` | **6.0000 pt** | ← 引擎当年写的就是这个（想要 5.52） |
| `a:spcPts val="560"` | 6.0000 pt | |
| `a:spcPts val="600"` | 6.0000 pt | |
| `a:spcPct val="100000"` | 4.1400 pt | = 3.45 × **1.20** |
| `a:spcPct val="133000"` | 5.5050 pt | = 3.45 × 1.20 × 1.33，**精确** |
| `a:spcPct val="140000"` | 5.7950 pt | 精确 |
| `a:spcPct val="150000"` | 6.2100 pt | 精确 |

两条结论：

1. **`spcPts` 被四舍五入到整点。** 想要 5.52pt 却实排 6.0pt —— 一行高 8.7%，
   九行变十行，第十行就溢到框外。**所以行距一律写 `spcPct`。**
2. **`spcPct` 的 100% = 字号 × 1.20**（这是 PowerPoint 语义的"单倍行距"，
   不是字号本身）。换算入口：`typography.line_spacing_pct()`，
   基准常量 `config.LINE_REF_FACTOR = 1.20`。

`sz`（字号）同样会被抬档：实测 3.45pt 与 3.50pt 量到**同一个字宽**
（每 16 个全角字 57.000pt），3.40pt 是 55.000pt。所以"精确到 0.05pt 调字号"
在这类渲染器上是没有意义的 —— 只能靠 `spcPct` 把行距做准。

### 断行：`push` / `hang` / `pull`

- **`push`**：行尾放不下的标点（`，。、`）**推到下一行** → 行首出现标点。
- **`hang`**：标点挤在本行末尾、悬挂出边界 → 行首干净。
  **只有在 `bodyPr/@hangingPunct="1"` 时才成立。**
- **`pull`**：行首禁则 —— 标点不能落行首，于是把它**连同前一个单元**
  一起移到下一行。这是渲染器在悬挂标点也挤不下时的真实兜底，**行数最多**。

引擎画预览时用 `hang`（好看），于是必须在 pptx 里也写上 `hangingPunct="1"`，
否则渲染器按 `push` 排。配置开关 `config.HANG_PUNCT`。
**求解尺寸时三种策略都要试、取行数最悲观的那个**（`typography.lines_count_worst`）——
漏掉 `pull` 会少算一行（2026-10-16 实际发生：估算 10 行、实排 11 行，末行压到图上）。

另有一个**与策略无关**的宽度偏差：Pillow 量出的行宽比渲染器实排窄约 1.5%，
由 `config.MEASURE_SAFETY` 打折抵掉，实测需 ≤ 0.97 才兜得住临界行。

---

## 2. 怎么量：用 `<a:br/>` 消掉行框高

`BoundHeight` 是"文字墨迹框高"，它满足

```
H(n) = (n - 1) × 行距 + 行框高
```

其中**行框高是未知的**（跟字体度量有关）。但只要你**强制**折行数，就可以用两行数相减
把它消掉：

```
行距 = (H(n2) - H(n1)) / (n2 - n1)
```

强制折行的办法是在段落里插 `<a:br/>`：行数就由我们说了算，而不是由渲染器断行决定。
字宽同理：一整行放 16 个全角字，`BoundWidth / 16` 就是每字推进宽。

工具：

```bash
# 1) 生成探针 pptx（一堆变体：不同 lnSpc × 2 / 6 行，不同字号 × 一行 16 字）
.venv/Scripts/python.exe tools/probe_metrics.py tools/_probe/probe_metrics.pptx

# 2) 让 PowerPoint 打开它、用 COM 把 BoundHeight / BoundWidth 读回来
powershell -NoProfile -File tools/read_probe.ps1 tools/_probe/probe_metrics.pptx
```

中间产物写在 `tools/_probe/`（已 gitignore）。**注意 `read_probe.ps1` 里的除数（16）
必须与 `probe_metrics.py` 的 `GLYPHS` 保持一致**，否则算出的字宽会整体偏一档。

---

## 3. 怎么复核产物：让渲染器报墨迹框

```bash
python tools/ppt_verify.py dist/demo-poster.pptx
# 退出码：0 通过 / 1 发现问题 / 2 本机没装 PowerPoint
```

底层是 `tools/ppt_dump.ps1`：COM 打开产物，把每个形状的
`left/top/width/height` 与 `TextRange.BoundLeft/BoundTop/BoundWidth/BoundHeight`
导成 TSV，再由 `src/poster/renderer_check.py` 判两件事：

1. **文字压到图**：某个文本框的墨迹框与某张配图的框相交；
2. **文字溢出声明框**：墨迹高 > 该文本框自己声明的高度（"迟早会溢出去"）。

**取巧之处**：不看引擎自己的布局结果，只看**渲染器报回来的墨迹框** ——
这样"模型算错了"这件事本身就会被暴露，而不是被同一套错误模型再确认一遍。

### 一个必须记住的边界：模板形状不归引擎管

写 pptx 是**只追加**的：产物里前 `n_tpl` 个形状（`ctx.n_tpl`）原封不动来自模板
（作品名称 / 项目成员 / 作品编号 …）。它们的框高、字号是模板作者定的 ——

- **"溢出声明框"只判引擎排进去的文字**（否则换一份字段排得紧的模板就一片假警报）；
- **"文字压到图"对全部文字生效**（万一我们排的图盖住了模板字段，那是真事故）。

`C8` / `C9` / `C11` / `F5` 都遵循这条界线。

---

## 4. 环境相关

- **没有 PowerPoint 就跳过。** `renderer_check.available()` 检查
  平台是 Windows、`tools/ppt_dump.ps1` 在、`powershell` 在。任一不满足，
  合规 `F5` 记为「跳过」而**不是**「不合格」—— 在别的机器上不该因此红。
- **Windows PowerShell 5.1 按 ANSI（本地代码页）读 `.ps1`。**
  脚本里写中文注释会导致解析失败（`MissingEndCurlyBrace` 之类），
  所以 `tools/*.ps1` **一律纯 ASCII**。
- **PowerShell 的 stdout 是本地代码页（中文 GBK），不是 UTF-8。**
  Python 侧调它时**不要**用 `text=True`，要 `capture_output=True` 后自己
  `.decode("utf-8", "replace")`（见 `renderer_check.dump_shapes`）。
- **传给 PowerShell 的路径要先 `Path(x).resolve()`**：相对路径会以 PowerShell 的
  当前目录解析，容易报 `DirectoryNotFoundException`。
- **`a:bodyPr` 的自动调整子元素是三选一**：`noAutofit` / `normAutofit` / `spAutoFit`。
  **同时挂两个**会让 PowerPoint 判定整份文件损坏、拒绝打开（`HRESULT:0x80070570`）。
  要挂之前先把三种都删干净（`writer._force_no_autofit`）。
