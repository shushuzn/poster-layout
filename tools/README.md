# tools —— 独立小工具

不参与 `build.py` 主流水线，按需手动调用。

## `make_demo.py` —— 生成示例任务的素材

`tasks/demo/` 里的模板与占位配图都是**合成**的，由本脚本生成：

```bash
.venv/Scripts/python.exe tools/make_demo.py
```

产出 `tasks/demo/assets/` 下的 `template.pptx`（8×20cm，含可被 `area: auto`
探到的「作品简介」占位框）与 4 张占位图（4:1 流程条 / 1.5:1 阵列 /
1:1 分布 / 2:1 柱状，分别对应四种排法）。想换一套演示素材，重跑即可。

**为什么是生成器而不是手工塞文件**：示例模板必须满足
`template.is_placeholder()` 的判定条件，否则 `area: auto` 探不到内容区 ——
这件事由代码保证，比手抄坐标可靠；而且仓库里不该出现来源不明的二进制素材。

## `edsdk_call.py` —— 以自定义超时调用 editor_sdk

调用本机 `editor_sdk` 的 MCP 工具，用于把素材文档（`.doc` / `.docx`）
渲染成逐页图片、或把内嵌图提取出来。

**为什么需要它**：editor_sdk 默认客户端超时只有 30s，而几十 MB 的文档做
`doc_to_image` / `doc_get_images` 时上传 + 转换远超这个时限，必然失败。

```bash
python tools/edsdk_call.py <工具名> '<JSON 参数>' [超时秒数]

# 例：把素材文档渲染成逐页图片（默认 600s）
python tools/edsdk_call.py doc_to_image '{"file_path":"D:/path/to/material.doc"}' 600

# 例：列出文档结构
python tools/edsdk_call.py doc_resolve_document_structure '{"file_path":"D:/path/to/material.doc"}'
```

会自动在 `127.0.0.1:39099–39108` 探测 editor_sdk 端点（并绕过系统代理）。

> 这是**素材准备阶段**的工具，正常构建不需要它；也不会随仓库分发任何素材。

## `ppt_verify.py` + `ppt_dump.ps1` —— PowerPoint 实排复核

合规清单里的 **F5** 就是它：用 PowerPoint COM 打开产物，读**渲染器自己算出的文字墨迹框**，
判「文字压图」与「文字溢出声明框」。

```bash
.venv/Scripts/python.exe tools/ppt_verify.py [dist/demo-poster.pptx]
# 退出码：0 通过 / 1 发现问题 / 2 本机没装 PowerPoint
```

**为什么要它**：`spcPts` 会被 PowerPoint 取整到整点（写 5.52pt 实排 6.0pt），
所以「按我们自己的模型算一遍」证明不了渲染结果 —— 2026-10-14 的事故就是自查全绿、
PowerPoint 里却压字。必须让真渲染器说话。

`ppt_dump.ps1` 是它的底层：把每个形状的几何量导出成 TSV 供 Python 解析。
**全 ASCII**：Windows PowerShell 5.1 按 ANSI 读 `.ps1`，中文注释会导致解析失败。

## `probe_metrics.py` + `read_probe.ps1` —— 量渲染器的真行距 / 真字宽

一次性探针。**结论已固化进 `config.py` / `typography.py`**，
并记在 [`docs/rendering-notes.md`](../docs/rendering-notes.md)；
需要重新量渲染器口径时才跑：

```bash
.venv/Scripts/python.exe tools/probe_metrics.py tools/_probe/probe_metrics.pptx
powershell -NoProfile -File tools/read_probe.ps1 tools/_probe/probe_metrics.pptx
```

**技巧**：用 `<a:br/>` **强制**折行数，于是行数我们说了算；`BoundHeight =
(n−1) × 行距 + 行框高`，两个不同行数**相减**就把未知的行框高消掉，直接得行距：

```
行距 = (H(n2) − H(n1)) / (n2 − n1)        每字宽 = BoundWidth / 16
```

`read_probe.ps1` 里的除数（16）必须与 `probe_metrics.py` 的 `GLYPHS` 一致。

实测结论：`spcPts 552` → 6.0000pt（被取整）；`spcPct 133000` → 5.5050pt（精确）；
`sz` 也会被抬档（3.45 与 3.50 量到同一字宽）。
中间产物写到 `tools/_probe/`（已 gitignore）。
