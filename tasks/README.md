# tasks/ —— 每个墙报任务一个目录

```
tasks/
├── README.md            本文件
├── _template.md         新任务的骨架（python build.py --new <id> 会复制它）
└── demo/                ★ 随仓库发布的示例任务
    ├── task.md              任务文档：排版设置 + 全部图文
    └── assets/              该任务的模板与配图（由 tools/make_demo.py 合成）
```

> **本仓库只把 `demo` 与骨架入库。** 具体墙报的文案与素材是作品内容，
> 由 `.gitignore` 排除（`/tasks/*` 只放行 `README.md`、`_template.md`、`demo/`）。
> 详见根目录 README 的「仓库收录范围」。

## 常用命令

```bash
python build.py --list                # 列出所有任务（被 .gitignore 排掉的不算数）
python build.py --new my-poster       # 新建任务
python build.py demo                  # 构建（产物进 dist/，同时出预览图）
python build.py demo --check          # 构建 + 合规自检
python build.py demo --no-preview -q  # 只要产物、不要预览、别啰嗦
```

## 写新墙报

1. `python build.py --new my-poster`
2. 把模板与配图丢进 `tasks/my-poster/assets/`
3. 编辑 `tasks/my-poster/task.md`：
   - `template:` 指向模板 pptx（相对本文档）
   - 按 `## 小标题` 划分分块，每块一段正文
   - 有素材的块写 `<!-- figure: 文件 | 图注 -->`；多张图用**分号**分隔，
     每张各写一条图注（一个块最多 3 张，版面上**等高并排成一「图条」**）
   - 关键数据用 `**两个星号**` 标红
4. `python build.py my-poster`，看 `dist/preview_intro.png` 是否满意，
   不满意就调 `grid` / `mode` / `style`
5. 交付前 `python build.py my-poster --check`，**硬性项全绿**再发

格式细节全在 [`../docs/task-format.md`](../docs/task-format.md)；
版式为什么会这样，见 [`../docs/design.md`](../docs/design.md)；
想照着一个能跑的样例改，就从 `tasks/demo/` 复制一份。

## 注意

- 目录名就是默认的任务 id；`task.md` 是唯一入口，**不要改名**。
- 每个任务独占 `.cache/<任务id>/` 存配图预处理中间产物，互不干扰。
- `area: auto` 依赖模板里的「作品简介」占位框（判定规则见 `config.PH_TOP` 等）。
  **换模板后先看构建日志里打印的内容区是否合理**，不合理就在文档里写死
  `area: [x0, y0, x1, y1]`。
- `tasks/` 下有多个任务时，`python build.py` 不带参数会报错 —— 必须指名道姓
  （测试也因此显式写 `TASK = "demo"`，免得本机多放一个任务就红）。
