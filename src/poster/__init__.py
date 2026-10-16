# -*- coding: utf-8 -*-
"""poster —— 通用墙报（学术海报）排版引擎。

一份墙报 = 一份任务文档（``tasks/<id>/task.md``）。引擎读文档、算版面、填模板。

模块划分
--------
- :mod:`poster.spec`       任务文档解析（Markdown + YAML front-matter）
- :mod:`poster.template`   模板探测（画布 / 内容区 / 背景图 / 占位框）
- :mod:`poster.config`     默认路径 + 版面几何 + 配色（可被任务覆盖）
- :mod:`poster.typography` 字体缓存 + 分词 + 断行量算
- :mod:`poster.imaging`    配图去透明底 + 裁白边
- :mod:`poster.layout`     网格求解 + 排法分派 + 尺寸搜索
- :mod:`poster.writer`     写 PPTX
- :mod:`poster.compliance` 合规自检（清单见 ``docs/compliance.md``）
- :mod:`poster.renderer_check`  真渲染器复核（PowerPoint COM 读文字墨迹框）
- :mod:`poster.repro`      可复现构建（归一化 zip 时间戳）
- :mod:`poster.preview`    PIL 预览渲染
- :mod:`poster.cli`        命令行入口
"""
from .cli import build, main, prepare  # noqa: F401
from .config import OUT_NAME, PT2CM, PX_PER_CM, template_path  # noqa: F401
from .spec import Section, Task, parse_task  # noqa: F401

__all__ = ["build", "main", "prepare", "parse_task", "Task", "Section",
           "template_path", "OUT_NAME", "PT2CM", "PX_PER_CM"]
__version__ = "2.8.0"
