# -*- coding: utf-8 -*-
"""命令行入口。

    python build.py                  # 构建 tasks/ 下唯一的任务
    python build.py demo        # 指定任务（id / 目录 / task.md 路径）
    python build.py --list           # 列出所有任务
    python build.py --new my-poster  # 从模板新建一个任务

流水线：读任务文档 → 配图预处理 → 版面求解 → 写 PPTX → 渲染预览。
"""
from __future__ import annotations

import argparse
import sys
import time

from . import config as C
from . import layout as L
from . import preview as P
from . import spec
from . import template as T
from .imaging import prepare_figures
from .writer import build_pptx


def prepare(task_ref=None):
    """解析任务文档并把任务级配置应用到这个进程。"""
    task = spec.parse_task(task_ref)
    C.configure(task)
    if task.area is None:                     # area: auto -> 由模板占位框推导
        area = T.detect_area(task.template)
        if area is None:
            print("警告：模板里没找到「作品简介」占位框，内容区退回 %s"
                  % (C.FALLBACK_AREA,), file=sys.stderr)
            area = C.FALLBACK_AREA
        C.set_area(*area)
    C.set_canvas(*T.slide_size_cm(task.template))
    return task


def build(task_ref=None, with_preview: bool = True, out_name: str | None = None,
          verbose: bool = True) -> dict:
    t0 = time.time()
    task = prepare(task_ref)
    C.ensure_dirs()

    if verbose:
        print(task.summary())
        print("[1/4] 预处理配图 ...")
    figs = prepare_figures(task)
    if verbose:
        for i, (sec, bf) in enumerate(zip(task.sections, figs)):
            if bf.empty:
                print("      #%d 无配图（纯文本块）" % i)
            else:
                print("      #%d 图条 %d 张 [%s]  单位高宽 Σ(w/h)=%.2f"
                      % (i, len(bf.paths), "、".join(sec.figures), bf.total))

    if verbose:
        print("[2/4] 求解版面 ...")
    lay = L.build_layout(task, figs, verbose=verbose)

    out_path = task.dist / (out_name or task.output)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    if verbose:
        print("[3/4] 写 PPTX ...")
    n_shapes = build_pptx(lay, task, figs, task.template, out_path)

    if verbose:
        print(L.describe(lay, task, figs))
        print("      shapes=%d | 网格底 %.2f（内区底 %.2f）"
              % (n_shapes, lay.rows[-1].y + lay.rows[-1].bh, C.Y1 - C.PAD_B))

    result = {
        "task": task.id,
        "pptx": out_path,
        "body_pt": lay.body_pt,
        "ih_cap": lay.ih_cap,
        "total_h": lay.total,
        "shapes": n_shapes,
        "elapsed": time.time() - t0,
        "task_obj": task,
        "figs": figs,
        "layout": lay,
    }

    if with_preview:
        if verbose:
            print("[4/4] 渲染预览 ...")
        p_intro, p_full = P.save_previews(lay, task, figs)
        result["preview_intro"] = p_intro
        result["preview_full"] = p_full
        if verbose:
            print("      %s" % p_intro)
            print("      %s" % p_full)

    if verbose:
        print("完成：%s（%.1fs）" % (out_path, result["elapsed"]))
    return result


def check(task_ref=None, built: dict | None = None, quiet: bool = False) -> int:
    """跑一遍合规自检，打印报告并落盘。返回硬性未通过项数。"""
    from . import compliance as K

    built = built or build(task_ref=task_ref, with_preview=True, verbose=not quiet)
    task, figs, lay = built["task_obj"], built["figs"], built["layout"]
    ctx = K.load(task, figs, lay, built["pptx"])
    results = K.run(ctx)
    text = K.report(results, task, ctx)
    out = task.dist / ("compliance-%s.md" % task.id)
    # 固定 LF：仓库 core.autocrlf=false，若按平台默认写 CRLF，
    # 每次构建都会让这份 Markdown 在工作区里变脏（内容其实没变）。
    out.write_text(text + "\n", encoding="utf-8", newline="\n")
    if quiet:
        bad = K.failures(results)
        for r in bad:
            print("FAIL %s %s — %s" % (r.item.id, r.item.title, r.detail))
        print("合规自检：硬性 %d 项未通过（详见 %s）" % (len(bad), out))
    else:
        print(text)
        print("报告已写入：%s" % out)
    return len(K.failures(results))


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(
        prog="poster",
        description="通用墙报排版引擎：读一份任务文档（tasks/<id>/task.md），"
                    "把图文自动排版并填进官方模板 pptx。")
    ap.add_argument("task", nargs="?", default=None,
                    help="任务：id / 目录 / task.md 路径（省略则取 tasks/ 下唯一的任务）")
    ap.add_argument("-o", "--out", default=None, help="覆盖输出文件名（写入任务 dist/）")
    ap.add_argument("--no-preview", action="store_true", help="跳过预览渲染")
    ap.add_argument("-q", "--quiet", action="store_true", help="只输出错误")
    ap.add_argument("-l", "--list", action="store_true", help="列出所有任务")
    ap.add_argument("--check", action="store_true",
                    help="构建后跑合规自检（清单见 docs/compliance.md），报告写入任务 dist/")
    ap.add_argument("--new", metavar="ID", default=None,
                    help="从 tasks/_template.md 新建一个任务")
    args = ap.parse_args(argv)

    if args.list:
        rows = spec.list_tasks()
        if not rows:
            print("tasks/ 下还没有任务")
        for tid, title in rows:
            print("  %-16s %s" % (tid, title))
        return 0

    if args.new:
        try:
            p = spec.new_task(args.new)
        except (FileNotFoundError, FileExistsError) as e:
            print("错误：%s" % e, file=sys.stderr)
            return 2
        print("已创建任务文档：%s\n编辑它，然后运行：python build.py %s" % (p, args.new))
        return 0

    try:
        built = build(task_ref=args.task, with_preview=not args.no_preview,
                      out_name=args.out, verbose=not args.quiet)
        if not args.check:
            return 0
        n_bad = check(args.task, built, quiet=args.quiet)
    except (FileNotFoundError, ValueError, KeyError) as e:
        print("错误：%s" % e, file=sys.stderr)
        return 2
    return 2 if n_bad else 0


if __name__ == "__main__":
    raise SystemExit(main())
