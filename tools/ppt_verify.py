# -*- coding: utf-8 -*-
"""渲染器真值复核（命令行薄壳）。

真正的判定逻辑在 :mod:`poster.renderer_check` —— 那里也被合规清单的 F5 复用。

用法：
    .venv/Scripts/python.exe tools/ppt_verify.py [pptx] [--tol 0.02]
退出码 0 = 通过；1 = 有文字压图 / 溢出；2 = 本机做不了复核（没 PowerPoint）。
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from poster import renderer_check as RC                       # noqa: E402


def main() -> int:
    ap = argparse.ArgumentParser(description="用 PowerPoint 复核产物有没有文字被图压住")
    ap.add_argument("pptx", nargs="?", default=None)
    ap.add_argument("--tol", type=float, default=0.02, help="容差 cm（默认 0.02）")
    ap.add_argument("-q", "--quiet", action="store_true")
    a = ap.parse_args()

    if not RC.available():
        print("本机没有 PowerPoint / PowerShell，跳过渲染器复核", file=sys.stderr)
        return 2

    pptx = Path(a.pptx) if a.pptx else next((ROOT / "dist").glob("*.pptx"))
    print("复核：%s" % pptx)
    if not a.quiet:
        for s in RC.dump_shapes(pptx):
            if s.has_tf and s.tlen > 0:
                print("  %-16s 框高 %.3f 墨迹 y %.3f..%.3f（高 %.3f，%d字）"
                      % (s.name, s.h, s.by, s.b_bottom, s.bh, s.tlen))
    bad = RC.verify(pptx, a.tol)
    if bad:
        print("\n发现 %d 处问题：" % len(bad))
        for b in bad:
            print("  ✗ " + b)
        return 1
    print("\n✓ 渲染器复核通过：没有文字被配图盖住，也没有文字溢出自己声明的框")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
