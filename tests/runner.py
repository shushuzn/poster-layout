# -*- coding: utf-8 -*-
"""测试运行器：不依赖 pytest。

    python tests/runner.py          # 跑全部
    python tests/runner.py spec     # 只跑名字含 spec 的用例
    pytest                          # 装了 pytest 也能直接跑

本目录每个 ``test_*.py`` 都只依赖标准库 + 项目依赖，用例函数命名为 ``test_*``。
"""
from __future__ import annotations

import importlib
import sys
import traceback
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(Path(__file__).resolve().parent))


def discover() -> list[str]:
    return sorted(p.stem for p in Path(__file__).resolve().parent.glob("test_*.py"))


def main(pattern: str | None = None) -> int:
    ok = fail = 0
    for name in discover():
        mod = importlib.import_module(name)
        cases = [(n, f) for n, f in vars(mod).items()
                 if n.startswith("test_") and callable(f)]
        if not cases:
            continue
        print("\n== %s ==" % name)
        for n, fn in sorted(cases):
            label = "%s::%s" % (name, n)
            if pattern and pattern not in label:
                continue
            try:
                fn()
                print("  PASS  %s" % n)
                ok += 1
            except AssertionError as e:
                fail += 1
                print("  FAIL  %s\n        %s" % (n, e))
            except Exception as e:                              # noqa: BLE001
                fail += 1
                print("  ERROR %s\n        %r" % (n, e))
                traceback.print_exc()
    print("\n%d 通过 / %d 失败" % (ok, fail))
    return 1 if fail else 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1] if len(sys.argv) > 1 else None))
