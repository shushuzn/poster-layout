#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""项目入口：``python build.py``

等价于 ``python -m poster``。
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "src"))

from poster.cli import main  # noqa: E402

if __name__ == "__main__":
    raise SystemExit(main())
