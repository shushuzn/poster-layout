# -*- coding: utf-8 -*-
"""可复现构建：把 pptx 内的 zip 时间戳归一化。

python-pptx 保存时给每个 zip 条目写当前时间，导致**同样的输入产出不同的字节**。
对一个 16MB 的二进制交付物来说，这意味着每跑一次构建就在 git 里多留一份 16MB
的历史垃圾 —— 也是本仓库 `.git` 曾经膨胀到 265MB 的主因。

做法：按原顺序重写 zip，把所有条目的 ``date_time`` 固定为 1980-01-01
（zip 格式的纪元，也是可复现构建社区的通行取值），
其余属性（压缩方式、外部属性、内容）原样保留。

保留 `SOURCE_DATE_EPOCH` 约定的接口：若设置了该环境变量，则用它换算日期。
"""
from __future__ import annotations

import os
import time
import zipfile
from pathlib import Path

# zip 能表示的最早时间；可复现构建社区的通行常量
ZIP_EPOCH = (1980, 1, 1, 0, 0, 0)


def _fixed_date_time() -> tuple[int, int, int, int, int, int]:
    """优先读 SOURCE_DATE_EPOCH，否则用 zip 纪元。"""
    env = os.environ.get("SOURCE_DATE_EPOCH")
    if not env:
        return ZIP_EPOCH
    try:
        t = time.gmtime(int(env))
    except (ValueError, OverflowError, OSError):
        return ZIP_EPOCH
    # zip 不支持 1980 年之前
    if t.tm_year < 1980:
        return ZIP_EPOCH
    return (t.tm_year, t.tm_mon, t.tm_mday, t.tm_hour, t.tm_min, t.tm_sec)


def freeze_zip(path: str | Path, date_time=None) -> None:
    """就地重写 zip，使内容相同则字节相同。"""
    path = Path(path)
    dt = date_time or _fixed_date_time()
    tmp = path.with_name(path.name + ".tmp")

    with zipfile.ZipFile(path) as zin:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                data = zin.read(info.filename)
                ni = zipfile.ZipInfo(info.filename, date_time=dt)
                ni.compress_type = info.compress_type
                ni.create_system = info.create_system
                ni.external_attr = info.external_attr
                ni.internal_attr = info.internal_attr
                zout.writestr(ni, data)

    os.replace(tmp, path)


def is_reproducible(path: str | Path) -> bool:
    """自检：再规范化一次，字节数不变即说明已冻结。"""
    path = Path(path)
    before = path.read_bytes()
    freeze_zip(path)
    return path.read_bytes() == before
