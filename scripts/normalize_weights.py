# -*- coding: utf-8 -*-
"""权重归一化（复用 intent.py normalize_weights 逻辑）。"""
from __future__ import annotations

import math


def normalize_weights(criteria, total=100):
    """把 criteria 的权重归一到 total（整数；最大余数法，余数按小数部分分配）。

    返回 (新 criteria, 偏离度, 是否偏离>50%)。非法权重条目跳过。
    防御：任一条目结果 <=0 时恢复为 1（极端偏离场景不产出负权重）。
    """
    clean = []
    for c in criteria:
        w = c.get("weight", 0)
        if isinstance(w, bool) or not isinstance(w, (int, float)) or not math.isfinite(w):
            continue
        clean.append(c)
    criteria = clean
    # 修正（2026-11-14，F3）：空输入此前会在余数分配处 ZeroDivisionError
    if not criteria:
        return [], 0.0, False
    s = sum(c.get("weight", 0) for c in criteria) or 1
    if s == total:
        # 早退分支同样执行非正权重钳制（此前 0/负权重会原样放行）
        return [dict(c, weight=1) if (not isinstance(c.get("weight"), (int, float))
                                      or isinstance(c.get("weight"), bool)
                                      or c["weight"] <= 0) else c
                for c in criteria], 0.0, False
    scale = total / s
    norm, fracs = [], []
    for c in criteria:
        w = c["weight"] * scale
        norm.append(dict(c, weight=int(w)))
        fracs.append(w - int(w))
    residual = total - sum(x["weight"] for x in norm)
    order = sorted(range(len(norm)), key=lambda i: -fracs[i])
    for i in range(residual):
        idx = order[i % len(order)]
        norm[idx] = dict(norm[idx], weight=norm[idx]["weight"] + 1)
    for i, x in enumerate(norm):
        if x["weight"] <= 0:
            norm[i] = dict(x, weight=1)
    dev = abs(s - total) / total
    return norm, round(dev, 3), dev > 0.5


if __name__ == "__main__":
    import json
    import sys

    with open(sys.argv[1], encoding="utf-8") as f:
        data = json.load(f)
    criteria = data.get("criteria", data) if isinstance(data, dict) else data
    norm, dev, warn = normalize_weights(criteria)
    print(json.dumps(norm, ensure_ascii=False, indent=1))
    print("dev=%s warn=%s" % (dev, warn))
