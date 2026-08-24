# -*- coding: utf-8 -*-
"""scope 定位器（§4.6 分层表）。

scope 是定位器，回答"从文档哪里取特征"。resolve_scope 返回定位结果 dict：
    {"kind": "paragraphs"|"tables"|"images"|"cells"|"all", "indices": set}
或 None（无法确定性定位 → 调用方走 LLM 回落路径，绝不伪造）。

分层：
- 已有：title / body / text / para（段落索引）
- 新增：table / cell / header / footer / image / comment / formula（元素索引）
- 受限：page（仅 first/last，用分节信息近似；文档需写明精度限制）
- 后置：shape 不支持
"""
from __future__ import annotations

import re

_TITLE_STYLES = frozenset({
    "Title", "标题", "Heading",
    *(f"Heading {n}" for n in range(1, 10)),
    *(f"标题 {n}" for n in range(1, 10)),
})


def _norm(t):
    return re.sub(r"\s+", "", t or "").lower()


def _title_indices(doc):
    idx = {i for i, p in enumerate(doc.paragraphs)
           if p.get("style") in _TITLE_STYLES}
    if not idx:
        for i, p in enumerate(doc.paragraphs):
            if p.get("text", "").strip():
                idx.add(i)
                break
    return idx


def resolve_scope(doc, scope):
    """scope: dict {type, ...} 或 None。返回定位 dict 或 None。"""
    if not scope or not isinstance(scope, dict):
        return None
    st = scope.get("type")
    paras = doc.paragraphs

    if st in ("title", "body", "text", "para"):
        title_idx = _title_indices(doc)
        if st == "title":
            return {"kind": "paragraphs", "indices": title_idx}
        if st == "body":
            return {"kind": "paragraphs",
                    "indices": {i for i in range(len(paras)) if i not in title_idx}}
        if st == "text":
            anchor = _norm(scope.get("val") or "")
            return {"kind": "paragraphs",
                    "indices": {i for i, p in enumerate(paras)
                                if anchor and anchor in _norm(p.get("text", ""))}}
        if st == "para":
            try:
                n = int(scope.get("index"))
            except (TypeError, ValueError):
                return None
            body = sorted({i for i in range(len(paras)) if i not in title_idx})
            return {"kind": "paragraphs",
                    "indices": {body[n - 1]} if 1 <= n <= len(body) else set()}

    if st == "table":
        try:
            n = int(scope.get("index"))
        except (TypeError, ValueError):
            return None
        if 1 <= n <= len(doc.tables):
            return {"kind": "tables", "indices": {n - 1}}
        return {"kind": "tables", "indices": set()}

    if st == "cell":
        try:
            t = int(scope.get("table")) - 1
            r = int(scope.get("row")) - 1
            c = int(scope.get("col")) - 1
        except (TypeError, ValueError):
            return None
        if 0 <= t < len(doc.tables):
            tbl = doc.tables[t]
            if 0 <= r < tbl["rows"] and 0 <= c < tbl["cols"]:
                return {"kind": "cells", "indices": {(t, r, c)}}
        return {"kind": "cells", "indices": set()}

    if st == "image":
        try:
            n = int(scope.get("index"))
        except (TypeError, ValueError):
            return None
        return {"kind": "images",
                "indices": {n - 1} if 1 <= n <= len(doc.images) else set()}

    if st == "comment":
        try:
            n = int(scope.get("index"))
        except (TypeError, ValueError):
            return None
        return {"kind": "comments",
                "indices": {n - 1} if 1 <= n <= len(doc.comments) else set()}

    if st == "formula":
        try:
            n = int(scope.get("index"))
        except (TypeError, ValueError):
            return None
        return {"kind": "formulas",
                "indices": {n - 1} if 1 <= n <= max(doc.formulas, 1) else set()}

    if st == "header":
        return {"kind": "header", "indices": set()}
    if st == "footer":
        return {"kind": "footer", "indices": set()}

    if st == "page":
        which = scope.get("which")
        if which == "first":
            return {"kind": "page_first", "indices": set()}
        if which == "last":
            return {"kind": "page_last", "indices": set()}
        return None  # 精确页码不做（渲染依赖），交 LLM 回落

    return None


def applies(doc, loc, predicate):
    """在 loc 定位范围内筛选元素：predicate(el) 为 True 的元素数 / 总元素数。

    返回 (命中数, 总数)。loc 为 None 时作用域不限（总数=全文档）。
    """
    if loc is None:
        els = doc.paragraphs
        return sum(1 for el in els if predicate(el)), len(els)
    kind = loc["kind"]
    idx = loc["indices"]
    if kind == "paragraphs":
        els = [p for p in doc.paragraphs if p["index"] in idx]
        return sum(1 for el in els if predicate(el)), len(els)
    if kind == "tables":
        els = [t for t in doc.tables if t["index"] in idx]
        return sum(1 for el in els if predicate(el)), len(els)
    if kind == "images":
        els = [im for im in doc.images if im["index"] in idx]
        return sum(1 for el in els if predicate(el)), len(els)
    if kind == "comments":
        return (len(idx) if doc.comments else 0, len(doc.comments))
    if kind == "cells":
        total = len(idx)
        return (total if idx else 0, total)
    if kind in ("header", "footer"):
        return (0, 1)  # 由原语直接读 header_text/footer_text，这里仅占位
    if kind == "page_first":
        return (1, 1)  # 受限 page：原语自行处理，这里占位
    if kind == "page_last":
        return (1, 1)
    return (0, 0)
