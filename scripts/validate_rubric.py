# -*- coding: utf-8 -*-
"""量规校验（复用 intent_schema 的校验规则）。

校验项：comparator 白名单、expected 形态、weight 有限正数、scope.type 白名单、
重复 (label, scope)、类型×比较器兼容。返回 (合法列表, 错误列表)。
"""
from __future__ import annotations

import json
import math

try:
    from .primitives import COMPARATORS, resolve
except ImportError:
    from primitives import COMPARATORS, resolve  # type: ignore

ALLOWED_COMPARATORS = set(COMPARATORS)
ALLOWED_SCOPE_TYPES = {"title", "body", "text", "para", "table", "cell",
                       "header", "footer", "image", "comment", "formula", "page"}
LIST_CMP = {"one_of", "any_all"}
PAIR_CMP = {"range", "tol"}

# 原语类型 × 比较器兼容（与 GradingServer 词汇表一致）
COMPARATOR_BY_TYPE = {
    "set": {"eq", "one_of", "any_all"},
    "bool": {"eq"},
    "number": {"eq", "one_of", "geq", "range", "tol"},
    "pair": {"eq"},
    "object": {"eq"},
    "string": {"eq", "one_of"},
}


def _to_number(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def _valid_expected(cmp_, expected):
    if cmp_ in LIST_CMP:
        return isinstance(expected, (list, tuple, set)) and len(expected) > 0
    if cmp_ in PAIR_CMP:
        return (isinstance(expected, (list, tuple)) and len(expected) == 2
                and all(_to_number(x) is not None for x in expected))
    if cmp_ == "geq":
        return _to_number(expected) is not None
    if isinstance(expected, (list, tuple, set)) and len(expected) == 0:
        return False
    return expected is not None


def _primitive_type(criterion):
    name = criterion.get("primitive")
    if not name:
        return None
    item = resolve(name)
    if item is None:
        return None
    return item[1].get("type")


def validate_rubric(rubric):
    """rubric: {criteria: [...], total, ...}。返回 (criteria, errors)。"""
    criteria = rubric.get("criteria", []) if isinstance(rubric, dict) else []
    out, errors = [], []
    seen = set()
    for i, c in enumerate(criteria):
        if not isinstance(c, dict):
            errors.append((i, "not dict"))
            continue
        bad = False
        label = c.get("label")
        if not isinstance(label, str) or not label.strip():
            errors.append((i, "missing/empty label"))
            bad = True
        else:
            key = (label, _scope_key(c.get("scope")))
            if key in seen:
                errors.append((i, "duplicate (label, scope): %s" % label))
                bad = True
            else:
                seen.add(key)

        cmp_ = c.get("comparator")
        if not isinstance(cmp_, str) or cmp_ not in ALLOWED_COMPARATORS:
            errors.append((i, "bad comparator: %r" % (cmp_,)))
            bad = True
        # 类型×比较器兼容：primitive 已知且有类型时校验
        t = _primitive_type(c)
        if (t and t in COMPARATOR_BY_TYPE and isinstance(cmp_, str)
                and cmp_ in ALLOWED_COMPARATORS
                and cmp_ not in COMPARATOR_BY_TYPE[t]):
            errors.append((i, "comparator %s 与原语类型 %s 不兼容" % (cmp_, t)))
            bad = True

        if "expected" not in c:
            errors.append((i, "missing expected"))
            bad = True
        elif isinstance(cmp_, str) and not _valid_expected(cmp_, c.get("expected")):
            errors.append((i, "bad expected for %s: %r" % (cmp_, c.get("expected"))))
            bad = True

        w = c.get("weight", 0)
        if isinstance(w, bool) or not isinstance(w, (int, float)) \
                or not math.isfinite(w) or w <= 0:
            errors.append((i, "bad weight: %r" % w))
            bad = True

        scope = c.get("scope")
        if scope is not None:
            if not isinstance(scope, dict):
                errors.append((i, "bad scope: %r" % (scope,)))
                bad = True
            else:
                st = scope.get("type")
                if st not in ALLOWED_SCOPE_TYPES:
                    errors.append((i, "bad scope: unknown type %r" % (st,)))
                    bad = True
                elif st == "para":
                    try:
                        int(scope.get("index"))
                    except (TypeError, ValueError):
                        errors.append((i, "bad scope: para needs int index"))
                        bad = True
                elif st == "text":
                    val = scope.get("val")
                    if not (isinstance(val, str) and val.strip()):
                        errors.append((i, "bad scope: text needs non-empty val"))
                        bad = True

        if not bad:
            out.append(c)
    return out, errors


def _scope_key(scope):
    if scope is None:
        return None
    try:
        return json.dumps(scope, sort_keys=True)
    except (TypeError, ValueError):
        return repr(scope)


if __name__ == "__main__":
    import sys

    with open(sys.argv[1], encoding="utf-8") as f:
        data = json.load(f)
    ok, errs = validate_rubric(data if isinstance(data, dict) else {"criteria": data})
    print("valid: %d, errors: %s" % (len(ok), errs))
