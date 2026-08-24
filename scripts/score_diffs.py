# -*- coding: utf-8 -*-
"""计分（复用 scoring.py 逻辑：百分制/等级 90/75/60）+ 逐条判定封装。"""
from __future__ import annotations

import json

LEVEL_EXCELLENT = 90
LEVEL_GOOD = 75
LEVEL_PASS = 60
_LEVELS = ((LEVEL_EXCELLENT, "优秀"), (LEVEL_GOOD, "良好"),
           (LEVEL_PASS, "合格"), (0, "需改进"))


def score_diffs(diffs, total=None):
    """diffs: [{weight, ok}] → {score, max_score, pct, level}。"""
    max_score = sum(d["weight"] for d in diffs) or 1
    score = sum(d["weight"] for d in diffs if d["ok"])
    scale = (total / max_score) if (total and max_score > 0) else 1.0
    score = round(score * scale, 1)
    max_score = round(max_score * scale, 1)
    pct = (score / max_score * 100) if max_score else 0.0
    pct = round(pct, 1)
    level = "需改进"
    for threshold, name in _LEVELS:
        if pct >= threshold:
            level = name
            break
    return {"score": score, "max_score": max_score, "pct": pct, "level": level}


def grade_criteria(parsed, criteria, judge_llm=None):
    """对一份 ParsedDoc 按量规逐条判定。

    - primitive_resolved 且抽取成功：确定性判定；
    - 否则调用 judge_llm(criterion, parsed) 返回 (ok, evidence)；
      judge_llm 为 None、或回调判不出（返回 None/空依据）/回调抛错 时，
      该条标记 need_review（而非硬给 ok=False）——"判不出"不等于"不通过"。
    返回 [criterion, ok, weight, evidence, resolved, need_review, actual]。
    """
    try:
        from .primitives import compare, extract
    except ImportError:
        from primitives import compare, extract  # type: ignore

    results = []
    for c in criteria:
        name = c.get("primitive")
        v, ok_ext = extract(name, parsed, c.get("scope")) if name else (None, False)
        if ok_ext:
            verdict = compare(c.get("comparator"), v, c.get("expected"))
            if verdict is not None:
                results.append({
                    "criterion": c, "ok": bool(verdict), "weight": c["weight"],
                    "evidence": "抽取 %s=%r，比较器 %s，期望 %r" % (
                        name, v, c.get("comparator"), c.get("expected")),
                    "resolved": True, "need_review": False, "actual": v,
                })
                continue
        if judge_llm is not None:
            try:
                judged = judge_llm(c, parsed)
                ok, evidence = judged
            except Exception as e:
                results.append({
                    "criterion": c, "ok": False, "weight": c["weight"],
                    "evidence": "LLM 判定回调异常（%s），待复核" % e,
                    "resolved": False, "need_review": True, "actual": None,
                })
                continue
            if judged is None or not evidence:
                # 回调判不出（返回 None 或无定位依据）→ 待复核，不硬判通过/不通过
                results.append({
                    "criterion": c, "ok": False, "weight": c["weight"],
                    "evidence": (evidence or "LLM 判定未给出可信依据"),
                    "resolved": False, "need_review": True, "actual": None,
                })
                continue
            results.append({
                "criterion": c, "ok": bool(ok), "weight": c["weight"],
                "evidence": evidence or "LLM 判定", "resolved": False,
                "need_review": False, "actual": None,
            })
        else:
            results.append({
                "criterion": c, "ok": False, "weight": c["weight"],
                "evidence": "无法确定性抽取且无 LLM 判定器", "resolved": False,
                "need_review": True, "actual": None,
            })
    return results


def summarize(results, total=100):
    diffs = [{"weight": r["weight"], "ok": r["ok"]} for r in results]
    return score_diffs(diffs, total)


if __name__ == "__main__":
    import sys

    # 用法：python score_diffs.py <diffs.json> [total]
    with open(sys.argv[1], encoding="utf-8") as f:
        diffs = json.load(f)
    print(json.dumps(score_diffs(diffs, float(sys.argv[2]) if len(sys.argv) > 2 else None),
                     ensure_ascii=False))
