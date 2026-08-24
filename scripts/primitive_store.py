# -*- coding: utf-8 -*-
"""动态原语存储（§4.5）：seed 表（代码注册表）只读，用户级增长存储可读写。

存储位置：$OFFICE_DOC_GRADING_DIR 或 ~/.office-doc-grading/primitives.json。
条目形态（LLM 模式原语）：
    {id, file_type, aliases: [], hint, comparator, scopes: [],
     prompt_template, status: candidate|active, usage: n, created_at}

去重：机械层按 id/alias 去重；语义层由调用方（agent）在 add 前用 LLM 判断
是否已有等价原语，等价则只更新 aliases/usage（见 SKILL.md §4.5）。
"""
from __future__ import annotations

import json
import os
import time
from pathlib import Path

DEFAULT_DIR = Path.home() / ".office-doc-grading"
DEFAULT_FILE = DEFAULT_DIR / "primitives.json"


def store_path():
    env = os.environ.get("OFFICE_DOC_GRADING_DIR")
    return (Path(env) / "primitives.json") if env else DEFAULT_FILE


def load(create=True):
    p = store_path()
    if not p.exists():
        if create:
            p.parent.mkdir(parents=True, exist_ok=True)
            p.write_text("[]", encoding="utf-8")
        return []
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
        return data if isinstance(data, list) else []
    except (json.JSONDecodeError, OSError):
        return []


def save(entries):
    p = store_path()
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(json.dumps(entries, ensure_ascii=False, indent=1), encoding="utf-8")


def find(entries, name):
    for e in entries:
        if e.get("id") == name or name in e.get("aliases", []):
            return e
    return None


def add(entries, entry, bump_existing=True):
    """机械去重后新增；返回 (entries, added: bool, matched: str|None)。"""
    name = entry.get("id")
    if not name:
        return entries, False, None
    existing = find(entries, name)
    if existing is not None:
        if bump_existing:
            existing["usage"] = existing.get("usage", 0) + 1
            for a in entry.get("aliases", []):
                if a not in existing.setdefault("aliases", []):
                    existing["aliases"].append(a)
        return entries, False, existing["id"]
    entry.setdefault("status", "candidate")
    entry.setdefault("usage", 0)
    entry.setdefault("created_at", time.strftime("%Y-%m-%dT%H:%M:%S"))
    entries.append(entry)
    return entries, True, None


def list_catalog(entries, file_type=None):
    """合并 seed 目录之外的用户原语（seed 在 primitives.PRIMITIVES）。"""
    out = {}
    for e in entries:
        if e.get("status") != "active":
            continue
        if file_type is None or e.get("file_type") == file_type:
            out[e["id"]] = {
                "hint": e.get("hint", ""),
                "comparator": e.get("comparator"),
                "scopes": e.get("scopes", []),
            }
    return out


def dedup(entries):
    """按 (file_type, id) 合并，alias 归并，usage 累加。返回去重后的列表。"""
    merged = {}
    for e in entries:
        key = (e.get("file_type"), e.get("id"))
        if key in merged:
            t = merged[key]
            t["usage"] = t.get("usage", 0) + e.get("usage", 0)
            for a in e.get("aliases", []):
                if a not in t.setdefault("aliases", []):
                    t["aliases"].append(a)
            t["status"] = "active" if t.get("status") == "active" or e.get("status") == "active" else "candidate"
        else:
            merged[key] = dict(e)
    return list(merged.values())


def stats(entries):
    total = len(entries)
    active = sum(1 for e in entries if e.get("status") == "active")
    by_type = {}
    for e in entries:
        ft = e.get("file_type") or "?"
        by_type[ft] = by_type.get(ft, 0) + 1
    return {"total": total, "active": active, "by_file_type": by_type}


if __name__ == "__main__":
    import sys

    cmd = sys.argv[1] if len(sys.argv) > 1 else "list"
    entries = load()
    if cmd == "list":
        for e in entries:
            print("%s [%s] ft=%s usage=%s aliases=%s" % (
                e.get("id"), e.get("status"), e.get("file_type"),
                e.get("usage"), e.get("aliases")))
    elif cmd == "stats":
        print(json.dumps(stats(entries), ensure_ascii=False))
    elif cmd == "dedup":
        entries = dedup(entries)
        save(entries)
        print("dedup done, total=%d" % len(entries))
    else:
        print("usage: python primitive_store.py [list|stats|dedup]")
