# -*- coding: utf-8 -*-
"""原语变化摘要（P8 原语反馈用）。

评分与报告完成后、在落盘任何原语变更之前，用本工具生成"本次会话原语变化摘要"，
供 agent 展示给教师并获得确认（见 SKILL.md §P8）。

输入：可选一个描述本次会话拟变更的 JSON（缺省无 session-changes，仅打印库总览）。
    {
      "added":    [ {id, file_type, hint, comparator, scopes, prompt_template} ],
      "upgraded": [ id, ... ],          # candidate -> active 的 id
      "aliased":  [ {id, aliases_added: [...]} ]
    }

输出：结构化中文摘要（可审计），含：
- 拟新增（candidate）
- 拟升级（candidate -> active）
- 仅追加 aliases/usage
- 用户原语库总览（total / active / by_file_type）
"""
from __future__ import annotations

import json
import sys
import io
from pathlib import Path

try:
    from . import primitive_store
except ImportError:  # 直接以脚本方式运行时（python primitive_summary.py）
    import primitive_store  # type: ignore


def _fmt_scopes(scopes):
    if not scopes:
        return "-"
    return ", ".join(scopes)


def write_summary(path_or_none, out=None):
    """path_or_none：本次会话变更 JSON 路径；None 表示无变更仅看库总览。"""
    out = out or io.StringIO()
    entries = primitive_store.load()

    adds, upgrades, aliased = [], [], []
    if path_or_none:
        with open(path_or_none, encoding="utf-8-sig") as f:
            data = json.load(f)
        adds = data.get("added", [])
        upgrades = data.get("upgraded", [])
        aliased = data.get("aliased", [])

    def w(s=""):
        out.write(s + "\n")

    w("=" * 52)
    w("原语变化摘要（P8 主动反馈）")
    w("=" * 52)

    if not (adds or upgrades or aliased):
        w("本次无原语变化。")
    else:
        if adds:
            w("\n【拟新增 - candidate】")
            for a in adds:
                w("  id=%s  type=%s" % (a.get("id"), a.get("file_type")))
                w("    hint: %s" % (a.get("hint") or "-"))
                w("    comparator=%s  scopes=[%s]" % (
                    a.get("comparator"), _fmt_scopes(a.get("scopes"))))
                pt = a.get("prompt_template")
                if pt:
                    w("    prompt: %s" % pt)
        if upgrades:
            w("\n【拟升级 - candidate -> active】")
            w("  " + ", ".join(upgrades))
        if aliased:
            w("\n【仅追加 aliases/usage（等价既有原语）】")
            for x in aliased:
                w("  id=%s +aliases=%s" % (x.get("id"), x.get("aliases_added")))

    st = primitive_store.stats(entries)
    w("\n" + "-" * 52)
    w("用户原语库总览：total=%d  active=%d  by_type=%s" % (
        st["total"], st["active"], st["by_file_type"]))
    w("-" * 52)

    return out.getvalue()


if __name__ == "__main__":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    arg = sys.argv[1] if len(sys.argv) > 1 else None
    print(write_summary(arg))
