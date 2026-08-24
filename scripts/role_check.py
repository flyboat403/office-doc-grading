# -*- coding: utf-8 -*-
"""文档角色判定助手（P1 角色确认 + P5 评分对象防呆校验）。

核心规则（与 SKILL P1/P5 一致）：
- 教师提供的文档 = 标准文档（基准），永远不是被评分的学生作业。
- 学生作业只在 P5 由教师明确提供（单个文件或目录）。
- 角色判定**不靠文件名猜测**，由教师明确指定；本脚本只提供"建议提示"辅助，
  以及 P5 的**强制校验**（同一文件自评、缺学生文件 这类错误直接拦截）。

用法：
    python role_check.py suggest <path>            # 单文档角色建议（仅供参考）
    python role_check.py check p5 <std_path> <student_path...>   # P5 强制校验，出错返回非 0
    python role_check.py list <dir>                # 列出目录所有 office 文档
"""
from __future__ import annotations

import sys
from pathlib import Path

OFFICE_EXT = {".docx", ".xlsx", ".pptx"}

# 文件名启发式（仅作建议，不作判定依据）
STD_HINTS = ("标准", "答案", "结果", "参考答案", "模型", "样例", "sample", "answer",
             "key", "solution", "_std", "教师", "教师版")
STU_HINTS = ("学生", "作业", "提交", "答题", "作答", "student", "homework", "submit",
             "提交版", "学号", "姓名")


def _ext(p: Path) -> str:
    return p.suffix.lower()


def is_office(p: Path) -> bool:
    return _ext(p) in OFFICE_EXT


def _suggest_name(p: Path):
    name = p.stem.lower()
    if any(h in name for h in STU_HINTS):
        return "student(疑似学生作业)"
    if any(h in name for h in STD_HINTS):
        return "std(疑似标准文档)"
    return "unknown(文件名无法判别)"


def _suggest_content(p: Path):
    """内容启发式：标准文档通常含批注/修订，或多节多页；仅提示，不判死。"""
    try:
        sys.path.insert(0, str(Path(__file__).resolve().parent))
        from parsers import parse_file
        d = parse_file(str(p))
        notes = []
        if d.comments:
            notes.append("含%d条批注" % len(d.comments))
        if d.track_changes:
            notes.append("含修订痕迹")
        if len(d.paragraphs or []) and len(d.sections or []) > 1:
            notes.append("多节")
        return "；".join(notes) if notes else "无特殊标记"
    except Exception:
        return "(无法解析)"


def suggest_role(path):
    p = Path(path)
    if not p.exists():
        return {"ok": False, "error": "文件不存在: %s" % p}
    if not is_office(p):
        return {"ok": False, "error": "非 office 文档（支持 .docx/.xlsx/.pptx）: %s" % p}
    name_hint = _suggest_name(p)
    content = _suggest_content(p)
    return {
        "ok": True, "path": str(p),
        "name_hint": name_hint, "content_markers": content,
        "conclusion": ("文件名与内容均为标准文档特征" if name_hint.startswith("std")
                       else ("文件名与学生作业特征" if name_hint.startswith("student")
                             else "无法从文件名判别，需教师指定角色")),
        "advisory": "角色以教师指定为准，本结果仅供参考",
    }


def list_dir(directory):
    p = Path(directory)
    if not p.is_dir():
        return {"ok": False, "error": "目录不存在: %s" % directory}
    files = sorted(f for f in p.iterdir() if f.is_file() and is_office(f))
    return {"ok": True, "files": [str(f) for f in files]}


def check_p5(std_path, student_paths):
    """P5 强制校验评分对象与标准文档的关系。

    拦截两类错误：
    - 学生文件与标准文档是同一文件（拿标准评标准）。
    - 未提供任何学生文件（P5 无对象却试图评分）。
    """
    std = Path(std_path).resolve()
    if not std.exists():
        return {"ok": False, "error": "标准文档不存在: %s" % std,
                "student_files": []}
    if not is_office(std):
        return {"ok": False, "error": "标准文档非 office 类型: %s" % std,
                "student_files": []}

    wholes = [Path(sp).resolve() for sp in student_paths]
    stu_files = []
    missing = []
    conflicts = []
    for w in wholes:
        if not w.exists():
            missing.append(str(w))
            continue
        if w.is_dir():
            subs = sorted(f for f in w.iterdir()
                          if f.is_file() and is_office(f) and f.resolve() != std)
            stu_files.extend(subs)
            for f in subs:
                if f.resolve() == std:
                    conflicts.append(str(f))
        else:
            if is_office(w):
                if w == std:
                    conflicts.append(str(w))
                else:
                    stu_files.append(w)
            else:
                missing.append(str(w))

    errors = []
    if not stu_files:
        errors.append("未提供任何学生作业文件（P5 无评分对象）")
    if conflicts:
        errors.append("标准文档被当作学生作业评分（自我评分）: %s" % conflicts)

    return {
        "ok": not errors and not missing,
        "errors": errors,
        "missing": missing,
        "student_files": [str(f) for f in stu_files],
        "note": "角色判定以教师指定为准；此处仅拦截'标准当学生'与'无学生文件'两类硬错误",
    }


def main(argv):
    if len(argv) < 3:
        print(__doc__)
        return 2
    cmd = argv[1]
    import json
    if cmd == "suggest" and len(argv) >= 3:
        print(json.dumps(suggest_role(argv[2]), ensure_ascii=False, indent=1))
        return 0
    if cmd == "list" and len(argv) >= 3:
        r = list_dir(argv[2])
        print(json.dumps(r, ensure_ascii=False, indent=1))
        return 0 if r.get("ok") else 1
    if cmd == "check" and argv[2] == "p5" and len(argv) >= 4:
        r = check_p5(argv[3], argv[4:])
        print(json.dumps(r, ensure_ascii=False, indent=1))
        return 0 if r.get("ok") else 1
    print(__doc__)
    return 2


if __name__ == "__main__":
    sys.exit(main(sys.argv))
