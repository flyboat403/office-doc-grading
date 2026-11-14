# -*- coding: utf-8 -*-
"""office-doc-grading scripts 单元回归（2026-11-14 补齐，F8）。

覆盖本次审查发现的高危边界：空量规（F3）、水印检测（F1）、批注标记层级（F2）、
比较器边界、归一化钳制、判定回调三态、报告损坏备份（F7）、页码误报（F6）、
run 级格式聚合（F4）。
"""
import io
import zipfile
from pathlib import Path

import pytest
from lxml import etree

from parsers import ParseError, parse_bytes, parse_file
from primitives import compare, extract
from validate_rubric import validate_rubric
from normalize_weights import normalize_weights
from locate import resolve_scope
from score_diffs import grade_criteria, score_diffs
from reporting import build_report
from annotate import annotate_docx
from role_check import check_p5

EVALS = Path(__file__).resolve().parents[1] / "evals" / "files"
STD_DOCX = EVALS / "标准_论文排版.docx"
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

VML_WATERMARK_P = (
    '<w:p><w:pict xmlns:v="urn:schemas-microsoft-com:vml" '
    'xmlns:o="urn:schemas-microsoft-com:office:office">'
    '<v:shape id="PowerPlusWaterMarkObject1" o:spid="_x0000_t136" '
    'type="#_x0000_t136" style="position:absolute;width:400pt;height:200pt">'
    '<v:textpath string="草稿"/></v:shape></w:pict></w:p>')


def _inject_watermark(src, dst):
    """往 document.xml 末尾注入真实 VML 水印段（检测路径：header/document 扫描）。"""
    with zipfile.ZipFile(src) as zin, \
            zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for n in zin.namelist():
            b = zin.read(n)
            if n == "word/document.xml":
                b = b.replace(b"</w:body>",
                              VML_WATERMARK_P.encode("utf-8") + b"</w:body>")
            zout.writestr(n, b)


class FakeDoc:
    """locate/extract 的最小 ParsedDoc 替身。"""

    def __init__(self, paragraphs, tables=None, images=None, comments=None,
                 formulas=0, sections=None, header_text="", footer_text="",
                 footer_page_field=False):
        self.paragraphs = paragraphs
        self.tables = tables or []
        self.images = images or []
        self.comments = comments or []
        self.formulas = formulas
        self.sections = sections or []
        self.header_text = header_text
        self.footer_text = footer_text
        self.footer_page_field = footer_page_field


# ---------------------------------------------------------------- compare

class TestCompare:
    def test_eq_scalars(self):
        assert compare("eq", 2.5, 2.5) is True
        assert compare("eq", "宋体", "宋体") is True
        assert compare("eq", "宋体", "黑体") is False

    def test_eq_set_exact(self):
        assert compare("eq", {"FF0000"}, ["FF0000"]) is True
        assert compare("eq", {"FF0000"}, {"FF0000", "000000"}) is False

    def test_eq_pair_list(self):
        assert compare("eq", [20.0, 28.0], [20.0, 28.0]) is True
        assert compare("eq", [20.0, 28.0], [28.0, 20.0]) is False  # 顺序敏感

    def test_one_of(self):
        assert compare("one_of", {"黑体"}, ["黑体", "微软雅黑"]) is True
        assert compare("one_of", "楷体", ["黑体"]) is False

    def test_geq_range_tol(self):
        assert compare("geq", 5, 3) is True
        assert compare("geq", 2, 3) is False
        assert compare("range", 2.5, [2.0, 3.0]) is True
        assert compare("range", 3.01, [2.0, 3.0]) is False
        assert compare("tol", 2.51, [2.5, 0.1]) is True   # 边界含等号
        assert compare("tol", 2.61, [2.5, 0.1]) is False

    def test_any_all(self):
        assert compare("any_all", {"a", "b"}, ["a"]) is True
        assert compare("any_all", {"a"}, ["a", "b"]) is False

    def test_none_actual_is_undecidable(self):
        assert compare("eq", None, True) is None
        assert compare("tol", None, [2.5, 0.1]) is None


# ---------------------------------------------------------- validate_rubric

def _crit(**kw):
    base = {"id": "c1", "label": "左页边距", "target": "word",
            "primitive": "margin_left", "primitive_resolved": True,
            "comparator": "tol", "expected": [2.5, 0.1], "weight": 10,
            "scope": None, "deduct": "x"}
    base.update(kw)
    return base


class TestValidateRubric:
    def test_good_rubric_zero_errors(self):
        ok, errs = validate_rubric({"criteria": [_crit()]})
        assert errs == [] and len(ok) == 1

    def test_empty_criteria_is_error(self):  # F3
        ok, errs = validate_rubric({"criteria": []})
        assert errs and ok == []

    def test_missing_criteria_is_error(self):  # F3
        ok, errs = validate_rubric({"title": "t"})
        assert errs and ok == []

    def test_bad_comparator(self):
        _, errs = validate_rubric({"criteria": [_crit(comparator="bogus")]})
        assert any("comparator" in e[1] for e in errs)

    @pytest.mark.parametrize("w", [0, -5, True, "10", float("nan")])
    def test_bad_weights(self, w):
        _, errs = validate_rubric({"criteria": [_crit(weight=w)]})
        assert any("weight" in e[1] for e in errs)

    def test_duplicate_label_scope(self):
        _, errs = validate_rubric({"criteria": [_crit(), _crit()]})
        assert any("duplicate" in e[1] for e in errs)

    def test_para_scope_needs_index(self):
        _, errs = validate_rubric({"criteria": [_crit(scope={"type": "para"})]})
        assert any("index" in e[1] for e in errs)

    def test_text_scope_needs_val(self):
        _, errs = validate_rubric({"criteria": [_crit(scope={"type": "text"})]})
        assert any("val" in e[1] for e in errs)

    def test_type_comparator_incompatible(self):
        # font_name 是 set 类型，geq 不兼容
        _, errs = validate_rubric({"criteria": [_crit(
            label="字体", primitive="font_name", comparator="geq", expected=1)]})
        assert any("不兼容" in e[1] for e in errs)


# -------------------------------------------------------- normalize_weights

class TestNormalizeWeights:
    def test_empty_no_crash(self):  # F3（此前 ZeroDivisionError）
        norm, dev, warn = normalize_weights([])
        assert norm == [] and dev == 0.0 and warn is False

    def test_single_scales_to_total(self):
        norm, _, _ = normalize_weights([_crit(weight=50)])
        assert norm[0]["weight"] == 100

    def test_largest_remainder(self):
        norm, _, _ = normalize_weights([_crit(weight=1), _crit(weight=1),
                                        _crit(weight=1)])
        assert sum(c["weight"] for c in norm) == 100  # 34+33+33

    def test_negative_clamped_in_early_return(self):
        norm, _, _ = normalize_weights([_crit(weight=-5), _crit(weight=105)])
        assert all(c["weight"] > 0 for c in norm)

    def test_zero_weight_clamped_in_early_return(self):
        norm, _, _ = normalize_weights([_crit(weight=0), _crit(weight=100)])
        assert all(c["weight"] > 0 for c in norm)

    def test_early_return_dev_reflects_drift(self):  # N4
        # [0,100] 钳成 [1,100]，和=101≠100，dev 不得仍报 0.0
        norm, dev, _ = normalize_weights([_crit(weight=0), _crit(weight=100)])
        assert dev > 0


class TestRoleCheckP5:
    def test_missing_std_includes_skipped_key(self, tmp_path):  # N5
        r = check_p5(str(tmp_path / "不存在.docx"), [str(tmp_path)])
        assert "skipped" in r and r["student_files"] == []

    def test_non_office_std_includes_skipped_key(self, tmp_path):  # N5
        f = tmp_path / "std.txt"
        f.write_text("x", encoding="utf-8")
        r = check_p5(str(f), [str(tmp_path)])
        assert "skipped" in r and "非 office" in r["error"]


# ------------------------------------------------------------------ locate

class TestLocate:
    def test_title_fallback_first_nonempty(self):
        doc = FakeDoc([{"index": 0, "text": "", "style": None},
                       {"index": 1, "text": "题目", "style": None}])
        loc = resolve_scope(doc, {"type": "title"})
        assert loc["indices"] == {1}

    def test_body_excludes_title(self):
        doc = FakeDoc([{"index": 0, "text": "题目", "style": "Title"},
                       {"index": 1, "text": "正文", "style": None}])
        loc = resolve_scope(doc, {"type": "body"})
        assert loc["indices"] == {1}

    def test_para_out_of_range_is_empty_not_none(self):
        doc = FakeDoc([{"index": 0, "text": "a", "style": "Title"}])
        loc = resolve_scope(doc, {"type": "para", "index": 5})
        assert loc is not None and loc["indices"] == set()

    def test_text_anchor(self):
        doc = FakeDoc([{"index": 0, "text": "量子芯片封装", "style": None}])
        loc = resolve_scope(doc, {"type": "text", "val": "芯片"})
        assert loc["indices"] == {0}

    def test_unknown_scope_returns_none(self):
        assert resolve_scope(FakeDoc([]), {"type": "shape"}) is None


# -------------------------------------------------- score_diffs/grade_criteria

class TestScore:
    def test_empty_diffs(self):
        r = score_diffs([])
        assert r["score"] == 0 and r["level"] == "需改进"

    def test_half_ok(self):
        r = score_diffs([{"weight": 10, "ok": True},
                         {"weight": 10, "ok": False}], 100)
        assert r["score"] == 50.0 and r["pct"] == 50.0

    def test_grade_deterministic(self):
        doc = FakeDoc([{"index": 0, "font_size": 12},
                       {"index": 1, "font_size": 14}])
        crit = {"label": "字号", "primitive": "font_size",
                "comparator": "eq", "expected": [12, 14], "weight": 10}
        res = grade_criteria(doc, [crit])
        assert res[0]["ok"] is True and res[0]["need_review"] is False

    def test_judge_normal(self):
        doc = FakeDoc([])
        crit = {"label": "x", "primitive": None, "comparator": "eq",
                "expected": True, "weight": 5}
        res = grade_criteria(doc, [crit], lambda c, d: (True, "第1段符合"))
        assert res[0]["ok"] is True and res[0]["need_review"] is False

    def test_judge_no_evidence_needs_review(self):
        doc = FakeDoc([])
        crit = {"label": "x", "primitive": None, "comparator": "eq",
                "expected": True, "weight": 5}
        res = grade_criteria(doc, [crit], lambda c, d: (False, ""))
        assert res[0]["need_review"] is True

    def test_judge_exception_needs_review(self):
        doc = FakeDoc([])
        crit = {"label": "x", "primitive": None, "comparator": "eq",
                "expected": True, "weight": 5}

        def boom(c, d):
            raise RuntimeError("boom")

        res = grade_criteria(doc, [crit], boom)
        assert res[0]["need_review"] is True and "boom" in res[0]["evidence"]


# ---------------------------------------------------- primitives: F6 页码

class TestPageNumberPrimitive:
    def test_pgnum_fmt_true(self):
        doc = FakeDoc([], sections=[{"page_number_fmt": "numberInDash"}])
        assert extract("page_number", doc) == (True, True)

    def test_footer_field_true(self):
        doc = FakeDoc([], footer_page_field=True)
        assert extract("page_number", doc) == (True, True)

    def test_digits_in_header_no_longer_counts(self):  # F6
        doc = FakeDoc([], header_text="2026年度报告")
        v, ok = extract("page_number", doc)
        assert ok is False and v is None  # 无可信证据 → 判不了


# ------------------------------------------------------------- reporting

def _results():
    return [
        {"criterion": {"label": "页边距<b>", "weight": 10, "deduct": "",
                       "scope": None},
         "ok": True, "weight": 10, "evidence": "抽取 margin_left=2.5",
         "need_review": False, "actual": 2.5, "expected": [2.5, 0.1],
         "position": ""},
        {"criterion": {"label": "标题颜色", "weight": 5, "deduct": "未用红色",
                       "scope": None},
         "ok": False, "weight": 5, "evidence": "font_color=set()",
         "need_review": False, "actual": None, "expected": ["FF0000"],
         "position": "标题段落"},
    ]


class TestReporting:
    _scoring = {"score": 90.0, "max_score": 100.0, "pct": 90.0, "level": "优秀"}

    def test_smoke_and_escape(self, tmp_path):
        ret = build_report("学生<一>", "张三.docx", _results(),
                           self._scoring, str(tmp_path), uncovered=["整体美观"])
        html = Path(ret["report_html"]).read_text(encoding="utf-8")
        assert "页边距<b>" not in html and "页边距&lt;b&gt;" in html
        assert (tmp_path / "summary.xlsx").exists()

    def test_corrupt_workbook_backed_up(self, tmp_path, capsys):  # F7
        (tmp_path / "summary.xlsx").write_bytes(b"not an xlsx")
        build_report("李四", "李四.docx", _results(), self._scoring, str(tmp_path))
        assert (tmp_path / "summary.xlsx.corrupt.bak").exists()
        assert "警告" in capsys.readouterr().out


# ---------------------------------------------------------------- parsers

class TestParsers:
    def test_parse_eval_doc(self):
        d = parse_file(str(STD_DOCX))
        assert d.file_type == "word" and d.sections
        assert d.watermark is False  # 该文档无水印

    def test_watermark_detection(self, tmp_path):  # F1
        dst = tmp_path / "wm.docx"
        _inject_watermark(STD_DOCX, dst)
        assert parse_file(str(dst)).watermark is True

    def test_bad_zip(self, tmp_path):
        p = tmp_path / "bad.docx"
        p.write_bytes(b"definitely not a zip")
        with pytest.raises(ParseError):
            parse_file(str(p))

    def test_unsupported_ext(self, tmp_path):
        p = tmp_path / "x.pdf"
        p.write_bytes(b"junk")
        with pytest.raises(ParseError):
            parse_file(str(p))

    def test_zip_entry_limit(self, tmp_path):
        p = tmp_path / "many.docx"
        with zipfile.ZipFile(p, "w") as z:
            for i in range(201):
                z.writestr("e%03d.txt" % i, "x")
        with pytest.raises(ParseError):
            parse_file(str(p))

    def test_run_level_bold_aggregation(self, tmp_path):  # F4
        from docx import Document
        p = tmp_path / "runs.docx"
        doc = Document()
        para = doc.add_paragraph()
        para.add_run("plain ")            # 首个 run 不加粗
        para.add_run("bold").bold = True  # 次个 run 加粗
        doc.save(str(p))
        d = parse_file(str(p))
        assert d.paragraphs[0]["bold"] is True  # 旧实现按首 run → False

    def test_style_level_bold_fallback(self, tmp_path):  # F4
        from docx import Document
        p = tmp_path / "style.docx"
        doc = Document()
        doc.add_paragraph("标题文本", style="Heading 1")  # 无 run 级加粗
        if not doc.styles["Heading 1"].font.bold:
            pytest.skip("默认模板 Heading 1 无样式级加粗，无法验证回退")
        doc.save(str(p))
        assert parse_file(str(p)).paragraphs[0]["bold"] is True


# --------------------------------------------------------------- annotate

def _diffs(n=2):
    return [{"label": "标题字号%d" % i, "position": "标题段落",
             "expected": [18], "actual": 15, "weight": 5,
             "ok": False, "evidence": "字号15"} for i in range(n)]


class TestAnnotate:
    def test_markers_inside_paragraph(self, tmp_path):  # F2
        ann = annotate_docx(STD_DOCX.read_bytes(), _diffs(2))
        root = etree.fromstring(
            zipfile.ZipFile(io.BytesIO(ann)).read("word/document.xml"))
        body = root.find(W + "body")
        # body 直接子级不允许出现范围标记（旧实现违规处）
        assert body.findall(W + "commentRangeStart") == []
        assert body.findall(W + "commentRangeEnd") == []
        croot = etree.fromstring(
            zipfile.ZipFile(io.BytesIO(ann)).read("word/comments.xml"))
        assert len(croot.findall(W + "comment")) == 2

    def test_no_fails_passthrough(self, tmp_path):
        src = STD_DOCX.read_bytes()
        ann = annotate_docx(src, [])
        root = etree.fromstring(
            zipfile.ZipFile(io.BytesIO(ann)).read("word/document.xml"))
        assert root.find(W + "body").findall(".//" + W + "commentRangeStart") == []

    def test_non_docx_rejected(self):
        with pytest.raises(ValueError):
            annotate_docx(b"not a zip", _diffs(1))
