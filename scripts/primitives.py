# -*- coding: utf-8 -*-
"""确定性抽取原语 + 比较器（§4.1-4.3）。

原语是"从 ParsedDoc 抽取一个可比较特征"的最小能力，粒度比维度更细。
量规生成时，LLM 为每个 criterion 选一个原语；命中（在 PRIMITIVES 中）就走
确定性抽取 + comparator 判定，miss 就回落 LLM 判定。

类型约定（与 GradingServer 词汇表一致）：
    bool   → eq
    number → eq/one_of/geq/range/tol
    set    → eq/one_of/any_all（返回 Python set，eq 精确比较，one_of 任一命中）
    string → eq/one_of
    pair   → eq（[a, b] 精确比较）
    object → eq

数值型段落原语返回作用域内的最大值（hint 中注明），保证 geq/range/tol 可用。
"""
from __future__ import annotations

try:
    from . import locate
except ImportError:  # 直接以脚本运行时（python primitives.py）
    import locate  # type: ignore

# --------------------------------------------------------------------------
# 比较器
# --------------------------------------------------------------------------

COMPARATORS = ("eq", "one_of", "geq", "range", "tol", "any_all")


def _norm_set(v):
    if isinstance(v, (list, tuple, set)):
        return set(v)
    return {v}


def compare(cmp_, actual, expected):
    """确定性比较。actual 为 None（无法抽取）时返回 None（无法判定）。"""
    if actual is None:
        return None
    if cmp_ == "eq":
        if isinstance(actual, set):
            return actual == _norm_set(expected)
        if isinstance(actual, (list, tuple)) or isinstance(expected, (list, tuple)):
            return list(actual) == list(expected)
        return actual == expected
    if cmp_ == "one_of":
        cands = list(expected) if isinstance(expected, (list, tuple, set)) else [expected]
        if isinstance(actual, set):
            return bool(actual & set(cands))
        return actual in cands
    if cmp_ == "any_all":
        cands = list(expected) if isinstance(expected, (list, tuple, set)) else [expected]
        if isinstance(actual, set):
            return set(cands) <= actual
        return actual in cands
    try:
        a = float(actual)
    except (TypeError, ValueError):
        return None
    if cmp_ == "geq":
        try:
            return a >= float(expected)
        except (TypeError, ValueError):
            return None
    if cmp_ == "range":
        try:
            lo, hi = float(expected[0]), float(expected[1])
            return lo <= a <= hi
        except (TypeError, ValueError, IndexError):
            return None
    if cmp_ == "tol":
        try:
            val, tol = float(expected[0]), float(expected[1])
            return abs(a - val) <= tol
        except (TypeError, ValueError, IndexError):
            return None
    return None


# --------------------------------------------------------------------------
# 抽取原语
# --------------------------------------------------------------------------

def _para_val(doc, scope, key, mode="max"):
    """作用域内段落属性聚合：mode=max 取最大数值；any 任一为真；set 收集。

    scope 为 None 表示不限位置（全文）；scope 给定但定位失败返回 None（无法判定）。
    """
    loc = locate.resolve_scope(doc, scope)
    if scope is not None and loc is None:
        return None  # 有 scope 但无法确定性定位 → 无法判定
    hit, total = locate.applies(doc, loc, lambda p: True)
    if total == 0:
        return None
    vals = []
    for p in doc.paragraphs:
        if loc is None or p["index"] in loc["indices"]:
            v = p.get(key)
            if v is not None:
                vals.append(v)
    if mode == "max":
        return max(vals) if vals else None
    if mode == "any":
        return any(vals)
    if mode == "set":
        return set(vals)


# ---- word ----

def _font_name(doc, scope):
    """作用域内字体名集合（含显式与主题字体名，展平）。"""
    loc = locate.resolve_scope(doc, scope)
    if scope is not None and loc is None:
        return None
    names = set()
    for p in doc.paragraphs:
        if loc is None or p["index"] in loc["indices"]:
            v = p.get("font_name")
            if isinstance(v, (set, list, tuple)):
                names.update(v)
            elif v:
                names.add(v)
    return names or None


def _font_size(doc, scope):
    return _para_val(doc, scope, "font_size", "set")


def _bold(doc, scope):
    return _para_val(doc, scope, "bold", "any")


def _italic(doc, scope):
    return _para_val(doc, scope, "italic", "any")


def _underline(doc, scope):
    return _para_val(doc, scope, "underline", "any")


def _align(doc, scope, name):
    aligns = _para_val(doc, scope, "align", "set") or set()
    return name in aligns


def _align_left(doc, scope):
    return _align(doc, scope, "left")


def _align_center(doc, scope):
    return _align(doc, scope, "center")


def _align_right(doc, scope):
    return _align(doc, scope, "right")


def _align_justify(doc, scope):
    return _align(doc, scope, "justify")


def _indent(doc, scope):
    return _para_val(doc, scope, "first_line_indent", "max")


def _indent_all(doc, scope):
    """作用域内所有非空段落是否都有首行缩进（整体性要求，如"正文首行缩进2字符"）。

    新增（2026-08-19）：indent 取最大值只判断"至少一段达标"，对"所有段落都要缩进"
    这类需求会漏判（学生只给一段设缩进也通过）。本原语要求全部命中。
    """
    loc = locate.resolve_scope(doc, scope)
    if scope is not None and loc is None:
        return None
    paras = [p for p in doc.paragraphs
             if (loc is None or p["index"] in loc["indices"]) and p.get("text", "").strip()]
    if not paras:
        return None
    return all(p.get("first_line_indent") is not None for p in paras)


def _line_spacing(doc, scope):
    return _para_val(doc, scope, "line_spacing", "max")


def _space_before(doc, scope):
    return _para_val(doc, scope, "space_before", "max")


def _space_after(doc, scope):
    return _para_val(doc, scope, "space_after", "max")


def _sec(doc, key):
    return doc.sections[0].get(key) if doc.sections else None


def _margin_top(doc, scope):
    return _sec(doc, "margin_top_cm")


def _margin_bottom(doc, scope):
    return _sec(doc, "margin_bottom_cm")


def _margin_left(doc, scope):
    return _sec(doc, "margin_left_cm")


def _margin_right(doc, scope):
    return _sec(doc, "margin_right_cm")


def _paper_size(doc, scope):
    w, h = _sec(doc, "paper_w_cm"), _sec(doc, "paper_h_cm")
    return [w, h] if w and h else None


def _paper_orientation(doc, scope):
    return _sec(doc, "orientation")


def _gutter(doc, scope):
    return _sec(doc, "gutter_cm")


def _gutter_pos(doc, scope):
    return _sec(doc, "gutter_pos")


def _header_text(doc, scope):
    return doc.header_text or None


def _footer_text(doc, scope):
    return doc.footer_text or None


def _page_number(doc, scope):
    """有页码设置（pgNumType 或页脚 PAGE 域或页眉页脚含数字）即认为有页码。"""
    for s in doc.sections:
        if s.get("page_number_fmt"):
            return True
    if getattr(doc, "footer_page_field", False):
        return True
    blob = (doc.header_text or "") + (doc.footer_text or "")
    return bool(blob and ("PAGE" in blob or any(c.isdigit() for c in blob)))


def _page_break(doc, scope):
    return doc.page_breaks > 0


def _para_count(doc, scope):
    return len(doc.paragraphs)


def _char_count(doc, scope):
    return len(doc.text.replace("\n", ""))


def _table_count(doc, scope):
    return len(doc.tables)


def _table_dim(doc, scope):
    t = doc.tables[0] if doc.tables else None
    return [t["rows"], t["cols"]] if t else None


def _table_align(doc, scope):
    t = doc.tables[0] if doc.tables else None
    return t["align"] if t else None


def _image_count(doc, scope):
    return len(doc.images)


def _image_size(doc, scope):
    im = doc.images[0] if doc.images else None
    return [im["width_cm"], im["height_cm"]] if im else None


def _comment(doc, scope):
    return bool(doc.comments)


def _comment_count(doc, scope):
    return len(doc.comments)


def _formula(doc, scope):
    return doc.formulas > 0


def _formula_count(doc, scope):
    return doc.formulas


def _footnote(doc, scope):
    return doc.footnotes > 0


def _endnote(doc, scope):
    return doc.endnotes > 0


def _track_changes(doc, scope):
    return doc.track_changes


def _watermark(doc, scope):
    return doc.watermark


def _toc(doc, scope):
    return doc.toc


def _bullet(doc, scope):
    return any((p.get("style") or "").lower() in ("list bullet", "list", "列表")
               for p in doc.paragraphs)


def _font_color(doc, scope):
    return _para_val(doc, scope, "font_color", "set")


def _highlight(doc, scope):
    return _para_val(doc, scope, "highlight", "any")


def _para_shading(doc, scope):
    """作用域内是否有有效段落底纹（有 shd 且 fill/themeFill/themeColor 非空）。

    修正（2026-08-19）：原实现只判断 shd 元素存在，学生文件常写
    <w:shd w:val="clear" w:color="auto"/> 无 fill，视觉上无底纹却判为有。
    """
    def valid(s):
        return bool(s and (s.get("fill") or s.get("themeFill")
                           or s.get("themeColor") or s.get("themeFillTint")))
    return any(valid(p.get("shading")) for p in doc.paragraphs
               if _in_scope(doc, scope, p))


def _para_border(doc, scope):
    """作用域内是否有有效段落边框（pBdr 存在且非全 none）。

    修正（2026-08-19）：原实现只看 pBdr 元素存在，学生文件常写各边
    val="none"（如 <w:top w:val="none"/>），视觉上无边框却判为有。
    """
    def valid(b):
        if not b:
            return False
        vals = set(v for v in b.values() if v)
        return bool(vals) and not vals.issubset({"none"})
    return any(valid(p.get("border")) for p in doc.paragraphs
               if _in_scope(doc, scope, p))


def _in_scope(doc, scope, p):
    loc = locate.resolve_scope(doc, scope)
    if scope is not None and loc is None:
        return False
    return loc is None or p["index"] in loc["indices"]


def _image_wrap(doc, scope):
    return {im.get("wrap") for im in doc.images if im.get("wrap")}


def _image_position(doc, scope):
    return {im.get("position") for im in doc.images if im.get("position")}


def _image_shadow(doc, scope):
    return any(im.get("shadow") for im in doc.images)


def _page_number_fmt(doc, scope):
    return _sec(doc, "page_number_fmt")


def _page_number_start(doc, scope):
    return _sec(doc, "page_number_start")


def _header_highlight(doc, scope):
    return _sec(doc, "header_highlight")


def _caption(doc, scope):
    """题注检测：段落文本或文本框文本含“图N/表N”模式（如“图1 中国芯…”）。"""
    import re
    pat = re.compile(r"(图|表)\s*\d")
    for p in doc.paragraphs:
        if pat.search(p.get("text", "")):
            return True
    for t in getattr(doc, "textbox_texts", []) or []:
        if pat.search(t):
            return True
    return False


# --------------------------------------------------------------------------
# word 扩展：直读 OOXML 控制标签的新增确定性原语（seed）
#   解析器在 paragraphs[].widow_control / top_line_punct，
#   sections[].cols_num / cols_sep / page_border 提供原始值。
# --------------------------------------------------------------------------

def _para_flag_any(doc, scope, key):
    """作用域内任一非空段落是否开启该布尔控制项（widowControl/topLinePunct）。"""
    hits, total = [], 0
    for p in doc.paragraphs:
        if not _in_scope(doc, scope, p):
            continue
        total += 1
        v = p.get(key)
        if v is True:
            hits.append(p)
    if total == 0:
        return False
    return bool(hits)


def _widow_control(doc, scope):
    """正文段落是否设置孤行控制：任一作用域段落 pPr 含 w:widowControl（非 val=0）。"""
    return _para_flag_any(doc, scope, "widow_control")


def _top_line_punct(doc, scope):
    """正文是否允许行首标点压缩：任一作用域段落 pPr 含 w:topLinePunct。"""
    return _para_flag_any(doc, scope, "top_line_punct")


def _any_section(doc, pred):
    if not getattr(doc, "sections", None):
        return None
    vals = [pred(s) for s in doc.sections if pred(s) is not None]
    if not vals:
        return None
    return any(vals)


def _page_border(doc, scope):
    """任一节是否设置有效页面边框（pgBorders 存在且至少一边 val 非 none）。"""
    return _any_section(doc, lambda s: s.get("page_border", {}).get("valid"))


def _column_two(doc, scope):
    """任一节是否为多栏（cols_num >= 2）。目标段落归属节精确判定必要时走 LLM。"""
    return _any_section(doc, lambda s: (s.get("cols_num") or 1) >= 2)


def _column_sep(doc, scope):
    """任一多栏节是否含分隔线（cols_num >= 2 且 sep=1）。"""
    return _any_section(doc, lambda s: (s.get("cols_num") or 1) >= 2 and s.get("cols_sep") is True)


# ---- xlsx ----

def _sheet_count(doc, scope):
    return doc.sheet_count


def _xlsx_freeze(doc, scope):
    return bool(getattr(doc, "xlsx_freeze", None))


def _xlsx_formula(doc, scope):
    return getattr(doc, "xlsx_formulas", 0) > 0


def _xlsx_formula_count(doc, scope):
    return getattr(doc, "xlsx_formulas", 0)


def _xlsx_protected(doc, scope):
    return bool(getattr(doc, "xlsx_protected", False))


def _xlsx_merged(doc, scope):
    return getattr(doc, "xlsx_merged", 0) > 0


def _xlsx_table(doc, scope):
    return len(doc.tables) > 0


def _xlsx_cond_format(doc, scope):
    return getattr(doc, "xlsx_cond_format", 0) > 0


def _xlsx_filter(doc, scope):
    return bool(getattr(doc, "xlsx_auto_filter", False))


def _xlsx_font_size(doc, scope):
    return set(getattr(doc, "xlsx_fonts", []))


def _xlsx_col_width(doc, scope):
    ws = getattr(doc, "xlsx_col_widths", [])
    return max(ws) if ws else None


def _xlsx_row_height(doc, scope):
    ws = getattr(doc, "xlsx_row_heights", [])
    return max(ws) if ws else None


def _xlsx_align(doc, scope):
    return set(getattr(doc, "xlsx_aligns", []))


def _xlsx_border(doc, scope):
    return bool(getattr(doc, "xlsx_border", False))


def _xlsx_fill(doc, scope):
    return bool(getattr(doc, "xlsx_fills", set()))


def _xlsx_wrap_text(doc, scope):
    return bool(getattr(doc, "xlsx_wrap_text", False))


def _xlsx_number_format(doc, scope):
    return set(getattr(doc, "xlsx_number_formats", []))


def _xlsx_data_validation(doc, scope):
    return (getattr(doc, "xlsx_data_validations", 0) or 0) > 0


def _xlsx_print_area(doc, scope):
    return bool(getattr(doc, "xlsx_print_area", False))


def _xlsx_hyperlink(doc, scope):
    return (getattr(doc, "xlsx_hyperlinks", 0) or 0) > 0


def _xlsx_merge_count(doc, scope):
    return getattr(doc, "xlsx_merged", 0)


def _xlsx_cond_type(doc, scope):
    return set(getattr(doc, "xlsx_cond_format_types", []))


# ---- pptx ----

def _slide_count(doc, scope):
    return doc.slide_count


def _pptx_font_size(doc, scope):
    sizes = set()
    for s in getattr(doc, "slides", []):
        sizes.update(s.get("fonts", []))
    return sizes


def _pptx_align_center(doc, scope):
    return any("center" in s.get("aligns", []) for s in getattr(doc, "slides", []))


def _pptx_bullet(doc, scope):
    return any(s.get("bullet") for s in getattr(doc, "slides", []))


def _pptx_animation(doc, scope):
    return any(s.get("animation") for s in getattr(doc, "slides", []))


def _pptx_transition(doc, scope):
    return any(s.get("transition") for s in getattr(doc, "slides", []))


def _pptx_notes(doc, scope):
    return any(s.get("notes") for s in getattr(doc, "slides", []))


def _pptx_hyperlink(doc, scope):
    return any(s.get("hyperlink") for s in getattr(doc, "slides", []))


def _pptx_image_count(doc, scope):
    return sum(s.get("images", 0) for s in getattr(doc, "slides", []))


def _pptx_text_align(doc, scope):
    aligns = set()
    for s in getattr(doc, "slides", []):
        aligns.update(s.get("aligns", []))
    return aligns or None


def _pptx_slide_has_title(doc, scope):
    return any(s.get("has_title") for s in getattr(doc, "slides", []))


def _pptx_aspect_ratio(doc, scope):
    return getattr(doc, "pptx_aspect", None)


def _pptx_layout(doc, scope):
    return set(getattr(doc, "pptx_layouts", [])) or None


def _pptx_shape_count(doc, scope):
    return sum(s.get("shape_count", 0) for s in getattr(doc, "slides", []))


# --------------------------------------------------------------------------
# 注册表
# --------------------------------------------------------------------------

PRIMITIVES = {
    # word
    "font_name": (_font_name, {"type": "set", "units": None,
                               "hint": "作用域内出现的字体名集合（如 {宋体}）"}),
    "font_size": (_font_size, {"type": "set", "units": "磅",
                               "hint": "作用域内出现的字号集合（磅，如 {12, 14}）"}),
    "bold": (_bold, {"type": "bool", "units": None, "hint": "作用域内是否有加粗"}),
    "italic": (_italic, {"type": "bool", "units": None, "hint": "作用域内是否有斜体"}),
    "underline": (_underline, {"type": "bool", "units": None, "hint": "作用域内是否有下划线"}),
    "align_left": (_align_left, {"type": "bool", "units": None, "hint": "作用域内是否有左对齐"}),
    "align_center": (_align_center, {"type": "bool", "units": None, "hint": "作用域内是否有居中对齐"}),
    "align_right": (_align_right, {"type": "bool", "units": None, "hint": "作用域内是否有右对齐"}),
    "align_justify": (_align_justify, {"type": "bool", "units": None, "hint": "作用域内是否有两端对齐"}),
"indent": (_indent, {"type": "number", "units": "厘米",
                          "hint": "作用域内首行缩进最大值（厘米，如 0.74≈2字符）"}),
    "indent_all": (_indent_all, {"type": "bool", "units": None,
                                  "hint": "作用域内所有非空段落是否都有首行缩进（整体性要求）"}),
    "line_spacing": (_line_spacing, {"type": "number", "units": "磅或倍数",
                                     "hint": "作用域内行距最大值（固定值磅/倍数，如 1.5）"}),
    "space_before": (_space_before, {"type": "number", "units": "磅", "hint": "作用域内段前距最大值"}),
    "space_after": (_space_after, {"type": "number", "units": "磅", "hint": "作用域内段后距最大值"}),
    "margin_top": (_margin_top, {"type": "number", "units": "厘米", "hint": "上页边距"}),
    "margin_bottom": (_margin_bottom, {"type": "number", "units": "厘米", "hint": "下页边距"}),
    "margin_left": (_margin_left, {"type": "number", "units": "厘米", "hint": "左页边距"}),
    "margin_right": (_margin_right, {"type": "number", "units": "厘米", "hint": "右页边距"}),
    "paper_size": (_paper_size, {"type": "pair", "units": "厘米",
                                 "hint": "纸张 [宽, 高]（如 [21, 29.7]）"}),
    "paper_orientation": (_paper_orientation, {"type": "string", "units": None,
                                               "hint": "portrait/landscape"}),
    "gutter": (_gutter, {"type": "number", "units": "厘米", "hint": "装订线宽度"}),
    "gutter_pos": (_gutter_pos, {"type": "string", "units": None, "hint": "装订线位置：left/right（w:pgMar gutterPos）"}),
    "header_text": (_header_text, {"type": "string", "units": None, "hint": "首页节页眉文本"}),
    "footer_text": (_footer_text, {"type": "string", "units": None, "hint": "首页节页脚文本"}),
    "page_number": (_page_number, {"type": "bool", "units": None, "hint": "是否有页码设置"}),
    "page_break": (_page_break, {"type": "bool", "units": None, "hint": "是否有分页符"}),
    "para_count": (_para_count, {"type": "number", "units": "个", "hint": "段落总数"}),
    "char_count": (_char_count, {"type": "number", "units": "字符", "hint": "正文字符总数"}),
    "table_count": (_table_count, {"type": "number", "units": "个", "hint": "表格总数"}),
    "table_dim": (_table_dim, {"type": "pair", "units": "行x列",
                               "hint": "第一个表格 [行数, 列数]"}),
    "table_align": (_table_align, {"type": "string", "units": None,
                                   "hint": "第一个表格对齐：left/center/right"}),
    "image_count": (_image_count, {"type": "number", "units": "张", "hint": "图片总数"}),
    "image_size": (_image_size, {"type": "pair", "units": "厘米",
                                 "hint": "第一张图片 [宽, 高]"}),
    "comment": (_comment, {"type": "bool", "units": None, "hint": "是否有批注"}),
    "comment_count": (_comment_count, {"type": "number", "units": "条", "hint": "批注条数"}),
    "formula": (_formula, {"type": "bool", "units": None, "hint": "是否有公式"}),
    "formula_count": (_formula_count, {"type": "number", "units": "个", "hint": "公式个数"}),
    "footnote": (_footnote, {"type": "bool", "units": None, "hint": "是否有脚注"}),
    "endnote": (_endnote, {"type": "bool", "units": None, "hint": "是否有尾注"}),
    "track_changes": (_track_changes, {"type": "bool", "units": None, "hint": "是否有修订"}),
    "watermark": (_watermark, {"type": "bool", "units": None, "hint": "是否有水印"}),
    "toc": (_toc, {"type": "bool", "units": None, "hint": "是否有目录"}),
    "bullet": (_bullet, {"type": "bool", "units": None, "hint": "是否有项目符号/编号"}),
    "font_color": (_font_color, {"type": "set", "units": "十六进制", "hint": "作用域内字体颜色集合（如 {FF0000}）"}),
    "highlight": (_highlight, {"type": "bool", "units": None, "hint": "作用域内是否有突出显示"}),
    "para_shading": (_para_shading, {"type": "bool", "units": None, "hint": "作用域内是否有段落底纹"}),
    "para_border": (_para_border, {"type": "bool", "units": None, "hint": "作用域内是否有段落边框"}),
    "image_wrap": (_image_wrap, {"type": "set", "units": None, "hint": "图片环绕方式集合：inline/square/tight/top_and_bottom/none/through"}),
    "image_position": (_image_position, {"type": "set", "units": None, "hint": "图片位置集合（H:relativeFrom/值 V:…，如 中间居右→H:column/right? V:margin/center?）"}),
    "image_shadow": (_image_shadow, {"type": "bool", "units": None, "hint": "图片是否有阴影效果"}),
    "page_number_fmt": (_page_number_fmt, {"type": "string", "units": None, "hint": "页码格式 fmt：numberInDash≈\\\"-1-,-2-,-3-\\\" 类型"}),
    "page_number_start": (_page_number_start, {"type": "number", "units": "页", "hint": "页码起始值（w:pgNumType start）"}),
    "header_highlight": (_header_highlight, {"type": "string", "units": None, "hint": "页眉突出显示颜色（yellow 等）"}),
    "caption": (_caption, {"type": "bool", "units": None, "hint": "是否有题注（段落/文本框含“图N/表N”模式）"}),
    "widow_control": (_widow_control, {"type": "bool", "units": None, "hint": "作用域内是否设置孤行控制（段落 pPr 含 w:widowControl 非 val=0）"}),
    "top_line_punct": (_top_line_punct, {"type": "bool", "units": None, "hint": "作用域内是否允许行首标点压缩（段落 pPr 含 w:topLinePunct）"}),
    "page_border": (_page_border, {"type": "bool", "units": None, "hint": "任一节是否设有效页面边框（pgBorders 存在且至少一边 val 非 none；颜色/磅值需在 expected 或复核中核对）"}),
    "column_two": (_column_two, {"type": "bool", "units": None, "hint": "任一节是否多栏（cols_num>=2）；目标段落归属节精确判定必要时走 LLM"}),
    "column_sep": (_column_sep, {"type": "bool", "units": None, "hint": "任一多栏节是否含分隔线（cols_num>=2 且 sep=1）"}),
    # xlsx
    "sheet_count": (_sheet_count, {"type": "number", "units": "个", "hint": "工作表数量"}),
    "xlsx_freeze": (_xlsx_freeze, {"type": "bool", "units": None, "hint": "是否冻结窗格"}),
    "xlsx_formula": (_xlsx_formula, {"type": "bool", "units": None, "hint": "是否含公式"}),
    "xlsx_formula_count": (_xlsx_formula_count, {"type": "number", "units": "个", "hint": "公式数量"}),
    "xlsx_protected": (_xlsx_protected, {"type": "bool", "units": None, "hint": "工作表是否保护"}),
    "xlsx_merged": (_xlsx_merged, {"type": "bool", "units": None, "hint": "是否有合并单元格"}),
    "xlsx_table": (_xlsx_table, {"type": "bool", "units": None, "hint": "是否有表格样式"}),
    "xlsx_cond_format": (_xlsx_cond_format, {"type": "bool", "units": None, "hint": "是否有条件格式"}),
    "xlsx_filter": (_xlsx_filter, {"type": "bool", "units": None, "hint": "是否有筛选"}),
    "xlsx_font_size": (_xlsx_font_size, {"type": "set", "units": "磅", "hint": "抽样单元格字号集合"}),
    "xlsx_col_width": (_xlsx_col_width, {"type": "number", "units": "字符", "hint": "抽样列宽最大值"}),
    "xlsx_row_height": (_xlsx_row_height, {"type": "number", "units": "磅", "hint": "抽样行高最大值"}),
    "xlsx_align": (_xlsx_align, {"type": "set", "units": None, "hint": "抽样单元格水平对齐集合（left/center/right 等）"}),
    "xlsx_border": (_xlsx_border, {"type": "bool", "units": None, "hint": "是否有单元格边框"}),
    "xlsx_fill": (_xlsx_fill, {"type": "bool", "units": None, "hint": "是否有单元格底纹/填充色"}),
    "xlsx_wrap_text": (_xlsx_wrap_text, {"type": "bool", "units": None, "hint": "是否有自动换行单元格"}),
    "xlsx_number_format": (_xlsx_number_format, {"type": "set", "units": None, "hint": "抽查单元格数字格式集合（0.00%/#,##0 等）"}),
    "xlsx_data_validation": (_xlsx_data_validation, {"type": "bool", "units": None, "hint": "是否有数据验证"}),
    "xlsx_print_area": (_xlsx_print_area, {"type": "bool", "units": None, "hint": "是否设置打印区域"}),
    "xlsx_hyperlink": (_xlsx_hyperlink, {"type": "bool", "units": None, "hint": "是否有单元格超链接"}),
    "xlsx_merge_count": (_xlsx_merge_count, {"type": "number", "units": "个", "hint": "合并单元格数量"}),
    "xlsx_cond_type": (_xlsx_cond_type, {"type": "set", "units": None, "hint": "条件格式类型集合（cellIs/dataBar/colorScale/iconSet 等）"}),
    # pptx
    "slide_count": (_slide_count, {"type": "number", "units": "张", "hint": "幻灯片数量"}),
    "pptx_font_size": (_pptx_font_size, {"type": "set", "units": "磅", "hint": "各页字号集合"}),
    "pptx_align_center": (_pptx_align_center, {"type": "bool", "units": None, "hint": "是否有居中对齐文本"}),
    "pptx_bullet": (_pptx_bullet, {"type": "bool", "units": None, "hint": "是否有项目符号"}),
    "pptx_animation": (_pptx_animation, {"type": "bool", "units": None, "hint": "是否有动画"}),
    "pptx_transition": (_pptx_transition, {"type": "bool", "units": None, "hint": "是否有切换效果"}),
    "pptx_notes": (_pptx_notes, {"type": "bool", "units": None, "hint": "是否有演讲者备注"}),
    "pptx_hyperlink": (_pptx_hyperlink, {"type": "bool", "units": None, "hint": "是否有超链接"}),
    "pptx_image_count": (_pptx_image_count, {"type": "number", "units": "张", "hint": "图片总数"}),
    "pptx_text_align": (_pptx_text_align, {"type": "set", "units": None, "hint": "各页文本对齐集合（left/center/right/justify）"}),
    "pptx_slide_has_title": (_pptx_slide_has_title, {"type": "bool", "units": None, "hint": "是否有含文本的标题占位符"}),
    "pptx_aspect_ratio": (_pptx_aspect_ratio, {"type": "string", "units": None, "hint": "宽高比：16:9 / 4:3"}),
    "pptx_layout": (_pptx_layout, {"type": "set", "units": None, "hint": "使用的幻灯片版式集合（Title/Title and Content/Blank 等）"}),
    "pptx_shape_count": (_pptx_shape_count, {"type": "number", "units": "个", "hint": "形状总数"}),
}

FILE_TYPES = {
    "word": {"font_name", "font_size", "bold", "italic", "underline",
             "align_left", "align_center", "align_right", "align_justify",
             "indent", "indent_all", "line_spacing", "space_before", "space_after",
             "margin_top", "margin_bottom", "margin_left", "margin_right",
             "paper_size", "paper_orientation", "gutter", "gutter_pos",
             "header_text",
             "footer_text", "page_number", "page_break", "para_count",
             "char_count", "table_count", "table_dim", "table_align",
             "image_count", "image_size", "comment", "comment_count",
             "formula", "formula_count", "footnote", "endnote",
             "track_changes", "watermark", "toc", "bullet", "font_color",
             "highlight", "para_shading", "para_border", "image_wrap",
             "image_position", "image_shadow", "page_number_fmt",
             "page_number_start", "header_highlight", "caption",
             "widow_control", "top_line_punct", "page_border",
             "column_two", "column_sep"},
    "xlsx": {"sheet_count", "xlsx_freeze", "xlsx_formula", "xlsx_formula_count",
             "xlsx_protected", "xlsx_merged", "xlsx_table", "xlsx_cond_format",
             "xlsx_filter", "xlsx_font_size", "xlsx_col_width", "xlsx_row_height",
             "xlsx_align", "xlsx_border", "xlsx_fill", "xlsx_wrap_text",
             "xlsx_number_format", "xlsx_data_validation", "xlsx_print_area",
             "xlsx_hyperlink", "xlsx_merge_count", "xlsx_cond_type"},
    "pptx": {"slide_count", "pptx_font_size", "pptx_align_center", "pptx_bullet",
             "pptx_animation", "pptx_transition", "pptx_notes", "pptx_hyperlink",
             "pptx_image_count", "pptx_text_align", "pptx_slide_has_title",
             "pptx_aspect_ratio", "pptx_layout", "pptx_shape_count"},
}


def resolve(name):
    """返回原语 (fn, meta) 或 None（未知原语 → LLM 回落）。"""
    item = PRIMITIVES.get(name)
    if item is None:
        return None
    return item


def catalog(file_type=None, extra=None):
    """供 LLM 映射选用的原语目录：name + hint。extra 为用户库中的 LLM 模式原语。"""
    out = {}
    for name, (_, meta) in PRIMITIVES.items():
        if file_type is None or name in FILE_TYPES.get(file_type, set()):
            out[name] = meta["hint"]
    for name, meta in (extra or {}).items():
        if file_type is None or meta.get("file_type") == file_type:
            out[name] = meta.get("hint", "")
    return out


def extract(name, parsed, scope=None):
    """按原语抽取；scope 为 None 或定位失败时按原语自身语义处理。

    返回 (value, ok)：ok=False 表示无法确定性抽取（原语未知或定位失败）。
    """
    item = resolve(name)
    if item is None:
        return None, False
    fn, meta = item
    try:
        v = fn(parsed, scope)
    except Exception:
        v = None
    if v is None:
        return None, False
    return v, True


if __name__ == "__main__":
    import json
    import sys

    from parsers import parse_file

    if len(sys.argv) >= 3:
        d = parse_file(sys.argv[1])
        names = sys.argv[2].split(",")
        for n in names:
            v, ok = extract(n, d)
            print("%s = %r (ok=%s)" % (n, v, ok))
