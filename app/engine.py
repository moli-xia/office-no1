"""排版引擎：把预设格式套用到 .docx 文档（Word / WPS 通用格式）。

处理内容：
- 页面设置：纸张、方向、页边距、禁用文档网格（保证行距真实生效）
- 正文：中/西文字体、字号、行距、段前段后、首行缩进（按字符）、对齐
- 标题：识别 Word 内置标题样式 / 大纲级别，按级别套用字体与段落格式
- 页码：页脚/页眉插入 PAGE / NUMPAGES 域，支持多种样式、起始页码、首页隐藏
- 页眉/页脚文字

说明：正文只规范字体族与字号，不改动加粗等局部强调；标题则按预设强制统一。
"""

import math
import os
import re

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

PAGE_SIZES_CM = {
    "A4": (21.0, 29.7),
    "B5": (18.2, 25.7),
    "Letter": (21.59, 27.94),
    "16K": (18.4, 26.0),
}

ALIGN_MAP = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}

POSITION_LABELS = {
    "footer_center": "页面底端居中",
    "footer_right": "页面底端右侧",
    "footer_left": "页面底端左侧",
    "footer_outside": "底端：单页居右、双页居左（公文）",
    "header_center": "页面顶端居中",
    "header_right": "页面顶端右侧",
}

FORMAT_LABELS = {
    "plain": "纯数字：1",
    "dash": "短横线：- 1 -",
    "gb_dash": "公文一字线：— 1 —",
    "cn": "中文：第 1 页",
    "cn_total": "中文总页数：第 1 页 共 N 页",
    "en": "英文：Page 1 of N",
}

GRID_MODE_LABELS = {
    "none": "无（行距按设置值，推荐）",
    "lines_chars": "指定行和字符网格（公文：每页 22 行 × 每行 28 字）",
    "keep": "保持原文档设置",
}

HEADING_LEVELS = (1, 2, 3, 4)

# OOXML 中 w:rPr 子元素的规范顺序（截取本项目会用到的完整序列）
RPR_SEQ = [
    "w:rStyle", "w:rFonts", "w:b", "w:bCs", "w:i", "w:iCs", "w:caps", "w:smallCaps",
    "w:strike", "w:dstrike", "w:outline", "w:shadow", "w:emboss", "w:imprint",
    "w:noProof", "w:snapToGrid", "w:vanish", "w:webHidden", "w:color", "w:spacing",
    "w:w", "w:kern", "w:position", "w:sz", "w:szCs", "w:highlight", "w:u", "w:effect",
    "w:bdr", "w:shd", "w:fitText", "w:vertAlign", "w:rtl", "w:cs", "w:em", "w:lang",
    "w:eastAsianLayout", "w:specVanish", "w:oMath",
]

# w:pPr 中 snapToGrid 之后的兄弟元素（用于按顺序插入）
PPR_AFTER_SNAP = [
    "w:spacing", "w:ind", "w:contextualSpacing", "w:mirrorIndents", "w:suppressOverlap",
    "w:jc", "w:textDirection", "w:textAlignment", "w:textboxTightWrap", "w:outlineLvl",
    "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr", "w:pPrChange",
]

# w:sectPr 中 pgNumType / docGrid 之后的兄弟元素
SECT_AFTER_PAGENUM = [
    "w:cols", "w:formProt", "w:vAlign", "w:noEndnote", "w:titlePg", "w:textDirection",
    "w:bidi", "w:rtlGutter", "w:docGrid", "w:printerSettings", "w:sectPrChange",
]
SECT_AFTER_DOCGRID = ["w:printerSettings", "w:sectPrChange"]


# ---------------------------------------------------------------------------
# 底层 XML 工具
# ---------------------------------------------------------------------------

def _rpr_child(rpr, tag, attrs):
    """在 rPr 中查找或按 schema 顺序插入子元素并设置属性。"""
    el = rpr.find(qn(tag))
    if el is None:
        el = OxmlElement(tag)
        successors = RPR_SEQ[RPR_SEQ.index(tag) + 1:]
        rpr.insert_element_before(el, *successors)
    for a, v in attrs.items():
        el.set(qn(a), v)
    return el


def _apply_rpr(rpr, east=None, west=None, size_pt=None, bold=None, reset_style=False):
    """在给定 rPr 上设置中/西文字体、字号、加粗（bold=None 表示不改）。
    reset_style=True（标题）时去掉斜体并把颜色恢复为自动（黑），清掉 Word 内置标题样式的主题色。"""
    rfonts = rpr.get_or_add_rFonts()
    # 主题字体属性优先级高于显式字体，必须移除，否则设置不生效
    for theme in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
        if rfonts.get(qn(theme)) is not None:
            del rfonts.attrib[qn(theme)]
    if west:
        rfonts.set(qn("w:ascii"), west)
        rfonts.set(qn("w:hAnsi"), west)
        rfonts.set(qn("w:cs"), west)
    if east:
        rfonts.set(qn("w:eastAsia"), east)
    if size_pt is not None:
        half = str(int(round(size_pt * 2)))
        _rpr_child(rpr, "w:sz", {"w:val": half})
        _rpr_child(rpr, "w:szCs", {"w:val": half})
    if bold is not None:
        val = "1" if bold else "0"
        _rpr_child(rpr, "w:b", {"w:val": val})
        _rpr_child(rpr, "w:bCs", {"w:val": val})
    if reset_style:
        # 内置标题样式常带斜体（Heading 4）与蓝色主题色，公文 / 论文标题都不需要
        _rpr_child(rpr, "w:i", {"w:val": "0"})
        _rpr_child(rpr, "w:iCs", {"w:val": "0"})
        color = _rpr_child(rpr, "w:color", {"w:val": "auto"})
        for a in ("w:themeColor", "w:themeTint", "w:themeShade"):
            if color.get(qn(a)) is not None:
                del color.attrib[qn(a)]


def _set_run_font(r_el, east, west, size_pt, bold=None):
    _apply_rpr(r_el.get_or_add_rPr(), east, west, size_pt, bold)


def _set_snap(ppr, on):
    """设置段落“对齐到文档网格”。不用网格时必须关闭，否则固定/倍数行距在网格文档中会失真；
    用网格排版（公文 22 行 × 28 字）时则必须打开，行高由网格决定。"""
    snap = ppr.find(qn("w:snapToGrid"))
    if snap is None:
        snap = OxmlElement("w:snapToGrid")
        ppr.insert_element_before(snap, *PPR_AFTER_SNAP)
    snap.set(qn("w:val"), "1" if on else "0")


def _set_snap_off(ppr):
    _set_snap(ppr, False)


def _grid_on(preset):
    return preset["page"].get("grid_mode") == "lines_chars"


def _set_first_line_indent(p_el, chars, size_pt, char_pitch_twips=None):
    """按“字符”设置首行缩进（w:firstLineChars），并写入磅值兜底。chars<=0 时显式清零。

    字符网格模式下改写成略小于 N 个网格的缇值：Word 把“2 字符”按四舍五入后的网格间距换算，
    结果比 2 格多出零点几磅，首字会被挤到第 3 格（首行只剩 25 字），这是 Word 的老毛病。
    """
    ind = p_el.get_or_add_pPr().get_or_add_ind()
    if chars and chars > 0:
        if char_pitch_twips:
            if ind.get(qn("w:firstLineChars")) is not None:
                del ind.attrib[qn("w:firstLineChars")]
            ind.set(qn("w:firstLine"), str(int(math.floor(chars * char_pitch_twips))))
        else:
            ind.set(qn("w:firstLineChars"), str(int(round(chars * 100))))
            ind.set(qn("w:firstLine"), str(int(round(size_pt * chars * 20))))
    else:
        for a in ("w:firstLine", "w:firstLineChars"):
            ind.set(qn(a), "0")


def _grid_metrics(page_cfg, body_size_pt):
    """字符网格参数：(行距缇, 字符间距缇, charSpace)。与 Word“指定行和字符网格”写出的 XML 一致：
    linePitch = 版心高度 / 每页行数（缇，向下取整）；charSpace = (字符间距 - 正文字号) × 4096。"""
    w, h = PAGE_SIZES_CM[page_cfg["size"]]
    if page_cfg["orientation"] == "landscape":
        w, h = h, w
    lines = max(1, int(page_cfg.get("grid_lines", 22)))
    chars = max(1, int(page_cfg.get("grid_chars", 28)))
    # 用与 python-docx 写入 pgSz / pgMar 相同的取整方式换算成缇，保证和 Word 读到的数值一致
    tw = lambda cm_val: Cm(float(cm_val)).twips  # noqa: E731
    text_h_twips = tw(h) - tw(page_cfg["margin_top_cm"]) - tw(page_cfg["margin_bottom_cm"])
    text_w_twips = tw(w) - tw(page_cfg["margin_left_cm"]) - tw(page_cfg["margin_right_cm"])
    char_pitch_twips = text_w_twips / chars
    # 向下取整：字符间距只能略小于精确值，否则 Word 会算成每行少 1 个字（27 而不是 28）
    char_space = math.floor((char_pitch_twips / 20 - float(body_size_pt)) * 4096)
    return int(text_h_twips / lines), char_pitch_twips, char_space


def _set_line_spacing(pf, body_cfg, grid=False):
    if grid:
        # 网格模式：单倍行距 + 对齐网格，每行高度即网格行距（版心高度 / 每页行数）
        pf.line_spacing = 1.0
        return
    t, v = body_cfg["line_spacing_type"], body_cfg["line_spacing_value"]
    if t == "multiple":
        pf.line_spacing = float(v)  # 浮点数 => lineRule=auto，即多倍行距
    else:
        pf.line_spacing = Pt(float(v))
        pf.line_spacing_rule = (WD_LINE_SPACING.EXACTLY if t == "exact"
                                else WD_LINE_SPACING.AT_LEAST)


def _set_indent_chars(p_el, left_chars=None, right_chars=None):
    """按字符设置左 / 右缩进（w:leftChars / w:rightChars，单位 1/100 字符）。"""
    ind = p_el.get_or_add_pPr().get_or_add_ind()
    if left_chars is not None:
        ind.set(qn("w:leftChars"), str(int(round(left_chars * 100))))
    if right_chars is not None:
        ind.set(qn("w:rightChars"), str(int(round(right_chars * 100))))


def _outline_level_of(ppr):
    """读取 pPr 中的大纲级别：返回 1-9 表示标题级别，0 表示明确的正文（val=9），None 表示未设置。"""
    if ppr is None:
        return None
    ol = ppr.find(qn("w:outlineLvl"))
    if ol is None:
        return None
    try:
        lvl = int(ol.get(qn("w:val")))
    except (TypeError, ValueError):
        return None
    return lvl + 1 if 0 <= lvl <= 8 else 0


def _style_heading_level(style):
    """沿样式继承链判断标题级别：先看样式里的大纲级别，再看内置标题样式名。

    很多模板用“章标题”“节标题”等自定义样式，它们基于 Heading 1/2，只查当前样式名会漏掉。
    """
    depth = 0
    while style is not None and depth < 12:
        depth += 1
        try:
            lvl = _outline_level_of(style.element.pPr)
            name = (style.name or "").strip()
        except Exception:
            return 0
        if lvl is not None:
            return lvl
        m = re.match(r"^(?:heading|标题)\s*(\d+)", name, re.IGNORECASE)
        if m:
            return int(m.group(1))
        if name.lower() in ("title", "标题"):
            return 1
        try:
            style = style.base_style
        except Exception:
            return 0
    return 0


def _get_heading_level(p):
    """返回标题级别（1-3 用于本工具配置），0 表示正文，>3 且未配置则跳过。"""
    lvl = _outline_level_of(p._p.pPr)  # 段落直接设置的大纲级别优先
    if lvl is not None:
        return lvl
    try:
        style = p.style
    except Exception:
        return 0
    return _style_heading_level(style)


def _has_numbering(p):
    """段落是否属于编号 / 项目符号列表（直接设置或样式中设置）。"""
    ppr = p._p.pPr
    if ppr is not None and ppr.find(qn("w:numPr")) is not None:
        return True
    try:
        sppr = p.style.element.pPr
    except Exception:
        return False
    return sppr is not None and sppr.find(qn("w:numPr")) is not None


def _clear_paragraph(p):
    for child in list(p._p):
        if child.tag == qn("w:pPr"):
            continue
        p._p.remove(child)


# ---------------------------------------------------------------------------
# 样式与段落格式化
# ---------------------------------------------------------------------------

def _find_paragraph_style(doc, names):
    lower = {n.lower() for n in names}
    for s in doc.styles:
        try:
            if s.type == WD_STYLE_TYPE.PARAGRAPH and (s.name or "").lower() in lower:
                return s
        except Exception:
            continue
    return None


def _update_normal_style(doc, preset, char_pitch=None):
    style = _find_paragraph_style(doc, ("Normal", "正文"))
    if style is None:
        return
    body = preset["body"]
    _apply_rpr(style.element.get_or_add_rPr(),
               east=body["font_east"], west=body["font_west"], size_pt=body["size_pt"])
    pf = style.paragraph_format
    _set_line_spacing(pf, body, _grid_on(preset))
    pf.space_before = Pt(float(body["space_before_pt"]))
    pf.space_after = Pt(float(body["space_after_pt"]))
    pf.alignment = ALIGN_MAP[body["align"]]
    # 预设为 0 时也要显式清零，否则样式里原有的缩进会保留下来
    _set_first_line_indent(style.element, body["first_line_indent_chars"], body["size_pt"], char_pitch)


def _update_heading_style(doc, level, hcfg, body_cfg, grid=False, char_pitch=None):
    style = _find_paragraph_style(doc, (f"Heading {level}", f"标题 {level}"))
    if style is None or hcfg is None:
        return
    _apply_rpr(style.element.get_or_add_rPr(),
               east=hcfg["font_east"], west=hcfg["font_west"],
               size_pt=hcfg["size_pt"], bold=bool(hcfg["bold"]), reset_style=True)
    pf = style.paragraph_format
    _set_line_spacing(pf, body_cfg, grid)
    pf.space_before = Pt(float(hcfg["space_before_pt"]))
    pf.space_after = Pt(float(hcfg["space_after_pt"]))
    pf.alignment = ALIGN_MAP[hcfg["align"]]
    _set_first_line_indent(style.element, hcfg.get("first_line_indent_chars", 0), hcfg["size_pt"], char_pitch)


def _format_body_paragraph(p, body, in_table=False, grid=False, char_pitch=None):
    """规范正文段落。

    - 编号 / 项目符号段落：保留其原有缩进（悬挂缩进由编号定义控制，强行首行缩进会错位）
    - 表格内段落：不加首行缩进、不改对齐，只统一字体、字号与行距
    """
    pf = p.paragraph_format
    _set_line_spacing(pf, body, grid)
    pf.space_before = Pt(float(body["space_before_pt"]))
    pf.space_after = Pt(float(body["space_after_pt"]))
    ppr = p._p.get_or_add_pPr()
    _set_snap(ppr, grid)
    if in_table:
        _set_first_line_indent(p._p, 0, body["size_pt"])
    elif not _has_numbering(p):
        pf.alignment = ALIGN_MAP[body["align"]]
        _set_first_line_indent(p._p, body["first_line_indent_chars"], body["size_pt"], char_pitch)
    else:
        pf.alignment = ALIGN_MAP[body["align"]]
    for r in p._p.iter(qn("w:r")):
        _set_run_font(r, body["font_east"], body["font_west"], body["size_pt"])


def _format_heading_paragraph(p, hcfg, body_cfg, grid=False, char_pitch=None):
    pf = p.paragraph_format
    pf.alignment = ALIGN_MAP[hcfg["align"]]
    _set_line_spacing(pf, body_cfg, grid)  # 标题行距继承正文设置
    pf.space_before = Pt(float(hcfg["space_before_pt"]))
    pf.space_after = Pt(float(hcfg["space_after_pt"]))
    ppr = p._p.get_or_add_pPr()
    # 网格模式下，比正文大的标题（公文 2 号标题）不对齐字符网格，否则字距会被撑开
    _set_snap(ppr, grid and float(hcfg["size_pt"]) <= float(body_cfg["size_pt"]))
    _set_first_line_indent(p._p, hcfg.get("first_line_indent_chars", 0), hcfg["size_pt"], char_pitch)
    for r in p._p.iter(qn("w:r")):
        _apply_rpr(r.get_or_add_rPr(), hcfg["font_east"], hcfg["font_west"], hcfg["size_pt"],
                   bold=bool(hcfg["bold"]), reset_style=True)


def _iter_table_paragraphs(table):
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                yield p
            for t in cell.tables:
                yield from _iter_table_paragraphs(t)


# ---------------------------------------------------------------------------
# 页面 / 页码 / 页眉页脚
# ---------------------------------------------------------------------------

def _apply_page_setup(section, page_cfg, body_size_pt=12.0):
    w, h = PAGE_SIZES_CM[page_cfg["size"]]
    if page_cfg["orientation"] == "landscape":
        w, h = h, w
        section.orientation = WD_ORIENT.LANDSCAPE
    else:
        section.orientation = WD_ORIENT.PORTRAIT
    section.page_width = Cm(w)
    section.page_height = Cm(h)
    section.top_margin = Cm(float(page_cfg["margin_top_cm"]))
    section.bottom_margin = Cm(float(page_cfg["margin_bottom_cm"]))
    section.left_margin = Cm(float(page_cfg["margin_left_cm"]))
    section.right_margin = Cm(float(page_cfg["margin_right_cm"]))
    if page_cfg.get("header_distance_cm"):
        section.header_distance = Cm(float(page_cfg["header_distance_cm"]))
    if page_cfg.get("footer_distance_cm"):
        section.footer_distance = Cm(float(page_cfg["footer_distance_cm"]))

    mode = page_cfg.get("grid_mode", "none")
    if mode == "keep":
        return
    sect_pr = section._sectPr
    grid = sect_pr.find(qn("w:docGrid"))
    if grid is None:
        grid = OxmlElement("w:docGrid")
        sect_pr.insert_element_before(grid, *SECT_AFTER_DOCGRID)
    for attr in ("w:linePitch", "w:charSpace"):
        if grid.get(qn(attr)) is not None:
            del grid.attrib[qn(attr)]
    if mode == "lines_chars":
        line_pitch, _pitch, char_space = _grid_metrics(page_cfg, body_size_pt)
        grid.set(qn("w:type"), "linesAndChars")
        grid.set(qn("w:linePitch"), str(line_pitch))
        grid.set(qn("w:charSpace"), str(char_space))
    else:
        grid.set(qn("w:type"), "default")


def _set_page_start(section, start_at):
    sect_pr = section._sectPr
    pgnum = sect_pr.find(qn("w:pgNumType"))
    if pgnum is None:
        pgnum = OxmlElement("w:pgNumType")
        sect_pr.insert_element_before(pgnum, *SECT_AFTER_PAGENUM)
    if start_at and int(start_at) > 1:
        pgnum.set(qn("w:start"), str(int(start_at)))
    elif pgnum.get(qn("w:start")) is not None:
        del pgnum.attrib[qn("w:start")]


def _add_text_run(p, text, east, west, size_pt):
    r = p.add_run(text)
    _set_run_font(r._r, east, west, size_pt)
    return r


def _add_field_run(p, instr, east, west, size_pt):
    """插入 Word 域（PAGE / NUMPAGES），附带缓存显示值，未刷新域前也有内容。"""
    r = p.add_run()
    _set_run_font(r._r, east, west, size_pt)
    begin = OxmlElement("w:fldChar"); begin.set(qn("w:fldCharType"), "begin")
    instr_el = OxmlElement("w:instrText"); instr_el.set(qn("xml:space"), "preserve")
    instr_el.text = f" {instr} "
    sep = OxmlElement("w:fldChar"); sep.set(qn("w:fldCharType"), "separate")
    cache = OxmlElement("w:t"); cache.text = "1"
    end = OxmlElement("w:fldChar"); end.set(qn("w:fldCharType"), "end")
    for el in (begin, instr_el, sep, cache, end):
        r._r.append(el)


def _build_page_number_runs(p, fmt, east, west, size_pt):
    if fmt == "plain":
        _add_field_run(p, "PAGE", east, west, size_pt)
    elif fmt == "dash":
        _add_text_run(p, "- ", east, west, size_pt)
        _add_field_run(p, "PAGE", east, west, size_pt)
        _add_text_run(p, " -", east, west, size_pt)
    elif fmt == "gb_dash":
        # GB/T 9704：数字左右各放一条一字线（全角破折号“—”正好一个字宽）
        _add_text_run(p, "— ", east, west, size_pt)
        _add_field_run(p, "PAGE", east, west, size_pt)
        _add_text_run(p, " —", east, west, size_pt)
    elif fmt == "cn":
        _add_text_run(p, "第 ", east, west, size_pt)
        _add_field_run(p, "PAGE", east, west, size_pt)
        _add_text_run(p, " 页", east, west, size_pt)
    elif fmt == "cn_total":
        _add_text_run(p, "第 ", east, west, size_pt)
        _add_field_run(p, "PAGE", east, west, size_pt)
        _add_text_run(p, " 页 共 ", east, west, size_pt)
        _add_field_run(p, "NUMPAGES", east, west, size_pt)
        _add_text_run(p, " 页", east, west, size_pt)
    elif fmt == "en":
        _add_text_run(p, "Page ", east, west, size_pt)
        _add_field_run(p, "PAGE", east, west, size_pt)
        _add_text_run(p, " of ", east, west, size_pt)
        _add_field_run(p, "NUMPAGES", east, west, size_pt)
    else:
        _add_field_run(p, "PAGE", east, west, size_pt)


ALIGN_NAME = {"footer_center": "center", "header_center": "center",
              "footer_right": "right", "header_right": "right",
              "footer_left": "left", "footer_outside": "right"}


def _clear_part_paragraphs(part):
    """清空 header/footer 部件中的所有段落内容，返回可用的第一个空段落。"""
    paras = part.paragraphs
    if not paras:
        base = part.add_paragraph()
    else:
        base = paras[0]
        for extra in paras[1:]:
            extra._p.getparent().remove(extra._p)
    _clear_paragraph(base)
    return base


def _plan_header_footer(preset):
    """决定要改写哪些部件及其内容：{part_name: [(kind, payload), ...]}。

    规则：只改写本次要写入内容的部件；启用页码时页眉页脚都会重建（避免原文档
    另一侧残留旧页码）。页码未启用且没有页眉页脚文字时返回空，原文档页眉页脚原样保留。
    """
    pn = preset["page_number"]
    plan = {}
    for part_name in ("header", "footer"):
        contents = []
        text_cfg = preset[part_name]
        if text_cfg.get("enabled") and text_cfg.get("text", "").strip():
            contents.append(("text", text_cfg["text"].strip()))
        if pn["enabled"] and pn["position"].startswith(part_name):
            contents.append(("page_number", None))
        if contents or pn["enabled"]:
            plan[part_name] = contents
    return plan


def _plain_hf_paragraph(p, pn):
    """页眉页脚段落用单倍行距、无段前段后、不对齐网格，位置不受正文行距 / 网格影响。
    段落标记也设成页码字体字号，这样“空一字”按页码的 4 号字计算，而不是正文的 3 号字。"""
    pf = p.paragraph_format
    pf.line_spacing = 1.0
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    ppr = p._p.get_or_add_pPr()
    _set_snap_off(ppr)
    rpr = ppr.find(qn("w:rPr"))
    if rpr is None:
        rpr = OxmlElement("w:rPr")
        ppr.insert_element_before(rpr, "w:sectPr", "w:pPrChange")
    _apply_rpr(rpr, pn["font_east"], pn["font_west"], pn["size_pt"])


def _write_part(part, part_name, contents, pn, even_page=False):
    part.is_linked_to_previous = False
    base = _clear_part_paragraphs(part)
    if not contents:
        return
    paras = [base]
    while len(paras) < len(contents):
        paras.append(part.add_paragraph())
    outside = pn["position"] == "footer_outside"
    for p, (kind, payload) in zip(paras, contents):
        _clear_paragraph(p)
        _plain_hf_paragraph(p, pn)
        if kind == "text":
            # 页眉文字固定居中；页脚文字跟随页码位置
            key = f"{part_name}_center" if part_name == "header" else pn["position"]
            align = ALIGN_NAME.get(key, "center")
        else:
            align = ALIGN_NAME.get(pn["position"], "center")
        if outside and part_name == "footer":
            # GB/T 9704：单页码居右空一字，双页码居左空一字
            align = "left" if even_page else "right"
            if even_page:
                _set_indent_chars(p._p, left_chars=1)
            else:
                _set_indent_chars(p._p, right_chars=1)
        p.alignment = ALIGN_MAP.get(align, ALIGN_MAP["center"])
        if kind == "text":
            _add_text_run(p, payload, pn["font_east"], pn["font_west"], pn["size_pt"])
        else:
            _build_page_number_runs(p, pn["format"], pn["font_east"],
                                    pn["font_west"], pn["size_pt"])


def _apply_header_footer(doc, preset):
    """页码与页眉页脚文字写入第一节，其余分节链接到前一节，保证全文统一、连续编号。"""
    plan = _plan_header_footer(preset)
    if not plan:
        return
    pn = preset["page_number"]
    hide_first = bool(pn["enabled"] and pn.get("hide_on_first_page"))
    odd_even = bool(pn["enabled"] and pn["position"] == "footer_outside")
    first = doc.sections[0]

    for part_name, contents in plan.items():
        _write_part(getattr(first, part_name), part_name, contents, pn)
        if odd_even:
            _write_part(getattr(first, f"even_page_{part_name}"), part_name, contents, pn,
                        even_page=True)

    if pn["enabled"]:
        # 奇偶页不同是文档级设置：只有“单右双左”才打开，否则统一用同一页脚
        doc.settings.odd_and_even_pages_header_footer = odd_even
        _set_page_start(first, pn.get("start_at", 1))
        first.different_first_page_header_footer = hide_first
        if hide_first:
            for part_name in plan:
                part = getattr(first, f"first_page_{part_name}")
                part.is_linked_to_previous = False
                _clear_part_paragraphs(part)

    for section in doc.sections[1:]:
        for part_name in plan:
            getattr(section, part_name).is_linked_to_previous = True
            if odd_even:
                getattr(section, f"even_page_{part_name}").is_linked_to_previous = True
        if pn["enabled"]:
            # 页码由第一节统一控制：后续分节不再单独“首页不同”，也不重新起始编号
            section.different_first_page_header_footer = False
            pgnum = section._sectPr.find(qn("w:pgNumType"))
            if pgnum is not None and pgnum.get(qn("w:start")) is not None:
                del pgnum.attrib[qn("w:start")]


# ---------------------------------------------------------------------------
# 入口
# ---------------------------------------------------------------------------

def apply_preset(input_path, preset, output_path):
    """把预设套用到单个 docx 文件。返回结果 dict，不抛异常。"""
    try:
        doc = Document(input_path)
        stats = {"body": 0, "h1": 0, "h2": 0, "h3": 0, "h4": 0, "table": 0, "skipped": 0}
        grid = _grid_on(preset)
        body = preset["body"]
        pitch = _grid_metrics(preset["page"], body["size_pt"])[1] if grid else None

        _update_normal_style(doc, preset, pitch)
        headings = {h["level"]: h for h in preset.get("headings", []) if h["level"] in HEADING_LEVELS}
        for lvl in HEADING_LEVELS:
            _update_heading_style(doc, lvl, headings.get(lvl), body, grid, pitch)

        for p in doc.paragraphs:
            lvl = _get_heading_level(p)
            if lvl == 0:
                _format_body_paragraph(p, body, grid=grid, char_pitch=pitch)
                stats["body"] += 1
            elif lvl in headings:
                _format_heading_paragraph(p, headings[lvl], body, grid, pitch)
                stats[f"h{lvl}"] += 1
            else:
                stats["skipped"] += 1

        if body.get("format_tables"):
            for t in doc.tables:
                for p in _iter_table_paragraphs(t):
                    _format_body_paragraph(p, body, in_table=True, grid=grid, char_pitch=pitch)
                    stats["table"] += 1

        for section in doc.sections:
            _apply_page_setup(section, preset["page"], body["size_pt"])
        _apply_header_footer(doc, preset)

        doc.save(output_path)
        summary = (f"正文 {stats['body']} 段，标题 "
                   f"{stats['h1']}/{stats['h2']}/{stats['h3']}/{stats['h4']}")
        if stats["table"]:
            summary += f"，表格 {stats['table']} 段"
        if stats["skipped"]:
            summary += f"，跳过 {stats['skipped']} 段"
        return {"ok": True, "output": output_path, "stats": stats, "message": summary}
    except Exception as e:  # 单文件失败不影响批处理中的其他文件
        return {"ok": False, "output": None, "error": f"{type(e).__name__}: {e}"}


def _office_convert(src, dst, progids, file_format):
    """用本机 Word / WPS 的 COM 接口把 src 另存为 dst（file_format 为 WdSaveFormat 值）。成功返回 True。"""
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        return False
    src, dst = os.path.abspath(src), os.path.abspath(dst)
    pythoncom.CoInitialize()
    try:
        for progid in progids:
            app = None
            try:
                app = win32com.client.DispatchEx(progid)
                app.Visible = False
                try:
                    app.DisplayAlerts = 0  # 不弹兼容性/恢复对话框，避免后台线程被卡住
                except Exception:
                    pass
                doc = app.Documents.Open(src, False, True)  # ConfirmConversions, ReadOnly
                try:
                    doc.SaveAs2(dst, FileFormat=file_format)
                except Exception:
                    doc.SaveAs(dst, file_format)  # WPS 旧版本没有 SaveAs2
                doc.Close(False)
                if os.path.exists(dst):
                    return True
            except Exception:
                continue
            finally:
                if app is not None:
                    try:
                        app.Quit()
                    except Exception:
                        pass
    finally:
        pythoncom.CoUninitialize()
    return False


WD_FORMAT_DOC = 0            # wdFormatDocument：.doc；WPS 以 .wps 扩展名保存时即为 WPS 原生格式
WD_FORMAT_DOCX = 16          # wdFormatDocumentDefault
WPS_PROGIDS = ("KWPS.Application", "WPS.Application")
WORD_PROGIDS = ("Word.Application",)


def convert_doc_to_docx(doc_path, out_path):
    """用本机已安装的 Word 或 WPS 把 .doc / .wps 转成 .docx（需要 COM，失败返回 False）。"""
    progids = WORD_PROGIDS + WPS_PROGIDS
    if doc_path.lower().endswith(".wps"):
        progids = WPS_PROGIDS + WORD_PROGIDS  # .wps 优先交给 WPS
    return _office_convert(doc_path, out_path, progids, WD_FORMAT_DOCX)


def convert_docx_to_wps(docx_path, out_path):
    """用本机 WPS 把排版后的 .docx 存回 .wps（保持 WPS 原生格式，需要本机装有 WPS）。"""
    return _office_convert(docx_path, out_path, WPS_PROGIDS, WD_FORMAT_DOC)
