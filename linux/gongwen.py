# -*- coding: utf-8 -*-
"""
市委办公室材料排版 —— 核心格式引擎（无界面，可独立调用）

依据《市委办公室材料排版基本格式》实现：
  一、A4 纸（210mm×297mm）
  二、页边距：上3.9cm 下3.3cm 左2.7cm 右2.7cm
  三、每面22行、每行28字、撑满版心；文字颜色黑色
  四、标题：二号小标宋、居中、行距固定值33磅；成文日期：三号楷体
  五、正文：三号仿宋_GB2312（转发类通知用三号楷体_GB2312）、行距固定值28磅、
      段首空2字、回行顶格；层次序数 一、/（一）/1./（1）依次用
      黑体 / 楷体_GB2312加粗 / 仿宋_GB2312加粗；段间不空行
  六、页码：四号半角宋体阿拉伯数字，数字左右各一条一字线，
      上距版心下边缘7mm；单页码居右空一字，双页码居左空一字
  七、数字用 Times New Roman（所有中文 run 的西文字体统一设为 TNR，
      阿拉伯数字自动以 TNR 呈现，中文仍用中文字体）
"""
import os
import sys
import copy
import re
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph

# ────────────────────────── 版式常量（可按需微调） ──────────────────────────
PAGE_W_CM, PAGE_H_CM = 21.0, 29.7                 # 一、A4 幅面

LINES_PER_PAGE, CHARS_PER_LINE = 22, 28           # 每面22行、每行28字
# 版心高 22.5cm = 12756 twips；12756/560 = 22.78 → 每页最多 22 行（第23行放不下）
# 网格间距 = 行距 = 560tw(28磅)，行高与网格槽严格一致：
# 若网格间距(如579)与行距(560)不一致，跨页网格锚定偏差会导致部分页只排21行
LINE_PITCH = 560
# 每行28字：字距 = 版心宽442.2pt ÷ 28 = 15.79pt，比三号字(16pt)窄 0.207pt
# charSpace 单位为 1/4096 磅 → −0.207 × 4096 ≈ −849（实测 WPS/Word 按 28 字/行排版）
CHAR_SPACE = -849

SZ_TITLE_PT, SZ_BODY_PT, SZ_PAGENUM_PT = 22.0, 16.0, 14.0   # 二号 / 三号 / 四号
LINE_TITLE_PT = 33.0
LINE_BODY_PT = LINE_PITCH / 20.0                  # 正文行距固定值 = 网格间距 = 28磅

FONT_TNR, FONT_SONG = "Times New Roman", "宋体"

# ────────────────────────── 格式预设 ──────────────────────────
# 共同点：A4；每面22行、每行28字撑满版心（两格式版心均为 15.6×22.5cm）；
#   标题二号小标宋居中、固定值33磅；正文三号仿宋_GB2312、固定值28磅、段首空2字回行顶格；
#   页码四号半角宋体一字线、单页居右空一字/双页居左空一字、距版心下边缘约7mm；
#   数字/西文 Times New Roman；标题下空一行；小标题接排正文分体设字。
PRESETS = {
    # GB/T 9704—2012 党政机关公文格式（默认）
    # 页边距：天头37mm、订口28mm、版心156×225mm（GB/T 9704 第5.1条）
    # 层次序数：一、黑体／（一）楷体加粗／1.、（1）仿宋加粗
    #（GB 原文未要求二三四层加粗，此处按用户单位实务习惯统一加粗）
    # 成文日期：3号仿宋、右空四字；发文机关署名以成文日期为准居中（GB 15.7条）
    'gb9704': {
        'key': 'gb9704',
        'name': 'GB/T 9704—2012 党政机关公文格式',
        'm_top': 3.7, 'm_bottom': 3.5, 'm_left': 2.8, 'm_right': 2.6,
        'footer_dist': 2.23,                       # 下边距3.5cm 时一字线距版心约7mm
        'level_bold': {1: False, 2: True, 3: True, 4: True},
        'sig_date_font': 'fs',                     # 成文日期三号仿宋
        'date_right_chars': 4.0,                   # 右空四字
        'sig_center': True,                        # 署名以成文日期为准居中
        'sig_blank_lines': 3,                      # 落款与正文之间空行数
    },
    # 市委办公室材料排版基本格式
    # 层次序数二/三/四层加粗；日期三号楷体、右空二字
    'swb': {
        'key': 'swb',
        'name': '市委办公室材料排版基本格式',
        'm_top': 3.9, 'm_bottom': 3.3, 'm_left': 2.7, 'm_right': 2.7,
        'footer_dist': 2.0,                        # 下边距3.3cm 时一字线距版心约7mm
        'level_bold': {1: False, 2: True, 3: True, 4: True},
        'sig_date_font': 'kai',                    # 成文日期三号楷体
        'date_right_chars': 2.0,                   # 右空二字
        'sig_center': False,                       # 署名与日期一同右对齐
        'sig_blank_lines': 3,                      # 落款与正文之间空行数
    },
}
DEFAULT_PRESET = 'gb9704'

# ────────────────────────── OOXML 子元素顺序表 ──────────────────────────
PPR_ORDER = ['w:pStyle', 'w:keepNext', 'w:keepLines', 'w:pageBreakBefore', 'w:framePr',
             'w:widowControl', 'w:numPr', 'w:suppressLineNumbers', 'w:pBdr', 'w:shd',
             'w:tabs', 'w:suppressAutoHyphens', 'w:kinsoku', 'w:wordWrap',
             'w:overflowPunct', 'w:topLinePunct', 'w:autoSpaceDE', 'w:autoSpaceDN',
             'w:bidi', 'w:adjustRightInd', 'w:snapToGrid', 'w:spacing', 'w:ind',
             'w:contextualSpacing', 'w:mirrorIndents', 'w:suppressOverlap', 'w:jc',
             'w:textDirection', 'w:textAlignment', 'w:textboxTightWrap',
             'w:outlineLvl', 'w:divId', 'w:cnfStyle', 'w:rPr', 'w:sectPr', 'w:pPrChange']
RPR_ORDER = ['w:rStyle', 'w:rFonts', 'w:b', 'w:bCs', 'w:i', 'w:iCs', 'w:caps',
             'w:smallCaps', 'w:strike', 'w:dstrike', 'w:outline', 'w:shadow',
             'w:emboss', 'w:imprint', 'w:noProof', 'w:snapToGrid', 'w:vanish',
             'w:webHidden', 'w:color', 'w:spacing', 'w:w', 'w:kern', 'w:position',
             'w:sz', 'w:szCs', 'w:highlight', 'w:u', 'w:effect', 'w:bdr', 'w:shd',
             'w:fitText', 'w:vertAlign', 'w:rtl', 'w:cs', 'w:em', 'w:lang',
             'w:eastAsianLayout', 'w:specVanish', 'w:oMath']
SECTPR_ORDER = ['w:footnotePr', 'w:endnotePr', 'w:type', 'w:pgSz', 'w:pgMar',
                'w:paperSrc', 'w:pgBorders', 'w:lnNumType', 'w:pgNumType', 'w:cols',
                'w:formProt', 'w:vAlign', 'w:noEndnote', 'w:titlePg',
                'w:textDirection', 'w:bidi', 'w:rtlGutter', 'w:docGrid',
                'w:printerSettings', 'w:sectPrChange']
# settings.xml 中 evenAndOddHeaders 之后的元素（用于确定插入位置）
SETTINGS_AFTER_EOH = ['w:bookFoldRevPrinting', 'w:bookFoldPrinting', 'w:bookFoldPrintingSheets',
                      'w:drawingGridHorizontalSpacing', 'w:drawingGridVerticalSpacing',
                      'w:displayHorizontalDrawingGridEvery', 'w:displayVerticalDrawingGridEvery',
                      'w:doNotUseMarginsForDrawingGridOrigin', 'w:drawingGridHorizontalOrigin',
                      'w:drawingGridVerticalOrigin', 'w:doNotShadeFormData',
                      'w:noPunctuationKerning', 'w:characterSpacingControl', 'w:printTwoOnOne',
                      'w:strictFirstAndLastChars', 'w:noLineBreaksAfter', 'w:noLineBreaksBefore',
                      'w:savePreviewPicture', 'w:doNotValidateAgainstSchema', 'w:saveInvalidXml',
                      'w:ignoreMixedContent', 'w:alwaysShowPlaceholderText',
                      'w:doNotDemarcateInvalidXml', 'w:saveXmlDataOnly', 'w:useXSLTWhenSaving',
                      'w:saveThroughXslt', 'w:showXMLTags', 'w:alwaysMergeEmptyNamespace',
                      'w:saveCustomXmlProperties', 'w:alwaysSaveInDefaultEncoding',
                      'w:updateFields', 'w:hdrShapeDefaults', 'w:footnotePr', 'w:endnotePr',
                      'w:compat', 'w:docVars', 'w:rsids', 'w:mathPr', 'w:themeFontLang',
                      'w:clrSchemeMapping', 'w:doNotIncludeSubdocsInStats',
                      'w:doNotAutoCompressPictures', 'w:forceUpgrade', 'w:decimalSymbol',
                      'w:listSeparator']

# ────────────────────────── 结构识别 ──────────────────────────
DATE_RE = re.compile(r'^[（(]\s*\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日\s*[)）]$')
DATE_PLAIN_RE = re.compile(r'^\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日$')
L1_RE = re.compile(r'^[一二三四五六七八九十]{1,3}\s*、')            # 一、
L2_RE = re.compile(r'^[（(][一二三四五六七八九十]{1,3}[)）]')        # （一）
L3_RE = re.compile(r'^\d+\s*[.．、](?!\d)')                          # 1. （排除 3.5% 误判）
L4_RE = re.compile(r'^[（(]\d+[)）]')                                # （1）


def _is_date(t):
    """成文日期：带括号（标题区惯用）或纯日期（落款惯用）"""
    return bool(DATE_RE.match(t) or DATE_PLAIN_RE.match(t))


# 同行落款："单位名称　　2026年11月9日"（单位与日期以空格/全角空格分隔在同一行）
SAME_LINE_SIG_RE = re.compile(
    r'^(?P<sig>[^\s。；，、：！？""'']{2,20})[\s\u3000]+'
    r'(?P<date>[（(]\s*\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日\s*[)）]'
    r'|\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日)$')


def _level_of(t):
    if L1_RE.match(t):
        return 1
    if L2_RE.match(t):
        return 2
    if L3_RE.match(t):
        return 3
    if L4_RE.match(t):
        return 4
    return 0


# ────────────────────────── XML 工具 ──────────────────────────
def _name(el):
    t = el.tag
    return 'w:' + t.split('}')[-1] if '}' in t else t


def _rm(parent, *tags):
    """删除子元素（按前缀名）"""
    wanted = {qn(t) for t in tags}
    for child in list(parent):
        if child.tag in wanted:
            parent.remove(child)


def _set_child(parent, tag, attrs, order):
    """按 schema 顺序获取/创建子元素并设置属性"""
    el = parent.find(qn(tag))
    if el is None:
        el = OxmlElement(tag)
        pos = order.index(tag)
        insert_at = len(parent)
        for i, child in enumerate(parent):
            n = _name(child)
            if n in order and order.index(n) > pos:
                insert_at = i
                break
        parent.insert(insert_at, el)
    for k, v in attrs.items():
        el.set(qn(k), str(v))
    return el


def _runs(p):
    """段落内全部 w:r（含超链接等嵌套 run）"""
    return p._p.findall('.//' + qn('w:r'))


_WS = ' \u3000\u00a0\t'


def _strip_para_edges(p):
    """去掉段落首尾的多余空格/全角空格（避免破坏居中与顶格）"""
    ts = p._p.findall('.//' + qn('w:t'))
    for t in ts:
        if t.text:
            t.text = t.text.lstrip(_WS)
            if t.text:
                break
    for t in reversed(ts):
        if t.text:
            t.text = t.text.rstrip(_WS)
            if t.text:
                break


def _drop_para(p):
    el = p._p
    el.getparent().remove(el)


def _para_has_content(p):
    """空段落判定：无文字且无图/分节符/分页符等才视为空行"""
    if p.text.strip():
        return True
    el = p._p
    for tag in ('w:drawing', 'w:pict', 'w:object', 'w:br', 'w:fldSimple'):
        if el.find('.//' + qn(tag)) is not None:
            return True
    pPr = el.find(qn('w:pPr'))
    return pPr is not None and pPr.find(qn('w:sectPr')) is not None


# ────────────────────────── 格式化原语 ──────────────────────────
def style_run_el(r, east, size_pt, ascii_font=FONT_TNR, bold=None):
    """设置 run 字体：eastAsia=中文字体，ascii/hAnsi=西文字体（数字自动 TNR）"""
    rPr = r.find(qn('w:rPr'))
    if rPr is None:
        rPr = OxmlElement('w:rPr')
        r.insert(0, rPr)
    _rm(rPr, 'w:rFonts')
    _set_child(rPr, 'w:rFonts',
               {'w:ascii': ascii_font, 'w:hAnsi': ascii_font,
                'w:eastAsia': east, 'w:cs': ascii_font}, RPR_ORDER)
    # 清除原 run 上与公文无关的修饰
    _rm(rPr, 'w:i', 'w:iCs', 'w:u', 'w:highlight', 'w:vertAlign', 'w:shd', 'w:bdr',
        'w:em', 'w:spacing', 'w:w', 'w:position', 'w:outline', 'w:shadow',
        'w:emboss', 'w:imprint', 'w:caps', 'w:smallCaps', 'w:strike', 'w:dstrike',
        'w:b', 'w:bCs', 'w:sz', 'w:szCs', 'w:color')
    if bold is True:
        _set_child(rPr, 'w:b', {}, RPR_ORDER)
        _set_child(rPr, 'w:bCs', {}, RPR_ORDER)
    elif bold is False:
        _set_child(rPr, 'w:b', {'w:val': '0'}, RPR_ORDER)
        _set_child(rPr, 'w:bCs', {'w:val': '0'}, RPR_ORDER)
    if size_pt is not None:
        half = str(int(round(size_pt * 2)))
        _set_child(rPr, 'w:sz', {'w:val': half}, RPR_ORDER)
        _set_child(rPr, 'w:szCs', {'w:val': half}, RPR_ORDER)
    _set_child(rPr, 'w:color', {'w:val': '000000'}, RPR_ORDER)


def set_para_fmt(p, jc, line_pt, snap=True, ind=None, drawing_aware=True):
    """段落格式：对齐、行距固定值、缩进（字字符）、禁用自动编号等。

    jc: center/both/left/right；line_pt: 行距磅值（固定值）；
    snap: 是否对齐文档网格（正文 True → 22行撑满版心；标题 False）；
    ind: w:ind 属性字典（如 段首空2字）；None 表示清除缩进。
    """
    pPr = p._p.get_or_add_pPr()
    _rm(pPr, 'w:numPr', 'w:pBdr', 'w:shd')
    # 关闭孤行控制：Word/WPS 默认会把"段落首行单独落在页尾"的整段推到下页，
    # 导致该页只排 21 行；公文要求每面恒 22 行，段落跨页是允许的
    _set_child(pPr, 'w:widowControl', {'w:val': '0'}, PPR_ORDER)
    _set_child(pPr, 'w:snapToGrid', {'w:val': '1' if snap else '0'}, PPR_ORDER)
    _set_child(pPr, 'w:wordWrap', {'w:val': '0'}, PPR_ORDER)  # 数字/年份不在单词内断行折行
    # 整体重建 spacing：若只覆盖 before/after，原文档残留的"行"单位段距
    # （beforeLines/afterLines，如"段前0.5行"）与自动段距（Autospacing）会覆盖磅值
    _rm(pPr, 'w:spacing')
    if drawing_aware and (p._p.find('.//' + qn('w:drawing')) is not None
                          or p._p.find('.//' + qn('w:pict')) is not None
                          or p._p.find('.//' + qn('w:object')) is not None):
        # 含图片段落不用固定值（避免裁切图片），改单倍行距且不对齐网格
        _set_child(pPr, 'w:spacing',
                   {'w:before': '0', 'w:after': '0',
                    'w:beforeAutospacing': '0', 'w:afterAutospacing': '0',
                    'w:line': '240', 'w:lineRule': 'auto'},
                   PPR_ORDER)
        _set_child(pPr, 'w:snapToGrid', {'w:val': '0'}, PPR_ORDER)
    else:
        _set_child(pPr, 'w:spacing',
                   {'w:before': '0', 'w:after': '0',
                    'w:beforeAutospacing': '0', 'w:afterAutospacing': '0',
                    'w:line': str(int(round(line_pt * 20))), 'w:lineRule': 'exact'},
                   PPR_ORDER)
    _rm(pPr, 'w:ind')
    if ind:
        _set_child(pPr, 'w:ind', ind, PPR_ORDER)
    _set_child(pPr, 'w:jc', {'w:val': jc}, PPR_ORDER)


IND_FIRST2 = {'w:firstLineChars': '200', 'w:firstLine': '640'}   # 段首空2字


# ────────────────── 层次序数段：小标题 + 接排正文 分体设置 ──────────────────
def _heading_cut(t):
    """层次序数与正文同段时，小标题（含首个句末标点）结束位置；整段皆小标题返回 None"""
    for i, ch in enumerate(t):
        if ch in '。！？':
            return i + 1 if i + 1 < len(t) else None
    return None


def _truncate_run_texts(r_el, lo, hi):
    """仅保留 run 文本中第 [lo, hi) 个字符（按 run 内 w:t 顺序累计）"""
    cum = 0
    for t in list(r_el.findall(qn('w:t'))):
        s = t.text or ''
        a, b = cum, cum + len(s)
        cum = b
        na, nb = max(a, lo), min(b, hi)
        if na >= nb:
            r_el.remove(t)
        else:
            t.text = s[na - a:nb - a]


def _split_runs_at(p, cut):
    """在段落文本第 cut 个字符处把所在 run 一分为二（文本边界恰在 run 边界时不动作）"""
    cum = 0
    for r in _runs(p):
        ts = r.findall(qn('w:t'))
        if not ts:
            continue
        rlen = sum(len(t.text or '') for t in ts)
        if cum + rlen < cut:
            cum += rlen
            continue
        if cum + rlen == cut:
            return
        k = cut - cum
        new_r = copy.deepcopy(r)
        r.addnext(new_r)
        _truncate_run_texts(r, 0, k)
        _truncate_run_texts(new_r, k, rlen)
        return


def _style_level_para(p, east, bold, body_font):
    """层次序数段：序数小标题用层次字体；接排的其余正文用正文字体"""
    t = ''.join(x.text or '' for x in p._p.findall('.//' + qn('w:t')))
    cut = _heading_cut(t)
    if cut is None:
        for r in _runs(p):
            style_run_el(r, east, SZ_BODY_PT, bold=bold)
        return
    _split_runs_at(p, cut)
    cum = 0
    for r in _runs(p):
        rlen = sum(len(x.text or '') for x in r.findall(qn('w:t')))
        if cum < cut:
            style_run_el(r, east, SZ_BODY_PT, bold=bold)
        else:
            style_run_el(r, body_font, SZ_BODY_PT, bold=False)
        cum += rlen


def _new_body_blank_el():
    """新建一个正文格式（28磅固定值行距）的空段元素"""
    blank = OxmlElement('w:p')
    pPr = blank.get_or_add_pPr()
    _set_child(pPr, 'w:widowControl', {'w:val': '0'}, PPR_ORDER)
    _set_child(pPr, 'w:snapToGrid', {'w:val': '1'}, PPR_ORDER)
    _set_child(pPr, 'w:spacing',
               {'w:before': '0', 'w:after': '0',
                'w:beforeAutospacing': '0', 'w:afterAutospacing': '0',
                'w:line': '560', 'w:lineRule': 'exact'},
               PPR_ORDER)
    return blank


def _ensure_blank_after_title(end_p):
    """标题区（标题/日期）之后确保恰好空一行（无则插入；GB/T 9704：标题下空一行）"""
    end_el = end_p._p
    nxt = end_el.getnext()
    if nxt is None:
        return
    if nxt.tag == qn('w:p'):
        nxt_p = Paragraph(nxt, None)
        if not nxt_p.text.strip() and not _para_has_content(nxt_p):
            return                      # 已有空行
    end_el.addnext(_new_body_blank_el())


def _ensure_blank_before_sig(sig_p, count, sig_els):
    """落款署名与正文之间恰好空 count 行。

    多余空行删除、不足补插；紧邻上方为另一署名段（联合行文）时保持紧凑不插行。
    """
    el = sig_p._p
    blanks = []
    cur = el.getprevious()
    while cur is not None and cur.tag == qn('w:p'):
        po = Paragraph(cur, None)
        if po.text.strip() or _para_has_content(po):
            break
        blanks.append(cur)
        cur = cur.getprevious()
    if cur is not None and any(cur == se for se in sig_els):
        return                          # 联合署名紧邻上一署名，不插空行
    while len(blanks) > count:
        extra = blanks.pop()            # 删除最上方的多余空行
        extra.getparent().remove(extra)
    for _ in range(count - len(blanks)):
        el.addprevious(_new_body_blank_el())


# ────────────────────────── 页面设置与页码 ──────────────────────────
def _setup_section(sec, pre):
    sec.page_width = Cm(PAGE_W_CM)
    sec.page_height = Cm(PAGE_H_CM)
    sec.top_margin = Cm(pre['m_top'])
    sec.bottom_margin = Cm(pre['m_bottom'])
    sec.left_margin = Cm(pre['m_left'])
    sec.right_margin = Cm(pre['m_right'])
    sec.footer_distance = Cm(pre['footer_dist'])
    sec.different_first_page_header_footer = False
    _set_child(sec._sectPr, 'w:docGrid',
               {'w:type': 'linesAndChars', 'w:linePitch': LINE_PITCH,
                'w:charSpace': CHAR_SPACE}, SECTPR_ORDER)


def _ensure_even_odd_headers(doc):
    settings = doc.settings.element
    if settings.find(qn('w:evenAndOddHeaders')) is None:
        el = OxmlElement('w:evenAndOddHeaders')
        insert_at = len(settings)
        for i, child in enumerate(settings):
            if _name(child) in SETTINGS_AFTER_EOH:
                insert_at = i
                break
        settings.insert(insert_at, el)


def _add_footer_run(p, text, sz_half):
    r = OxmlElement('w:r')
    style_run_el(r, FONT_SONG, SZ_PAGENUM_PT, ascii_font=FONT_SONG, bold=False)
    if text is not None:
        t = OxmlElement('w:t')
        t.set(qn('xml:space'), 'preserve')
        t.text = text
        r.append(t)
    p._p.append(r)
    return r


def _add_page_field(p):
    """插入 PAGE 域：{ PAGE }"""
    r = _add_footer_run(p, None, 28)
    fc = OxmlElement('w:fldChar')
    fc.set(qn('w:fldCharType'), 'begin')
    r.append(fc)
    r = _add_footer_run(p, None, 28)
    it = OxmlElement('w:instrText')
    it.set(qn('xml:space'), 'preserve')
    it.text = ' PAGE '
    r.append(it)
    r = _add_footer_run(p, None, 28)
    fc = OxmlElement('w:fldChar')
    fc.set(qn('w:fldCharType'), 'separate')
    r.append(fc)
    _add_footer_run(p, '1', 28)          # 域缓存值，打开后自动更新
    r = _add_footer_run(p, None, 28)
    fc = OxmlElement('w:fldChar')
    fc.set(qn('w:fldCharType'), 'end')
    r.append(fc)


def _build_pagenum_footer(footer, side):
    """四号半角宋体页码「— N —」：单页居右空一字，双页居左空一字"""
    footer.is_linked_to_previous = False
    for tbl in list(footer.tables):
        tbl._tbl.getparent().remove(tbl._tbl)
    paras = footer.paragraphs
    p = paras[0]
    for extra in paras[1:]:
        _drop_para(extra)
    for child in list(p._p):
        if _name(child) != 'w:pPr':
            p._p.remove(child)
    pPr = p._p.get_or_add_pPr()
    _rm(pPr, 'w:spacing', 'w:ind', 'w:jc', 'w:snapToGrid', 'w:tabs')
    _set_child(pPr, 'w:spacing',
               {'w:before': '0', 'w:after': '0',
                'w:beforeAutospacing': '0', 'w:afterAutospacing': '0'},
               PPR_ORDER)
    _set_child(pPr, 'w:snapToGrid', {'w:val': '0'}, PPR_ORDER)
    _set_child(pPr, 'w:jc', {'w:val': 'right' if side == 'odd' else 'left'}, PPR_ORDER)
    # 显式覆盖两侧及首行缩进为 0：防止文档模板 Normal/Footer 样式残留的
    # 首行缩进等通过继承叠加进来（表现为偶数页页码左侧多出大段空白）
    base = {'w:firstLineChars': '0', 'w:firstLine': '0',
            'w:leftChars': '0', 'w:left': '0',
            'w:rightChars': '0', 'w:right': '0'}
    if side == 'odd':                     # 单页码居右空一字（右缩进一字，跨 Word/WPS 稳定）
        base.update({'w:rightChars': '100', 'w:right': '280'})
    else:                                 # 双页码居左空一字
        base.update({'w:leftChars': '100', 'w:left': '280'})
    _set_child(pPr, 'w:ind', base, PPR_ORDER)
    _add_footer_run(p, '— ', 28)
    _add_page_field(p)
    _add_footer_run(p, ' —', 28)


def build_page_numbers(doc, log=None):
    _ensure_even_odd_headers(doc)
    for i, sec in enumerate(doc.sections):
        _build_pagenum_footer(sec.footer, 'odd')
        _build_pagenum_footer(sec.even_page_footer, 'even')
        if log:
            log(f'  · 第{i + 1}节：页码已编排（单页右空一字 / 双页左空一字）')


# ────────────────────────── 标题块与主送识别 ──────────────────────────
def _find_title_block(ss):
    """返回 (标题段索引列表, 标题区日期段索引或 None)"""
    n = len(ss)
    i = 0
    while i < n and not ss[i]:
        i += 1
    if i >= n:
        return [], None
    if len(ss[i]) > 60:          # 超长首段不是标题（标题不应超过60字）
        return [], None
    title = [i]
    i += 1
    while i < n and len(title) < 3:
        t = ss[i]
        if not t or _is_date(t):
            break
        if (len(t) > 40 or t.endswith(('。', '；', '！', '？', '：', ':'))
                or _level_of(t)):
            break
        title.append(i)
        i += 1
    j = i
    blanks = 0
    while j < n and not ss[j]:
        j += 1
        blanks += 1
    if j < n and blanks <= 1 and _is_date(ss[j]):
        return title, j
    return title, None


def _is_zhusong(t):
    return (t.endswith(('：', ':')) and len(t) <= 60
            and '。' not in t and not _level_of(t))


def _next_nonempty(ss, k):
    j = k
    while j < len(ss) and not ss[j]:
        j += 1
    return j


# ────────────────────────── 段落格式化主流程 ──────────────────────────
def _style_normal(doc, body_font):
    st = doc.styles['Normal']
    st.font.name = FONT_TNR
    st.font.size = Pt(SZ_BODY_PT)
    st.font.color.rgb = RGBColor(0, 0, 0)
    rpr = st.element.get_or_add_rPr()
    _rm(rpr, 'w:rFonts')
    _set_child(rpr, 'w:rFonts',
               {'w:ascii': FONT_TNR, 'w:hAnsi': FONT_TNR,
                'w:eastAsia': body_font, 'w:cs': FONT_TNR}, RPR_ORDER)
    # 先清样式残留的 spacing（含行单位段距/自动段距）与缩进，再由 setter 重建
    st_pPr = st.element.find(qn('w:pPr'))
    if st_pPr is not None:
        _rm(st_pPr, 'w:ind', 'w:spacing')
    st.paragraph_format.space_before = Pt(0)
    st.paragraph_format.space_after = Pt(0)
    st_pPr = st.element.find(qn('w:pPr'))
    if st_pPr is not None:
        sp = st_pPr.find(qn('w:spacing'))
        if sp is not None:
            sp.set(qn('w:beforeAutospacing'), '0')
            sp.set(qn('w:afterAutospacing'), '0')


def _iter_tables(tables):
    for tbl in tables:
        yield tbl
        for nested in tbl.tables:
            yield from _iter_tables(nested)


def _fmt_tables(doc, body_font, log=None):
    cnt = 0
    for tbl in _iter_tables(doc.tables):
        for cell in tbl._cells:
            for p in cell.paragraphs:
                pPr = p._p.get_or_add_pPr()
                _set_child(pPr, 'w:snapToGrid', {'w:val': '0'}, PPR_ORDER)
                _rm(pPr, 'w:spacing')
                _set_child(pPr, 'w:spacing',
                           {'w:before': '0', 'w:after': '0',
                            'w:beforeAutospacing': '0', 'w:afterAutospacing': '0'},
                           PPR_ORDER)
                for r in _runs(p):
                    style_run_el(r, body_font, None, bold=None)  # 仅统一字体与黑色
        cnt += 1
    if cnt and log:
        log(f'  · 表格 {cnt} 个：已统一中文字体/黑色/数字 TNR（字号与加粗保持原样）')


def _visual_width(t):
    """估算文本视觉宽度（全角=1字，半角=0.5字），用于落款居中近似"""
    w = 0.0
    for c in t:
        w += 0.5 if ord(c) < 0x2E80 else 1.0
    return w


def _truncate_para_text(p_el, lo, hi):
    """仅保留段落文本第 [lo, hi) 个字符（跨 run 累计，按 w:t 顺序）"""
    cum = 0
    for r in p_el.findall('.//' + qn('w:r')):
        for t in list(r.findall(qn('w:t'))):
            s = t.text or ''
            a, b = cum, cum + len(s)
            cum = b
            na, nb = max(a, lo), min(b, hi)
            if na >= nb:
                r.remove(t)
            else:
                t.text = s[na - a:nb - a]


def _split_para_at(p, cut):
    """在段落文本第 cut 字符处拆分为两段（原段保留前半、新段承载后半）。

    用于"单位名称　　日期"同行落款 → 拆为署名、日期两行。
    返回新段落的 Paragraph 对象。
    """
    _split_runs_at(p, cut)                      # 先保证 cut 恰在 run 边界
    new_el = copy.deepcopy(p._p)
    p._p.addnext(new_el)
    total = sum(len(t.text or '') for t in p._p.findall('.//' + qn('w:t')))
    _truncate_para_text(p._p, 0, cut)
    _truncate_para_text(new_el, cut, total)
    return Paragraph(new_el, p._parent if hasattr(p, '_parent') else None)


def _fmt_signature(p, sig_text, date_text, pre, body_font):
    """落款署名：GB=以成文日期为准居中；市委办=与日期一同右对齐"""
    if pre['sig_center']:
        w_date = _visual_width(date_text)
        w_sig = _visual_width(sig_text)
        rc = pre['date_right_chars'] + max(0.0, (w_date - w_sig) / 2)
    else:
        rc = pre['date_right_chars']
    rc = round(rc * 100)
    set_para_fmt(p, 'right', LINE_BODY_PT,
                 ind={'w:rightChars': str(rc), 'w:right': str(int(rc * 3.2))})
    for r in _runs(p):
        style_run_el(r, body_font, SZ_BODY_PT, bold=False)


def _fmt_trailing_date(p, pre, fonts):
    """落款成文日期：GB=三号仿宋右空四字；市委办=三号楷体右空二字"""
    drc = round(pre['date_right_chars'] * 100)
    set_para_fmt(p, 'right', LINE_BODY_PT,
                 ind={'w:rightChars': str(drc), 'w:right': str(int(drc * 3.2))})
    dfont = fonts['kai'] if pre['sig_date_font'] == 'kai' else fonts['fs']
    for r in _runs(p):
        style_run_el(r, dfont, SZ_BODY_PT, bold=False)


def format_document(doc, fonts, preset=DEFAULT_PRESET, forwarding=False,
                    page_numbers=True, remove_blank=True, log=None):
    """在 Document 对象上执行全文排版，返回统计信息。

    preset：'gb9704'（GB/T 9704—2012，默认）或 'swb'（市委办公室材料排版基本格式）。
    """
    log = log or (lambda *_: None)
    pre = PRESETS[preset]
    body_font = fonts['kai'] if forwarding else fonts['fs']

    _style_normal(doc, body_font)
    log(f'⓪ 格式标准：{pre["name"]}')
    log(f'① 页面：A4，页边距 上{pre["m_top"]}/下{pre["m_bottom"]}/'
        f'左{pre["m_left"]}/右{pre["m_right"]}cm')
    for i, sec in enumerate(doc.sections):
        _setup_section(sec, pre)
    log(f'② 网格：每面{LINES_PER_PAGE}行、每行{CHARS_PER_LINE}字，撑满版心')
    if page_numbers:
        build_page_numbers(doc, log)
    else:
        log('② 页码：按设置未编排')

    paras = doc.paragraphs
    ss = [p.text.replace('\u3000', ' ').replace('\xa0', ' ').strip() for p in paras]
    n = len(paras)
    title_idx, date_idx = _find_title_block(ss)
    title_set = set(title_idx)

    zhusong = None
    if title_idx:
        k = (date_idx if date_idx is not None else title_idx[-1]) + 1
        j = _next_nonempty(ss, k)
        if j < n and _is_zhusong(ss[j]):
            zhusong = j

    # 标题区内（含标题与日期之间）、标题之前的空行保留（模拟红头纸空行）；其余正文空行可删
    protect = set(title_idx)
    if title_idx:
        for x in range(title_idx[0]):
            protect.add(x)
    if date_idx is not None:
        for x in range(title_idx[-1] + 1, date_idx):
            protect.add(x)
    keep_blank = None
    if title_idx:
        k = (date_idx if date_idx is not None else title_idx[-1]) + 1
        if k < n and not ss[k] and not _para_has_content(paras[k]):
            keep_blank = k

    stats = {'title': len(title_idx), 'date': 0, 'zhusong': bool(zhusong),
             'blanks_removed': 0, 'levels': 0, 'forwarding': forwarding,
             'sig_split': 0}

    # 落款署名：紧邻落款日期之前的短段落（无句末标点）→ {署名段索引: 日期段索引}
    sig_map = {}
    for j in range(n):
        if _is_date(ss[j]) and j != date_idx and j > (date_idx or 0):
            k = j - 1
            if 0 <= k < n and ss[k] and len(ss[k]) <= 20 and not _level_of(ss[k]) \
                    and not ss[k].endswith(('。', '；', '！', '？', '：', ':')):
                sig_map[k] = j

    # 文末段落（最后5个非空段）：同行落款"单位　　日期"仅在此范围拆分，避免误伤正文
    nonsparse = [i for i in range(n) if ss[i]]
    tail_set = set(nonsparse[-5:])
    sig_paras = []                       # 落款署名段（用于其后统一处理空行）
    if title_idx:
        log(f'③ 标题：第{"、".join(str(i + 1) for i in title_idx)}段，'
            f'二号{fonts["title"]}、居中、固定值33磅')

    for i, p in enumerate(paras):
        if ss[i]:
            _strip_para_edges(p)
        if i in title_set:                                   # 标题
            set_para_fmt(p, 'center', LINE_TITLE_PT, snap=False)
            for r in _runs(p):
                style_run_el(r, fonts['title'], SZ_TITLE_PT, bold=False)
        elif i == date_idx:                                  # 标题区成文日期：三号楷体居中
            set_para_fmt(p, 'center', LINE_BODY_PT)
            for r in _runs(p):
                style_run_el(r, fonts['kai'], SZ_BODY_PT, bold=False)
            stats['date'] += 1
        elif i == zhusong:                                   # 主送机关：顶格
            set_para_fmt(p, 'left', LINE_BODY_PT)
            for r in _runs(p):
                style_run_el(r, body_font, SZ_BODY_PT, bold=False)
        elif not ss[i]:                                      # 空行
            if _para_has_content(p) or i in protect or i == keep_blank or not remove_blank:
                set_para_fmt(p, 'both', LINE_BODY_PT, ind=IND_FIRST2)
                for r in _runs(p):
                    style_run_el(r, body_font, SZ_BODY_PT, bold=False)
            else:
                _drop_para(paras[i])
                stats['blanks_removed'] += 1
        elif i in sig_map:                                   # 落款署名（与日期分行）
            _fmt_signature(p, ss[i], ss[sig_map[i]], pre, body_font)
            sig_paras.append(p)
        elif _is_date(ss[i]):                                # 落款成文日期（独占一行）
            _fmt_trailing_date(p, pre, fonts)
            stats['date'] += 1
        elif i in tail_set and SAME_LINE_SIG_RE.match(ss[i]):   # 同行落款：单位+日期拆成两行
            m = SAME_LINE_SIG_RE.match(ss[i])
            p2 = _split_para_at(p, m.start('date'))
            _strip_para_edges(p)
            _strip_para_edges(p2)
            _fmt_signature(p, m.group('sig'), m.group('date'), pre, body_font)
            _fmt_trailing_date(p2, pre, fonts)
            stats['date'] += 1
            stats['sig_split'] = stats.get('sig_split', 0) + 1
            sig_paras.append(p)
        else:                                                # 正文
            lv = _level_of(ss[i])
            set_para_fmt(p, 'both', LINE_BODY_PT, ind=IND_FIRST2)
            if lv == 0:
                for r in _runs(p):
                    style_run_el(r, body_font, SZ_BODY_PT, bold=False)
            else:
                if lv == 1:
                    east = fonts['hei']                      # 第一层：黑体
                elif lv == 2:
                    east = fonts['kai']                      # 第二层：楷体
                else:
                    east = fonts['fs']                       # 第三/四层：仿宋
                bold = pre['level_bold'][lv]                 # 是否加粗由格式标准决定
                stats['levels'] += 1
                _style_level_para(p, east, bold, body_font)  # 小标题+接排正文分体

    # 标题区之后确保恰好空一行（标题下空一行）
    if title_idx:
        _ensure_blank_after_title(paras[date_idx] if date_idx is not None
                                  else paras[title_idx[-1]])

    # 落款与正文之间空 N 行（默认 3 行；不足补插、多余删除）
    if sig_paras:
        cnt = pre.get('sig_blank_lines', 3)
        sig_els = [sp._p for sp in sig_paras]
        for sp in sig_paras:
            _ensure_blank_before_sig(sp, cnt, sig_els)
        log(f'⑦ 落款与正文之间已空 {cnt} 行')

    log(f'④ 正文：三号{"楷体" if forwarding else "仿宋"}、固定值28磅、段首空2字、'
        f'回行顶格；层次序数黑体/楷体加粗/仿宋加粗 {stats["levels"]} 处；'
        f'删除空行 {stats["blanks_removed"]} 个；标题下空一行')
    if stats['sig_split']:
        log('⑥ 落款：同行"单位＋日期"已拆分为两行（署名在上、成文日期在下）')
    log('⑤ 数字/西文：Times New Roman；全文黑色')
    _fmt_tables(doc, body_font, log)
    return stats


# ────────────────────────── 旧格式（.doc/.wps）自动转换 ──────────────────────────
class ConverterUnavailable(ValueError):
    """旧格式转换引擎不可用（附原文件路径，供 GUI 提供「用 WPS 打开」选项）。"""

    def __init__(self, msg, src=None):
        super().__init__(msg)
        self.src = src


def _convert_via_com(src, log=None):
    """Windows：调用 WPS/Word 将 OLE2 旧格式转换为临时 .docx。

    使用 DispatchEx 独立进程，避免影响用户已打开的文档窗口。
    """
    try:
        import win32com.client
    except ImportError:
        raise ConverterUnavailable('旧格式自动转换需要 pywin32 组件，请先将文件另存为 .docx', src)

    app = None
    last_err = None
    for progid in ('KWPS.Application', 'Word.Application', 'WPS.Application'):
        try:
            try:
                app = win32com.client.DispatchEx(progid)   # 独立实例，不动用户窗口
            except Exception:
                app = win32com.client.Dispatch(progid)
            break
        except Exception as e:  # noqa: BLE001
            last_err = e
    if app is None:
        raise ConverterUnavailable(
            '未检测到可用的 WPS/Word，无法自动转换旧格式。\n'
            '请打开该文件另存为 .docx 后重新处理。', src)

    import tempfile
    fd, tmp = tempfile.mkstemp(suffix='.docx')
    os.close(fd)
    os.remove(tmp)                       # 仅占位取唯一路径
    try:
        app.Visible = False
        doc = app.Documents.Open(os.path.abspath(src), ReadOnly=True)
        try:
            doc.SaveAs2(tmp, 12)         # 12 = docx（WPS 兼容）
        except Exception:  # noqa: BLE001
            doc.SaveAs(tmp, 12)
        doc.Close(False)
        if log:
            log('  · 已通过 WPS/Word 将旧格式转换为 .docx 再排版')
        return tmp
    finally:
        try:
            app.Quit()
        except Exception:  # noqa: BLE001
            pass


def _convert_via_soffice(src, log=None):
    """Linux：调用 LibreOffice（soffice）无头模式将旧格式转换为临时 .docx。

    WPS Linux 版不支持后台命令行转换，故 Linux 端以 LibreOffice 为转换引擎。
    独立用户配置目录（-env:UserInstallation），不影响正在运行的 LibreOffice/WPS。
    """
    import tempfile
    import subprocess
    import shutil
    outdir = tempfile.mkdtemp(prefix='clpb_conv_')
    src_abs = os.path.abspath(src)
    stem = os.path.splitext(os.path.basename(src))[0]
    # PATH 中的命令 + 常见安装绝对路径
    candidates = ['soffice', 'libreoffice',
                  '/usr/bin/soffice', '/usr/bin/libreoffice',
                  '/usr/lib/libreoffice/program/soffice',
                  '/opt/libreoffice/program/soffice']
    tried = []
    for cmd in candidates:
        exe = shutil.which(cmd) or (cmd if os.path.isfile(cmd) else None)
        if not exe:
            continue
        tried.append(cmd)
        profile = outdir + '/loprofile'
        cmdline = [exe, '--headless', '--norestore', '--nologo',
                   '--nolockcheck', '--nodefault',
                   f'-env:UserInstallation=file://{profile}',
                   '--convert-to', 'docx', '--outdir', outdir, src_abs]
        try:
            r = subprocess.run(cmdline, capture_output=True, timeout=180)
        except (subprocess.TimeoutExpired, OSError):
            continue
        out = os.path.join(outdir, stem + '.docx')
        if r.returncode == 0 and os.path.exists(out):
            fd, final = tempfile.mkstemp(suffix='.docx')
            os.close(fd)
            os.remove(final)
            shutil.move(out, final)
            shutil.rmtree(outdir, ignore_errors=True)
            if log:
                log(f'  · 已通过 LibreOffice 将旧格式转换为 .docx 再排版')
            return final
    shutil.rmtree(outdir, ignore_errors=True)
    raise ConverterUnavailable(
        '无法自动转换旧格式（.doc/.wps）：系统未安装 LibreOffice，'
        '而 WPS Linux 版不支持后台转换。\n'
        '两种解决办法（任选其一）：\n'
        '① 安装转换引擎：sudo apt install libreoffice-writer，之后重新拖入；\n'
        '② 用 WPS 打开该文件，「另存为」→.docx，再把 .docx 拖入本工具。', src)


def _convert_to_docx_via_office(src, log=None):
    """旧格式转换入口：Windows 走 WPS/Word COM，Linux 走 LibreOffice。"""
    if sys.platform.startswith('win'):
        return _convert_via_com(src, log)
    return _convert_via_soffice(src, log)


def _prepare_source(src, log=None):
    """按文件魔数分流。返回 (工作文件, 临时文件或None, 容器类型 'zip'/'ole')。

    ZIP(OOXML) 直接用；OLE2(.doc/.wps) 先转换为 .docx；其余报错。
    """
    with open(src, 'rb') as f:
        magic = f.read(4)
    if magic == b'PK\x03\x04':
        return src, None, 'zip'
    if magic == b'\xd0\xcf\x11\xe0':
        if log:
            log('▶ 检测到旧格式二进制文档（.doc/.wps），正在调用 WPS/Word/LibreOffice 转换…')
        tmp = _convert_to_docx_via_office(src, log)
        return tmp, tmp, 'ole'
    raise ValueError('无法识别的文件格式（支持 .docx / .doc / .wps 文档）')


def _export_binary(docx_path, dst_path, log=None):
    """将排版后的 .docx 回写为二进制格式（.doc / .wps），保持与原文档相同的格式。

    Windows 走 WPS/Word COM（另存为 Word97 二进制，兼容 .wps 扩展）；
    Linux 走 LibreOffice --convert-to doc 后按需改扩展名。
    """
    if sys.platform.startswith('win'):
        import win32com.client
        app = None
        for progid in ('KWPS.Application', 'Word.Application', 'WPS.Application'):
            try:
                app = win32com.client.DispatchEx(progid)
                break
            except Exception:  # noqa: BLE001
                continue
        if app is None:
            raise RuntimeError('未检测到 WPS/Word')
        try:
            app.Visible = False
            doc = app.Documents.Open(os.path.abspath(docx_path), ReadOnly=True)
            try:
                doc.SaveAs2(os.path.abspath(dst_path), 0)     # 0 = Word97 二进制
            except Exception:  # noqa: BLE001
                doc.SaveAs(os.path.abspath(dst_path), 0)
            doc.Close(False)
        finally:
            try:
                app.Quit()
            except Exception:  # noqa: BLE001
                pass
        if log:
            log(f'  · 已按原格式回写：{os.path.basename(dst_path)}')
        return dst_path
    # Linux：LibreOffice 转 .doc，再按目标扩展名落盘（.wps 与 .doc 同为二进制容器）
    import subprocess
    import shutil
    import tempfile
    outdir = tempfile.mkdtemp(prefix='clpb_exp_')
    profile = outdir + '/loprofile'
    for cmd in ('soffice', 'libreoffice'):
        exe = shutil.which(cmd)
        if not exe:
            continue
        try:
            r = subprocess.run(
                [exe, '--headless', '--norestore', '--nologo', '--nodefault',
                 f'-env:UserInstallation=file://{profile}',
                 '--convert-to', 'doc', '--outdir', outdir,
                 os.path.abspath(docx_path)],
                capture_output=True, timeout=180)
        except (subprocess.TimeoutExpired, OSError):
            continue
        base = os.path.splitext(os.path.basename(docx_path))[0] + '.doc'
        out = os.path.join(outdir, base)
        if r.returncode == 0 and os.path.exists(out):
            shutil.move(out, dst_path)
            shutil.rmtree(outdir, ignore_errors=True)
            if log:
                log(f'  · 已按原格式回写：{os.path.basename(dst_path)}')
            return dst_path
    shutil.rmtree(outdir, ignore_errors=True)
    raise RuntimeError('未检测到 LibreOffice/WPS')



# ────────────────────────── 文件级入口 ──────────────────────────
def resolve_fonts(title_font=None, fs_font=None, kai_font=None, hei_font=None):
    return {
        'title': title_font or '方正小标宋简体',
        'fs': fs_font or '仿宋_GB2312',
        'kai': kai_font or '楷体_GB2312',
        'hei': hei_font or '黑体',
    }


def default_output(src):
    """默认输出：保持与原文档相同的扩展名（.docx / .doc / .wps）"""
    base, ext = os.path.splitext(src)
    if ext.lower() in ('.docx', '.doc', '.wps'):
        return base + '_排版' + ext.lower()
    return base + '_排版.docx'


def format_file(src, dst=None, *, preset=DEFAULT_PRESET, title_font=None,
                fs_font=None, kai_font=None, hei_font=None, forwarding=False,
                page_numbers=True, remove_blank=True, log=None):
    """排版一个文档文件（.docx / .doc / .wps）；默认输出「原名_排版.docx»。

    preset：'gb9704'（GB/T 9704—2012，默认）或 'swb'（市委办公室材料排版基本格式）。
    .doc/.wps 旧格式自动调用本机 WPS/Word（Windows）/LibreOffice（Linux）转换后处理。
    返回 (输出路径, 统计)。
    """
    if preset not in PRESETS:
        raise ValueError(f'未知格式标准：{preset}（可选：{" / ".join(PRESETS)}）')
    fonts = resolve_fonts(title_font, fs_font, kai_font, hei_font)
    work_src, tmp, kind = _prepare_source(src, log)
    try:
        doc = Document(work_src)
        stats = format_document(doc, fonts, preset=preset, forwarding=forwarding,
                                page_numbers=page_numbers, remove_blank=remove_blank, log=log)
        dst = dst or default_output(src)
        out_ext = os.path.splitext(dst)[1].lower()
        # 保持原文件格式：.doc 输出 .doc；二进制 .wps 输出 .wps（ZIP 容器 .wps 直接存 OOXML 内容）
        need_binary = out_ext == '.doc' or (out_ext == '.wps' and kind == 'ole')
        if need_binary:
            import tempfile
            fd, tdocx = tempfile.mkstemp(suffix='.docx')
            os.close(fd)
            os.remove(tdocx)
            try:
                doc.save(tdocx)
                _export_binary(tdocx, dst, log)
            except Exception as e:  # noqa: BLE001 —— 回写失败降级为 .docx 输出
                if log:
                    log(f'· 按原格式回写未成功（{e}），改为输出 .docx')
                dst = os.path.splitext(dst)[0] + '.docx'
                doc.save(dst)
            finally:
                if os.path.exists(tdocx):
                    os.remove(tdocx)
        else:
            doc.save(dst)
        return dst, stats
    finally:
        if tmp:
            try:
                os.remove(tmp)
            except OSError:
                pass
