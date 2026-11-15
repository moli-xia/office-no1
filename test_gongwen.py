# -*- coding: utf-8 -*-
"""对 gongwen.py 的自动化验证：先造样例，再排版，逐项断言输出 XML。"""
import os
import sys
import zipfile
sys.stdout.reconfigure(encoding='utf-8')

import make_sample
import gongwen

SRC = '样例_原始.docx'
DST = '样例_排版.docx'
make_sample.make(SRC)

fails = []


def check(name, cond, detail=''):
    tag = 'PASS' if cond else 'FAIL'
    print(f'[{tag}] {name}' + (f'  （{detail}）' if detail and not cond else ''))
    if not cond:
        fails.append(name)


from docx import Document  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402

# ── 一、市委办公室格式（preset='swb'）──
dst, stats = gongwen.format_file(
    SRC, DST, preset='swb', title_font='方正小标宋简体', fs_font='仿宋_GB2312',
    kai_font='楷体_GB2312', log=print)
doc = Document(DST)

# ── 一、幅面 二、页边距（市委办 3.9/3.3/2.7/2.7）──
sec = doc.sections[0]
check('A4 幅面 210×297mm',
      abs(sec.page_width.cm - 21.0) < 0.01 and abs(sec.page_height.cm - 29.7) < 0.01,
      f'{sec.page_width.cm}×{sec.page_height.cm}')
check('页边距 上3.9', round(sec.top_margin.cm, 2) == 3.9)
check('页边距 下3.3', round(sec.bottom_margin.cm, 2) == 3.3)
check('页边距 左2.7', round(sec.left_margin.cm, 2) == 2.7)
check('页边距 右2.7', round(sec.right_margin.cm, 2) == 2.7)

# ── 三、网格：22行×28字 ──
dg = sec._sectPr.find(qn('w:docGrid'))
check('文档网格存在', dg is not None)
check('每面22行（行距=网格560）', dg.get(qn('w:linePitch')) == '560',
      dg.get(qn('w:linePitch')))
check('每行28字（charSpace=-849）', dg.get(qn('w:charSpace')) == '-849',
      dg.get(qn('w:charSpace')))
check('网格类型 linesAndChars', dg.get(qn('w:type')) == 'linesAndChars')

# ── 六、页码 ──
settings = doc.settings.element
check('奇偶页设置 evenAndOddHeaders', settings.find(qn('w:evenAndOddHeaders')) is not None)


def footer_info(footer):
    p = footer.paragraphs[0]
    pPr = p._p.find(qn('w:pPr'))
    jc = pPr.find(qn('w:jc')).get(qn('w:val')) if pPr.find(qn('w:jc')) is not None else None
    ind = pPr.find(qn('w:ind'))
    ind_d = {k.split('}')[1]: ind.get(k) for k in ind.attrib} if ind is not None else {}
    texts = [t.text for t in p._p.findall('.//' + qn('w:t')) if t.text]
    instr = [i.text for i in p._p.findall('.//' + qn('w:instrText'))]
    szs = {r.get(qn('w:val')) for r in p._p.findall('.//' + qn('w:sz'))}
    easts = {r.get(qn('w:eastAsia')) for r in p._p.findall('.//' + qn('w:rFonts'))}
    return jc, ''.join(texts), instr, szs, easts, ind_d


fodd = footer_info(sec.footer)
feven = footer_info(sec.even_page_footer)
check('单页码居右', fodd[0] == 'right', str(fodd[0]))
check('双页码居左', feven[0] == 'left', str(feven[0]))
check('页码含 PAGE 域', any('PAGE' in (x or '') for x in fodd[2]))
check('页码一字线「— 1 —」', '—' in fodd[1] and '1' in fodd[1], fodd[1])
check('单页码右空一字（右缩进一字）', fodd[5].get('rightChars') == '100',
      str(fodd[5]))
check('双页码左空一字（左缩进一字）', feven[5].get('leftChars') == '100',
      str(feven[5]))
check('页码四号（sz=28 半磅）', fodd[3] == {'28'}, str(fodd[3]))
check('页码宋体', fodd[4] == {'宋体'}, str(fodd[4]))

# ── 四、标题 ──
paras = doc.paragraphs
texts = [p.text.strip() for p in paras]
i_title = next(i for i, t in enumerate(texts) if t)
tp = paras[i_title]
pPr = tp._p.find(qn('w:pPr'))
check('标题居中', pPr.find(qn('w:jc')).get(qn('w:val')) == 'center')
sp = pPr.find(qn('w:spacing'))
check('标题固定值33磅', sp.get(qn('w:line')) == '660' and sp.get(qn('w:lineRule')) == 'exact',
      f'{sp.get(qn("w:line"))}/{sp.get(qn("w:lineRule"))}')
check('标题不对齐网格', pPr.find(qn('w:snapToGrid')).get(qn('w:val')) == '0')
rf = tp._p.find('.//' + qn('w:rFonts'))
check('标题二号小标宋', rf.get(qn('w:eastAsia')) == '方正小标宋简体'
      and tp._p.find('.//' + qn('w:sz')).get(qn('w:val')) == '44',
      f'{rf.get(qn("w:eastAsia"))}')

# 日期段（标题区）
i_date = next(i for i, t in enumerate(texts) if '（2026年11月9日）' in t and i < 5)
dp = paras[i_date]
rf = dp._p.find('.//' + qn('w:rFonts'))
check('成文日期三号楷体', rf.get(qn('w:eastAsia')) == '楷体_GB2312'
      and dp._p.find('.//' + qn('w:sz')).get(qn('w:val')) == '32')
check('日期居中', dp._p.find(qn('w:pPr')).find(qn('w:jc')).get(qn('w:val')) == 'center')

# 标题区下自动空一行（原文无空行也应插入）
check('标题（日期）下自动空一行',
      not paras[i_date + 1].text.strip()
      and paras[i_date + 2].text.strip().endswith('：'))

# ── 主送机关 ──
i_zs = next(i for i, t in enumerate(texts) if t.endswith('：'))
zp = paras[i_zs]
ind = zp._p.find(qn('w:pPr')).find(qn('w:ind'))
check('主送顶格无缩进', ind is None)

# ── 五、正文 ──
i_body = next(i for i, t in enumerate(texts) if t.startswith('城市是经济发展'))
bp = paras[i_body]
pPr = bp._p.find(qn('w:pPr'))
sp = pPr.find(qn('w:spacing'))
check('正文固定值28磅', sp.get(qn('w:line')) == '560' and sp.get(qn('w:lineRule')) == 'exact',
      f'{sp.get(qn("w:line"))}/{sp.get(qn("w:lineRule"))}')
wc = pPr.find(qn('w:widowControl'))
check('孤行控制已关闭（保证每面恒22行）', wc is not None and wc.get(qn('w:val')) == '0')
spattrs = {k.split('}')[1]: v for k, v in sp.attrib.items()}
check('段前段后均为0且无行单位段距残留',
      spattrs.get('before') == '0' and spattrs.get('after') == '0'
      and 'beforeLines' not in spattrs and 'afterLines' not in spattrs
      and spattrs.get('beforeAutospacing', '0') == '0', str(spattrs))
ind = pPr.find(qn('w:ind'))
check('段首空2字（firstLineChars=200）', ind is not None
      and ind.get(qn('w:firstLineChars')) == '200')
check('正文两端对齐', pPr.find(qn('w:jc')).get(qn('w:val')) == 'both')
check('正文对齐网格（22行撑满版心）', pPr.find(qn('w:snapToGrid')).get(qn('w:val')) == '1')
rf = bp._p.find('.//' + qn('w:rFonts'))
check('正文三号仿宋_GB2312', rf.get(qn('w:eastAsia')) == '仿宋_GB2312'
      and bp._p.find('.//' + qn('w:sz')).get(qn('w:val')) == '32')
check('西文/数字 TNR', rf.get(qn('w:ascii')) == 'Times New Roman')
col = bp._p.find('.//' + qn('w:color'))
check('文字黑色', col is not None and col.get(qn('w:val')) == '000000')


def para_of(prefix):
    return next(p for p in paras if p.text.strip().startswith(prefix))


def bold_off(p):
    b = p._p.find('.//' + qn('w:b'))
    return b is None or b.get(qn('w:val')) in ('0', 'false')


# 层次序数字体
p1 = para_of('一、')
check('一层黑体不加粗', p1._p.find('.//' + qn('w:rFonts')).get(qn('w:eastAsia')) == '黑体'
      and bold_off(p1))
p2 = para_of('（一）')
check('二层楷体加粗（独立成段）', p2._p.find('.//' + qn('w:rFonts')).get(qn('w:eastAsia')) == '楷体_GB2312'
      and p2._p.find('.//' + qn('w:b')) is not None)
p3 = para_of('1. 加快')
check('三层仿宋加粗', p3._p.find('.//' + qn('w:rFonts')).get(qn('w:eastAsia')) == '仿宋_GB2312'
      and p3._p.find('.//' + qn('w:b')) is not None)
p4 = para_of('（1）完善')
check('四层仿宋加粗', p4._p.find('.//' + qn('w:rFonts')).get(qn('w:eastAsia')) == '仿宋_GB2312'
      and p4._p.find('.//' + qn('w:b')) is not None)
p35 = para_of('城市是经济')
check('“3.5%”未被误判为三层序数（仍为正文仿宋不加粗）',
      p35._p.find('.//' + qn('w:rFonts')).get(qn('w:eastAsia')) == '仿宋_GB2312'
      and bold_off(p35))

# 小标题与正文同段：序数部分楷体加粗，接排正文仿宋不加粗
pi = para_of('（二）发展目标。')
runs = pi._p.findall('.//' + qn('w:r'))
rf0 = runs[0].find(qn('w:rPr')).find(qn('w:rFonts'))
b0 = runs[0].find(qn('w:rPr')).find(qn('w:b'))
rf_last = runs[-1].find(qn('w:rPr')).find(qn('w:rFonts'))
b_last = runs[-1].find(qn('w:rPr')).find(qn('w:b'))
t_last = ''.join(x.text or '' for x in runs[-1].findall(qn('w:t')))
check('接排：序数段楷体_GB2312加粗', rf0.get(qn('w:eastAsia')) == '楷体_GB2312' and b0 is not None)
check('接排：后续正文仿宋不加粗', rf_last.get(qn('w:eastAsia')) == '仿宋_GB2312'
      and (b_last is None or b_last.get(qn('w:val')) in ('0', 'false'))
      and t_last.startswith('到2026年'), t_last[:12])

# 空行删除
check('空行已删除', stats['blanks_removed'] > 0, str(stats))
# 落款：原文"市政府办公室　　2026年11月9日"同一行 → 拆分为两行
check('同行落款已拆分为两行', stats['sig_split'] == 1, str(stats))
# 落款与正文之间恰好空3行
from docx.text.paragraph import Paragraph as _Para
sig_el = next(p for p in paras if p.text.strip() == '市政府办公室')._p
blanks_before = 0
cur = sig_el.getprevious()
while cur is not None and cur.tag == qn('w:p'):
    po = _Para(cur, None)
    if po.text.strip():
        break
    blanks_before += 1
    cur = cur.getprevious()
check('落款与正文之间空3行', blanks_before == 3, str(blanks_before))
sig_p = next(p for p in paras if p.text.strip() == '市政府办公室')
date_p = next(p for p in paras if p.text.strip() == '2026年11月9日')
sig_pPr = sig_p._p.find(qn('w:pPr'))
check('落款署名右对齐',
      sig_pPr.find(qn('w:jc')).get(qn('w:val')) == 'right')
p_ld = date_p
pPr = p_ld._p.find(qn('w:pPr'))
check('落款日期楷体+右对齐+右空二字',
      p_ld._p.find('.//' + qn('w:rFonts')).get(qn('w:eastAsia')) == '楷体_GB2312'
      and pPr.find(qn('w:jc')).get(qn('w:val')) == 'right'
      and pPr.find(qn('w:ind')).get(qn('w:rightChars')) == '200')

# 数字不折行设置
check('wordWrap=0（数字年份不折行）', pPr.find(qn('w:wordWrap')).get(qn('w:val')) == '0')

# ── 包完整性 ──
with zipfile.ZipFile(DST) as z:
    names = z.namelist()
check('输出包含奇偶页脚部件',
      any('footer' in n for n in names), str([n for n in names if 'footer' in n]))

# ══════════ GB/T 9704—2012（默认格式）断言 ══════════
DST_GB = '样例_排版_GB.docx'
gongwen.format_file(SRC, DST_GB, title_font='方正小标宋简体', fs_font='仿宋_GB2312',
                    kai_font='楷体_GB2312', log=lambda *_: None)
gdoc = Document(DST_GB)
gsec = gdoc.sections[0]
gparas = gdoc.paragraphs

check('GB 边距 上3.7', round(gsec.top_margin.cm, 2) == 3.7)
check('GB 边距 下3.5', round(gsec.bottom_margin.cm, 2) == 3.5)
check('GB 边距 左2.8', round(gsec.left_margin.cm, 2) == 2.8)
check('GB 边距 右2.6', round(gsec.right_margin.cm, 2) == 2.6)
check('GB 版心 156×225mm',
      round((gsec.page_width.cm - gsec.left_margin.cm - gsec.right_margin.cm) * 10) == 156
      and round((gsec.page_height.cm - gsec.top_margin.cm - gsec.bottom_margin.cm) * 10) == 225)
check('GB 页脚距离校准（2.23cm）', round(gsec.footer_distance.cm, 2) == 2.23)
gdg = gsec._sectPr.find(qn('w:docGrid'))
check('GB 网格与市委办一致（22行×28字）',
      gdg.get(qn('w:linePitch')) == '560' and gdg.get(qn('w:charSpace')) == '-849')


def gpara_of(prefix):
    return next(p for p in gparas if p.text.strip().startswith(prefix))


gp2 = gpara_of('（一）')
gb2 = gp2._p.find('.//' + qn('w:b'))
check('GB 二层楷体加粗', gp2._p.find('.//' + qn('w:rFonts')).get(qn('w:eastAsia')) == '楷体_GB2312'
      and gb2 is not None and gb2.get(qn('w:val')) not in ('0', 'false'))
gp3 = gpara_of('1. 加快')
gb3 = gp3._p.find('.//' + qn('w:b'))
check('GB 三层仿宋加粗', gp3._p.find('.//' + qn('w:rFonts')).get(qn('w:eastAsia')) == '仿宋_GB2312'
      and gb3 is not None and gb3.get(qn('w:val')) not in ('0', 'false'))

gdate = [p for p in gparas if p.text.strip() == '2026年11月9日'][-1]
gind = gdate._p.find(qn('w:pPr')).find(qn('w:ind'))
check('GB 落款日期三号仿宋右空四字',
      gdate._p.find('.//' + qn('w:rFonts')).get(qn('w:eastAsia')) == '仿宋_GB2312'
      and gind.get(qn('w:rightChars')) == '400', str(gind.attrib if gind is not None else None))
gsig = gpara_of('市政府办公室')
gsig_ind = gsig._p.find(qn('w:pPr')).find(qn('w:ind'))
gsig_rc = int(gsig_ind.get(qn('w:rightChars')))
# 日期 2026年11月9日 视觉宽 6.5 字，署名 6 字 → 右空 400 + 25 = 425
check('GB 署名以日期为准居中（右空>四字）', 400 < gsig_rc <= 500, str(gsig_rc))

gfodd_jc = gsec.footer.paragraphs[0]._p.find(qn('w:pPr')).find(qn('w:jc')).get(qn('w:val'))
check('GB 页码仍为单页居右', gfodd_jc == 'right')

# ══════════ 双行落款（单位一行、日期一行）══════════
from docx import Document as _Doc
dd = _Doc()
dd.add_paragraph('双行落款测试文')
dd.add_paragraph('正文内容若干。' * 40)
dd.add_paragraph('××市发展和改革委员会')
dd.add_paragraph('2026年12月1日')
dd.save('sig2_src.docx')
gongwen.format_file('sig2_src.docx', 'sig2_out.docx', log=lambda *_: None)
sd = _Doc('sig2_out.docx')
sp = next(p for p in sd.paragraphs if '发展和改革委员会' in p.text)
dp = next(p for p in sd.paragraphs if p.text.strip() == '2026年12月1日')
sp_ok = sp._p.find(qn('w:pPr')).find(qn('w:jc')).get(qn('w:val')) == 'right'
dp_ok = (dp._p.find(qn('w:pPr')).find(qn('w:ind')).get(qn('w:rightChars')) == '400'
         and dp._p.find('.//' + qn('w:rFonts')).get(qn('w:eastAsia')) == '仿宋_GB2312')
check('双行落款：署名右对齐+日期仿宋右空四字(GB默认)', sp_ok and dp_ok)
os.remove('sig2_src.docx')
os.remove('sig2_out.docx')

# ══════════ 输出保持原文件格式 ══════════
import win32com.client as _w32
_app = _w32.Dispatch('KWPS.Application')
_app.Visible = False
_d = _app.Documents.Open(os.path.abspath(SRC), ReadOnly=True)
_d.SaveAs(os.path.abspath('fmt_src.doc'), 0)
_d.Close(False)
_app.Quit()
with open('fmt_src.doc', 'rb') as f:
    check('测试源 .doc 为二进制格式', f.read(4) == b'\xd0\xcf\x11\xe0')
out_doc, _ = gongwen.format_file('fmt_src.doc', None, log=lambda *_: None)
with open(out_doc, 'rb') as f:
    magic = f.read(4)
check('.doc 输入 → .doc 输出（二进制容器）', out_doc.endswith('.doc') and magic == b'\xd0\xcf\x11\xe0',
      f'{out_doc} {magic}')
os.remove('fmt_src.doc')
os.remove(out_doc)

import shutil as _sh
_sh.copy(SRC, 'fmt_src.wps')          # ZIP 容器 .wps（docx 内容改后缀）
out_wps, _ = gongwen.format_file('fmt_src.wps', None, log=lambda *_: None)
with open(out_wps, 'rb') as f:
    magic = f.read(4)
check('.wps(ZIP) 输入 → .wps 输出（OOXML 容器）', out_wps.endswith('.wps') and magic == b'PK\x03\x04',
      f'{out_wps} {magic}')
os.remove('fmt_src.wps')
os.remove(out_wps)

print()
print('=' * 50)
print('样例统计：', stats)
if fails:
    print(f'✘ {len(fails)} 项未通过：', fails)
    sys.exit(1)
print('✔ 全部断言通过')
