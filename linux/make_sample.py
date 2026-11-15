# -*- coding: utf-8 -*-
"""生成一份"格式混乱"的测试样例，用于验证排版工具的归一化效果。"""
import sys
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH


def make(out='样例_原始.docx'):
    doc = Document()

    # 开头空行（模拟红头纸空两行）
    doc.add_paragraph('')

    # 标题：字体/字号/对齐全部错误
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r = p.add_run('在全市城市高质量发展推进会议上的讲话')
    r.font.name = '宋体'
    r.font.size = Pt(18)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)

    # 日期：错误字体错误字号（其后不空行，紧接主送——验证工具自动补插空行）
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r = p.add_run('（2026年11月9日）')
    r.font.name = 'Calibri'
    r.font.size = Pt(12)

    # 主送机关：错误字体，且带了缩进
    p = doc.add_paragraph()
    r = p.add_run('　　各县（市、区）人民政府，市政府各部门、各直属机构：')
    r.font.name = 'Calibri'
    r.font.size = Pt(11)

    body_texts = [
        '现将会议有关事项通知如下。',
        '一、提高站位，切实增强推进城市高质量发展的责任感使命感',
        '城市是经济发展的重要载体。2026年是"十四五"规划收官之年，各地各部门要完整准确全面贯彻新发展理念，'
        '把城市发展放在中国式现代化大局中谋划推进，以钉钉子精神抓好各项任务落实，确保2025—2026年行动计划的'
        '各项目标如期实现，全市地区生产总值增长3.5%以上。',
        '（一）聚焦重点领域精准发力。',
        '围绕先进制造业、现代服务业、数字经济三大方向，实施重点项目123个，完成投资480亿元，'
        '推动产业链供应链优化升级。',
        '（二）发展目标。到2026年，全市市场主体总量达到120万户以上，其中企业占比62%以上，'
        '年均新增市场主体2万户以上。',
        '1. 加快基础设施建设。',
        '统筹推进交通、能源、水利等重大工程，2026年底前实现市域铁路通车里程突破200公里，'
        '新建5G基站3000个以上。',
        '（1）完善前期工作机制，2026年6月底前完成全部项目可行性研究报告审批。',
        '（2）强化要素保障，优先保障重点项目建设用地和资金需求。',
        '1. 深化改革开放。',
        '深化"放管服"改革，持续优化营商环境，激发市场主体活力。',
        '二、强化保障，确保各项任务落地见效',
        '各级领导干部要带头改进作风，深入基层一线调查研究，着力解决群众急难愁盼问题，'
        '以实绩实效取信于民。',
    ]
    for i, t in enumerate(body_texts):
        p = doc.add_paragraph()
        r = p.add_run(t)
        r.font.name = 'Calibri'
        r.font.size = Pt(11)
        if i % 3 == 0:
            r.font.bold = True                    # 乱加粗
        if i % 5 == 0:
            r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)  # 乱设颜色
        if i == 1:                                # 乱设"行"单位段距（段前1行/段后0.5行）
            from docx.oxml.ns import qn as _qn
            from docx.oxml import OxmlElement as _oe
            _sp = _oe('w:spacing')
            _sp.set(_qn('w:beforeLines'), '100')
            _sp.set(_qn('w:afterLines'), '50')
            _sp.set(_qn('w:after'), '240')
            p._p.get_or_add_pPr().append(_sp)
        if i % 4 == 0:
            doc.add_paragraph('')                 # 乱插空行

    doc.add_paragraph('')
    # 落款：单位名称与日期写在同一行（材料中常见写法，工具应拆分为两行）
    p = doc.add_paragraph()
    r = p.add_run('市政府办公室\u3000\u30002026年11月9日')
    r.font.name = 'Calibri'
    r.font.size = Pt(11)

    doc.save(out)
    print('已生成：', out)


if __name__ == '__main__':
    make(sys.argv[1] if len(sys.argv) > 1 else '样例_原始.docx')
