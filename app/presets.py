"""预设定义：内置排版预设与默认参数结构。

预设为一个可 JSON 序列化的 dict，结构见 DEFAULT_PRESET。
用户自定义预设保存在 presets_user/ 目录下的 .json 文件中。
"""

import copy

DEFAULT_PRESET = {
    "name": "自定义",
    "page": {
        "size": "A4",                # A4 | B5 | Letter | 16K
        "orientation": "portrait",   # portrait | landscape
        "margin_top_cm": 2.54,
        "margin_bottom_cm": 2.54,
        "margin_left_cm": 3.17,
        "margin_right_cm": 3.17,
        "header_distance_cm": 1.5,   # 页眉距边界
        "footer_distance_cm": 1.75,  # 页脚距边界
        # 行距始终由 body.line_spacing 决定；引擎会禁用文档网格，保证行距真实生效
    },
    "body": {
        "font_east": "宋体",
        "font_west": "Times New Roman",
        "size_pt": 12.0,
        "align": "justify",          # left | center | right | justify
        "line_spacing_type": "multiple",  # multiple | exact | minimum
        "line_spacing_value": 1.5,   # multiple=倍数；exact/minimum=磅
        "space_before_pt": 0.0,
        "space_after_pt": 0.0,
        "first_line_indent_chars": 2.0,
        "format_tables": False,      # 是否同时格式化表格内的文字
    },
    "headings": [
        # line_spacing 继承正文设置
        {"level": 1, "font_east": "黑体", "font_west": "Times New Roman", "size_pt": 16.0,
         "bold": False, "align": "center", "space_before_pt": 12.0, "space_after_pt": 12.0,
         "first_line_indent_chars": 0.0},
        {"level": 2, "font_east": "黑体", "font_west": "Times New Roman", "size_pt": 14.0,
         "bold": False, "align": "left", "space_before_pt": 6.0, "space_after_pt": 6.0,
         "first_line_indent_chars": 0.0},
        {"level": 3, "font_east": "黑体", "font_west": "Times New Roman", "size_pt": 12.0,
         "bold": True, "align": "left", "space_before_pt": 6.0, "space_after_pt": 6.0,
         "first_line_indent_chars": 0.0},
        {"level": 4, "font_east": "宋体", "font_west": "Times New Roman", "size_pt": 12.0,
         "bold": True, "align": "left", "space_before_pt": 6.0, "space_after_pt": 6.0,
         "first_line_indent_chars": 0.0},
    ],
    "page_number": {
        "enabled": True,
        "position": "footer_center",  # footer_center | footer_right | footer_left | header_center | header_right
        "format": "dash",             # plain | dash | cn | cn_total | en
        "font_east": "宋体",
        "font_west": "Times New Roman",
        "size_pt": 9.0,
        "hide_on_first_page": False,
        "start_at": 1,
    },
    "header": {"enabled": False, "text": ""},
    "footer": {"enabled": False, "text": ""},
}


def merge_preset(data: dict) -> dict:
    """把用户 JSON 深合并到默认结构上，保证字段完整、可向后兼容。"""
    base = copy.deepcopy(DEFAULT_PRESET)

    def merge(dst, src):
        for k, v in src.items():
            if isinstance(v, dict) and isinstance(dst.get(k), dict):
                merge(dst[k], v)
            else:
                dst[k] = v

    if isinstance(data, dict):
        merge(base, data)
        page = base["page"]
        # 清理旧版预设的文档网格字段：行距一律由“正文”行距决定，网格已移除
        for legacy in ("grid_mode", "grid_lines", "grid_chars", "disable_doc_grid"):
            page.pop(legacy, None)
        # 旧版预设只有 3 级标题：补齐第 4 级
        levels = {h.get("level") for h in base.get("headings", [])}
        for h in DEFAULT_PRESET["headings"]:
            if h["level"] not in levels:
                base["headings"].append(copy.deepcopy(h))
        base["headings"].sort(key=lambda h: h["level"])
    return base


# ---------------------------------------------------------------------------
# 内置预设
# ---------------------------------------------------------------------------

BUILTIN_PRESETS = [
    {
        "name": "标准办公文档",
        "page": {"size": "A4", "orientation": "portrait",
                 "margin_top_cm": 2.54, "margin_bottom_cm": 2.54,
                 "margin_left_cm": 3.17, "margin_right_cm": 3.17},
        "body": {"font_east": "宋体", "font_west": "Times New Roman", "size_pt": 12.0,
                 "align": "justify", "line_spacing_type": "multiple", "line_spacing_value": 1.5,
                 "space_before_pt": 0.0, "space_after_pt": 0.0,
                 "first_line_indent_chars": 2.0, "format_tables": False},
        "headings": [
            {"level": 1, "font_east": "黑体", "font_west": "Times New Roman", "size_pt": 16.0,
             "bold": False, "align": "center", "space_before_pt": 12.0, "space_after_pt": 12.0,
             "first_line_indent_chars": 0.0},
            {"level": 2, "font_east": "黑体", "font_west": "Times New Roman", "size_pt": 14.0,
             "bold": False, "align": "left", "space_before_pt": 6.0, "space_after_pt": 6.0,
             "first_line_indent_chars": 0.0},
            {"level": 3, "font_east": "黑体", "font_west": "Times New Roman", "size_pt": 12.0,
             "bold": True, "align": "left", "space_before_pt": 6.0, "space_after_pt": 6.0,
             "first_line_indent_chars": 0.0},
            {"level": 4, "font_east": "宋体", "font_west": "Times New Roman", "size_pt": 12.0,
             "bold": True, "align": "left", "space_before_pt": 6.0, "space_after_pt": 6.0,
             "first_line_indent_chars": 0.0},
        ],
        "page_number": {"enabled": True, "position": "footer_center", "format": "dash",
                        "font_east": "宋体", "font_west": "Times New Roman", "size_pt": 9.0,
                        "hide_on_first_page": False, "start_at": 1},
        "header": {"enabled": False, "text": ""},
        "footer": {"enabled": False, "text": ""},
    },
    {
        # GB/T 9704-2012《党政机关公文格式》：
        # 5.1 A4；5.2 版心 156 mm × 225 mm，天头 37 mm、订口 28 mm（→ 下 35 mm、右 26 mm）；
        # 5.3 一般用 3 号仿宋体，每面 22 行、每行 28 字，撑满版心
        #     （→ 正文固定行距 28.9 磅：版心 225 mm ÷ 22 行 ≈ 28.95 磅，取 28.9 保证每页 22 行）；
        # 5.5 页码 4 号半角宋体阿拉伯数字，左右各一条一字线，一字线上距版心下边缘 7 mm，
        #     单页码居右空一字、双页码居左空一字；
        # 7.3.1 标题 2 号小标宋体居中，标题下空一行；
        # 7.3.3 正文每段左空二字；结构层次序数：第一层黑体、第二层楷体、第三 / 四层仿宋。
        # 本预设标题级别：1 级 = 公文标题（Word“标题/标题 1”样式），2 级 = “一、”，3 级 = “（一）”，4 级 = “1.”
        "name": "党政机关公文（GB/T 9704 参照）",
        "page": {"size": "A4", "orientation": "portrait",
                 "margin_top_cm": 3.7, "margin_bottom_cm": 3.5,
                 "margin_left_cm": 2.8, "margin_right_cm": 2.6,
                 "header_distance_cm": 1.5, "footer_distance_cm": 2.5},
        # 每页 22 行由正文固定行距实现：版心 225 mm ÷ 22 行 ≈ 28.95 磅，取 28.9 保证放得下 22 行
        "body": {"font_east": "仿宋_GB2312", "font_west": "Times New Roman", "size_pt": 16.0,
                 "align": "justify", "line_spacing_type": "exact", "line_spacing_value": 28.9,
                 "space_before_pt": 0.0, "space_after_pt": 0.0,
                 "first_line_indent_chars": 2.0, "format_tables": False},
        "headings": [
            # 公文标题：2 号小标宋，居中，标题下空一行（28.9 磅 ≈ 一行行距）
            {"level": 1, "font_east": "方正小标宋简体", "font_west": "Times New Roman", "size_pt": 22.0,
             "bold": False, "align": "center", "space_before_pt": 0.0, "space_after_pt": 28.9,
             "first_line_indent_chars": 0.0},
            # 第一层“一、”：黑体
            {"level": 2, "font_east": "黑体", "font_west": "Times New Roman", "size_pt": 16.0,
             "bold": False, "align": "justify", "space_before_pt": 0.0, "space_after_pt": 0.0,
             "first_line_indent_chars": 2.0},
            # 第二层“（一）”：楷体_GB2312
            {"level": 3, "font_east": "楷体_GB2312", "font_west": "Times New Roman", "size_pt": 16.0,
             "bold": False, "align": "justify", "space_before_pt": 0.0, "space_after_pt": 0.0,
             "first_line_indent_chars": 2.0},
            # 第三层“1.”：仿宋_GB2312
            {"level": 4, "font_east": "仿宋_GB2312", "font_west": "Times New Roman", "size_pt": 16.0,
             "bold": False, "align": "justify", "space_before_pt": 0.0, "space_after_pt": 0.0,
             "first_line_indent_chars": 2.0},
        ],
        # 4 号宋体（数字也用宋体，半角）；单页居右、双页居左；一字线
        "page_number": {"enabled": True, "position": "footer_outside", "format": "gb_dash",
                        "font_east": "宋体", "font_west": "宋体", "size_pt": 14.0,
                        "hide_on_first_page": False, "start_at": 1},
        "header": {"enabled": False, "text": ""},
        "footer": {"enabled": False, "text": ""},
    },
    {
        "name": "毕业论文（通用）",
        "page": {"size": "A4", "orientation": "portrait",
                 "margin_top_cm": 3.0, "margin_bottom_cm": 2.5,
                 "margin_left_cm": 3.0, "margin_right_cm": 2.5},
        "body": {"font_east": "宋体", "font_west": "Times New Roman", "size_pt": 12.0,
                 "align": "justify", "line_spacing_type": "multiple", "line_spacing_value": 1.5,
                 "space_before_pt": 0.0, "space_after_pt": 0.0,
                 "first_line_indent_chars": 2.0, "format_tables": False},
        "headings": [
            {"level": 1, "font_east": "黑体", "font_west": "Times New Roman", "size_pt": 16.0,
             "bold": False, "align": "center", "space_before_pt": 18.0, "space_after_pt": 12.0,
             "first_line_indent_chars": 0.0},
            {"level": 2, "font_east": "黑体", "font_west": "Times New Roman", "size_pt": 14.0,
             "bold": False, "align": "left", "space_before_pt": 12.0, "space_after_pt": 6.0,
             "first_line_indent_chars": 0.0},
            {"level": 3, "font_east": "黑体", "font_west": "Times New Roman", "size_pt": 12.0,
             "bold": True, "align": "left", "space_before_pt": 6.0, "space_after_pt": 6.0,
             "first_line_indent_chars": 0.0},
            {"level": 4, "font_east": "宋体", "font_west": "Times New Roman", "size_pt": 12.0,
             "bold": True, "align": "left", "space_before_pt": 6.0, "space_after_pt": 6.0,
             "first_line_indent_chars": 0.0},
        ],
        "page_number": {"enabled": True, "position": "footer_center", "format": "plain",
                        "font_east": "宋体", "font_west": "Times New Roman", "size_pt": 10.5,
                        "hide_on_first_page": False, "start_at": 1},
        "header": {"enabled": False, "text": ""},
        "footer": {"enabled": False, "text": ""},
    },
    {
        "name": "简洁现代报告",
        "page": {"size": "A4", "orientation": "portrait",
                 "margin_top_cm": 2.54, "margin_bottom_cm": 2.54,
                 "margin_left_cm": 2.8, "margin_right_cm": 2.8},
        "body": {"font_east": "微软雅黑", "font_west": "Calibri", "size_pt": 10.5,
                 "align": "justify", "line_spacing_type": "multiple", "line_spacing_value": 1.5,
                 "space_before_pt": 0.0, "space_after_pt": 6.0,
                 "first_line_indent_chars": 0.0, "format_tables": False},
        "headings": [
            {"level": 1, "font_east": "微软雅黑", "font_west": "Calibri", "size_pt": 18.0,
             "bold": True, "align": "left", "space_before_pt": 18.0, "space_after_pt": 8.0,
             "first_line_indent_chars": 0.0},
            {"level": 2, "font_east": "微软雅黑", "font_west": "Calibri", "size_pt": 14.0,
             "bold": True, "align": "left", "space_before_pt": 12.0, "space_after_pt": 6.0,
             "first_line_indent_chars": 0.0},
            {"level": 3, "font_east": "微软雅黑", "font_west": "Calibri", "size_pt": 12.0,
             "bold": True, "align": "left", "space_before_pt": 6.0, "space_after_pt": 6.0,
             "first_line_indent_chars": 0.0},
            {"level": 4, "font_east": "微软雅黑", "font_west": "Calibri", "size_pt": 10.5,
             "bold": True, "align": "left", "space_before_pt": 6.0, "space_after_pt": 3.0,
             "first_line_indent_chars": 0.0},
        ],
        "page_number": {"enabled": True, "position": "footer_right", "format": "plain",
                        "font_east": "微软雅黑", "font_west": "Calibri", "size_pt": 9.0,
                        "hide_on_first_page": False, "start_at": 1},
        "header": {"enabled": False, "text": ""},
        "footer": {"enabled": False, "text": ""},
    },
]
