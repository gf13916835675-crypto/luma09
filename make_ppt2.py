# -*- coding: utf-8 -*-
"""生成"学科布局"一页 PPT(交大红蓝配色,优化排版)"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

SJTU_RED = RGBColor(0xA7, 0x20, 0x38)
SJTU_BLUE = RGBColor(0x00, 0x45, 0x7C)
GOLD = RGBColor(0xC9, 0x9B, 0x2C)
LIGHT_BG = RGBColor(0xFA, 0xF7, 0xF7)
LIGHT_GREY = RGBColor(0xF2, 0xF2, 0xF2)
MID_GREY = RGBColor(0x9B, 0x9B, 0x9B)
DARK_TEXT = RGBColor(0x33, 0x33, 0x33)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
FONT = "微软雅黑"

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
slide = prs.slides.add_slide(prs.slide_layouts[6])
SW = prs.slide_width
SH = prs.slide_height


def set_font(run, size, color, bold=False):
    f = run.font
    f.name = FONT
    f.size = Pt(size)
    f.color.rgb = color
    f.bold = bold
    rPr = run._r.get_or_add_rPr()
    ea = rPr.find(qn('a:ea'))
    if ea is None:
        ea = rPr.makeelement(qn('a:ea'), {})
        rPr.append(ea)
    ea.set('typeface', FONT)


def add_box(x, y, w, h, fill=None, line=None, line_w=None, shape=MSO_SHAPE.RECTANGLE):
    sp = slide.shapes.add_shape(shape, x, y, w, h)
    sp.shadow.inherit = False
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid()
        sp.fill.fore_color.rgb = fill
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line
        sp.line.width = Pt(line_w or 1)
    return sp


def add_par(tf, runs, align=PP_ALIGN.CENTER, first=False, space_before=None):
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.alignment = align
    if space_before is not None:
        p.space_before = Pt(space_before)
    for text, size, color, bold in runs:
        r = p.add_run()
        r.text = text
        set_font(r, size, color, bold)
    return p

# ---------- 顶部红色标题条 ----------
add_box(0, 0, SW, Inches(1.0), fill=SJTU_RED)
add_box(0, Inches(1.0), SW, Inches(0.06), fill=SJTU_BLUE)
tb = slide.shapes.add_textbox(Inches(0.6), 0, Inches(9), Inches(1.0))
tf = tb.text_frame
tf.vertical_anchor = MSO_ANCHOR.MIDDLE
add_par(tf, [("学科布局", 30, WHITE, True)], align=PP_ALIGN.LEFT, first=True)
tb2 = slide.shapes.add_textbox(Inches(10.2), 0, Inches(2.6), Inches(1.0))
tf2 = tb2.text_frame
tf2.vertical_anchor = MSO_ANCHOR.MIDDLE
add_par(tf2, [("SJTU", 14, RGBColor(0xF0, 0xC0, 0xC8), False)], align=PP_ALIGN.RIGHT, first=True)

# ---------- 上半部分:两个学科卡片 ----------
disc_y = Inches(1.4)
disc_h = Inches(2.7)
disc_w = Inches(5.9)
gap = Inches(0.35)
start_x = (SW - disc_w * 2 - gap) / 2

# 左:计算机科学与技术
x = start_x
card = add_box(x, disc_y, disc_w, disc_h, fill=LIGHT_BG, line=SJTU_RED, line_w=1.5,
               shape=MSO_SHAPE.ROUNDED_RECTANGLE)
card.adjustments[0] = 0.05
add_box(x, disc_y + Inches(0.05), Inches(0.12), disc_h - Inches(0.1), fill=SJTU_RED)
tb = slide.shapes.add_textbox(x + Inches(0.35), disc_y + Inches(0.22), disc_w - Inches(0.6), disc_h - Inches(0.44))
tf = tb.text_frame
tf.word_wrap = True
add_par(tf, [("计算机科学与技术", 24, SJTU_RED, True)], first=True)
add_par(tf, [("教育部第五轮学科评估  ", 17, DARK_TEXT, True), ("A+", 34, SJTU_RED, True)], space_before=8)
add_par(tf, [("首批国家一级重点学科", 14, DARK_TEXT, False)], space_before=10)
add_par(tf, [("首批国家一流学科", 14, DARK_TEXT, False)], space_before=4)

# 右:网络空间安全
x = start_x + disc_w + gap
card = add_box(x, disc_y, disc_w, disc_h, fill=LIGHT_BG, line=SJTU_BLUE, line_w=1.5,
               shape=MSO_SHAPE.ROUNDED_RECTANGLE)
card.adjustments[0] = 0.05
add_box(x, disc_y + Inches(0.05), Inches(0.12), disc_h - Inches(0.1), fill=SJTU_BLUE)
tb = slide.shapes.add_textbox(x + Inches(0.35), disc_y + Inches(0.22), disc_w - Inches(0.6), disc_h - Inches(0.44))
tf = tb.text_frame
tf.word_wrap = True
add_par(tf, [("网络空间安全", 24, SJTU_BLUE, True)], first=True)
add_par(tf, [("国家一级学科", 26, SJTU_BLUE, True)], space_before=8)
add_par(tf, [("两次入选“一流网络安全学院", 14, DARK_TEXT, False)], space_before=10)
add_par(tf, [("建设示范项目”", 14, DARK_TEXT, False)], space_before=4)

# ---------- 下半部分:四个排名卡片 ----------
rank_y = Inches(4.45)
rank_h = Inches(2.55)
rank_w = Inches(2.9)
rgap = Inches(0.22)
start_x = (SW - rank_w * 4 - rgap * 3) / 2
badge_h = Inches(0.55)

ranks = [
    ("QS", SJTU_BLUE, [("全球第", 15, "20", 30), ("全国第", 15, "3", 30)]),
    ("U.S. News", SJTU_RED, [("全球第", 15, "6", 30), ("全国第", 15, "4", 30)]),
    ("CSRankings", SJTU_BLUE, [("全球第", 15, "1", 30), ("全国第", 15, "1", 30)]),
    ("ESI", SJTU_RED, [("全球第", 15, "16", 30), ("全球排名 前", 15, "1‰", 30)]),
]
for i, (name, color, rows) in enumerate(ranks):
    x = start_x + i * (rank_w + rgap)
    card = add_box(x, rank_y, rank_w, rank_h, fill=WHITE, line=color, line_w=1.25,
                   shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    card.adjustments[0] = 0.06
    badge = add_box(x, rank_y, rank_w, badge_h, fill=color)
    tbb = slide.shapes.add_textbox(x, rank_y, rank_w, badge_h)
    tfb = tbb.text_frame
    tfb.vertical_anchor = MSO_ANCHOR.MIDDLE
    add_par(tfb, [(name, 17, WHITE, True)], first=True)

    tb = slide.shapes.add_textbox(x + Inches(0.15), rank_y + badge_h + Inches(0.08),
                                  rank_w - Inches(0.3), rank_h - badge_h - Inches(0.2))
    tf = tb.text_frame
    tf.word_wrap = True
    add_par(tf, [("2025 年学科排名", 12, MID_GREY, False)], first=True)
    for j, (pre, pre_sz, num, num_sz) in enumerate(rows):
        add_par(tf, [(pre, pre_sz, DARK_TEXT, False), (num, num_sz, color, True)],
                space_before=8 if j == 0 else 2)

prs.save("/workspace/学科布局.pptx")
print("saved")
