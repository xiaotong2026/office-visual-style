# -*- coding: utf-8 -*-
"""PPT渲染自检demo: 生成2页pptx -> LibreOffice渲染 -> 逐页验收(无溢出/错位才交付)"""
import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'output')
os.makedirs(OUT, exist_ok=True)
NAVY = RGBColor(0x1F, 0x38, 0x64); BLUE = RGBColor(0x2E, 0x6F, 0xB7)
ORANGE = RGBColor(0xE8, 0x87, 0x3A); GRAY = RGBColor(0x59, 0x59, 0x59); WHITE = RGBColor(0xFF, 0xFF, 0xFF)

prs = Presentation()
prs.slide_width = Inches(13.333); prs.slide_height = Inches(7.5)

# 页1 标题
s1 = prs.slides.add_slide(prs.slide_layouts[6])
s1.background.fill.solid(); s1.background.fill.fore_color.rgb = NAVY
accent = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(5.87), Inches(3.5), Inches(1.6), Pt(4))
accent.fill.solid(); accent.fill.fore_color.rgb = ORANGE; accent.line.fill.background()
tb = s1.shapes.add_textbox(Inches(1), Inches(2.5), Inches(11.3), Inches(1.0))
p = tb.text_frame.paragraphs[0]; p.text = '季度经营简报'; p.alignment = PP_ALIGN.CENTER
p.font.size = Pt(40); p.font.bold = True; p.font.color.rgb = WHITE
tb2 = s1.shapes.add_textbox(Inches(1), Inches(3.9), Inches(11.3), Inches(0.6))
p2 = tb2.text_frame.paragraphs[0]; p2.text = 'office-visual-style · 可视化标准演示'
p2.alignment = PP_ALIGN.CENTER; p2.font.size = Pt(16); p2.font.color.rgb = RGBColor(0xBD, 0xD7, 0xEE)

# 页2 内容
s2 = prs.slides.add_slide(prs.slide_layouts[6])
s2.background.fill.solid(); s2.background.fill.fore_color.rgb = WHITE
pb = s2.shapes.add_textbox(Inches(0.6), Inches(0.35), Inches(12), Inches(0.8)).text_frame.paragraphs[0]
pb.text = '核心结论：统一配色与标注，让图表自己讲结论'
pb.font.size = Pt(26); pb.font.bold = True; pb.font.color.rgb = NAVY
accent2 = s2.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(0.45), Inches(0.12), Inches(0.6))
accent2.fill.solid(); accent2.fill.fore_color.rgb = BLUE; accent2.line.fill.background()
pts = ['洞见式标题：说结论，不说“统计”',
       '去图表垃圾：去外框、去竖网格线、去多余刻度',
       '低饱和专业配色：主蓝 #2E6FB7 + 强调橙 #E8873A',
       '数值直接标注，图例置底，眼睛不用来回对轴']
tf = s2.shapes.add_textbox(Inches(0.9), Inches(1.5), Inches(11.5), Inches(3.0)).text_frame
for i, t in enumerate(pts):
    para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
    para.text = '• ' + t; para.font.size = Pt(18); para.font.color.rgb = GRAY; para.space_after = Pt(12)
stat = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.9), Inches(5.6), Inches(11.5), Inches(1.2))
stat.fill.solid(); stat.fill.fore_color.rgb = RGBColor(0xEA, 0xF1, 0xF9); stat.line.fill.background()
stf = stat.text_frame; stf.text = '交付前用 LibreOffice 渲染逐页自检，无溢出/错位/错标才发出'
stf.paragraphs[0].font.size = Pt(16); stf.paragraphs[0].font.color.rgb = NAVY

path = os.path.join(OUT, 'slides_demo.pptx')
import sys as _sys, os as _os
_sys.path.insert(0, _os.path.dirname(_os.path.abspath(__file__)))
from chart_style import brand_ppt
brand_ppt(prs)
prs.save(path); print('生成:', path)
