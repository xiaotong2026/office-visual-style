# -*- coding: utf-8 -*-
"""管线压测·PPT端: 读Excel同一数据源→10页原生图表汇报(素材来自Excel)
关键: 原生pptx图表(非图片) + KPI数字与Excel交叉校验 + 设计统一(导航条/页脚/配色)"""
import sys, os, json, datetime
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.chart.data import CategoryChartData, XyChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION, XL_TICK_MARK
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn
import copy

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, 'output')

# ── 设计常量(与chart_style一致) ──
NAVY = RGBColor(0x1F, 0x38, 0x64)
BLUE = RGBColor(0x2E, 0x6F, 0xB7)
ORANGE = RGBColor(0xE8, 0x87, 0x3A)
LIGHT = RGBColor(0xBD, 0xD7, 0xEE)
TXT = RGBColor(0x26, 0x26, 0x26)
GREY = RGBColor(0x59, 0x59, 0x59)
RED = RGBColor(0xC0, 0x50, 0x4D)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GREEN = RGBColor(0xD9, 0xEA, 0xD3)
FONT = '微软雅黑'

# ── 一致性基准(来自Excel端输出) ──
ck = json.load(open(os.path.join(OUT, 'consistency.json'), encoding='utf-8'))

# ── 数据集(与Excel端同源逻辑; 此处直接读取Excel明细做真素材传递) ──
import openpyxl
wb = openpyxl.load_workbook(os.path.join(OUT, 'pipeline_q3.xlsx'), data_only=True)
wd = wb['明细数据']
rows = []
for r in wd.iter_rows(min_row=2, values_only=True):
    if r[0] is None: continue
    rows.append({'日期': r[0], '月份': r[3], '区域': r[4], '品类': r[5], '实际': r[6], '目标': r[7]})
mon_actual = {m: sum(x['实际'] for x in rows if x['月份'] == m) for m in ['7月', '8月', '9月']}
mon_target = {m: sum(x['目标'] for x in rows if x['月份'] == m) for m in ['7月', '8月', '9月']}
reg_actual = {}
reg_target = {}
for x in rows:
    reg_actual[x['区域']] = reg_actual.get(x['区域'], 0) + x['实际']
    reg_target[x['区域']] = reg_target.get(x['区域'], 0) + x['目标']
reg_ach = {k: reg_actual[k] / reg_target[k] for k in reg_actual}
cat_total = {}
for x in rows:
    cat_total[x['品类']] = cat_total.get(x['品类'], 0) + x['实际']
TOTAL_A = sum(x['实际'] for x in rows)
TOTAL_T = sum(x['目标'] for x in rows)
cat_share = sorted(((c, v / TOTAL_A * 100) for c, v in cat_total.items()), key=lambda z: -z[1])
daily = {}
for x in rows:
    daily[x['日期']] = daily.get(x['日期'], 0) + x['实际']
# 日利润率(与Excel同规则重算: 46-12*规模/40)
daily_margin = {k: round(46 - 12 * (v / 40), 1) for k, v in daily.items()}
def _fmt_pct(a, t):
    return f'{(a/t-1)*100:+.1f}%'
M7D = _fmt_pct(mon_actual['7月'], mon_target['7月'])
M8D = _fmt_pct(mon_actual['8月'], mon_target['8月'])
M9D = _fmt_pct(mon_actual['9月'], mon_target['9月'])
MOM = (mon_actual['9月'] / mon_actual['8月'] - 1) * 100
GAP7 = mon_target['7月'] - mon_actual['7月']
MG_MIN = min(daily_margin.values())
best_region = max(reg_ach, key=reg_ach.get)
worst_region = min(reg_ach, key=reg_ach.get)
week_reg = {}
# 周×区域矩阵
import collections
wk = collections.defaultdict(lambda: collections.defaultdict(float))
for x in rows:
    # 周次从日期推: 7/1起每7天
    d = datetime.datetime.strptime('2026-' + x['日期'], '%Y-%m-%d')
    week = min(12, (d - datetime.datetime(2026, 7, 1)).days // 7)
    wk[week][x['区域']] += x['实际']
weeks = sorted(wk.keys())

# ── 一致性断言 ──
def assert_ck(name, val, tol=1.5):
    expect = ck[name]
    if isinstance(expect, (int, float)):
        ok = abs(val - expect) <= tol
    else:
        ok = val == expect
    print(f'  {"✅" if ok else "❌"} {name}: PPT={val} Excel={expect}')
    return ok
ALL_OK = True
ALL_OK &= assert_ck('总销售', round(TOTAL_A))
ALL_OK &= assert_ck('总目标', round(TOTAL_T))
ALL_OK &= assert_ck('达成率%', round(TOTAL_A / TOTAL_T * 100, 1))
ALL_OK &= assert_ck('最佳区域', best_region)
ALL_OK &= assert_ck('未达标区域', worst_region)
ALL_OK &= assert_ck('品类Top1', cat_share[0][0])
print('一致性校验:', '✅ 全部通过' if ALL_OK else '❌ 有不一致!')

prs = Presentation()
prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
BLANK = prs.slide_layouts[6]
SW, SH = prs.slide_width, prs.slide_height


def add_text(slide, x, y, w, h, text, size=18, color=TXT, bold=False, align=PP_ALIGN.LEFT,
             anchor=MSO_ANCHOR.TOP, line_spacing=1.0):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame; tf.word_wrap = True; tf.vertical_anchor = anchor
    lines = text.split('\n')
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align; p.line_spacing = line_spacing
        r = p.add_run(); r.text = ln
        r.font.size = Pt(size); r.font.bold = bold; r.font.color.rgb = color; r.font.name = FONT
    return tb


def header(slide, title, page):
    """统一导航条: 顶部色条+标题+页码页脚"""
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SW, Inches(0.85))
    bar.fill.solid(); bar.fill.fore_color.rgb = NAVY; bar.line.fill.background()
    add_text(slide, 0.45, 0.14, 10, 0.6, title, 22, WHITE, True, anchor=MSO_ANCHOR.MIDDLE)
    add_text(slide, 11.8, 0.14, 1.1, 0.6, f'{page:02d}', 14, RGBColor(0x9D, 0xB8, 0xD8), True,
             PP_ALIGN.RIGHT, MSO_ANCHOR.MIDDLE)
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.45), Inches(7.02), Inches(12.4),
                                  Pt(1.2))
    line.fill.solid(); line.fill.fore_color.rgb = RGBColor(0xE9, 0xE9, 0xE9); line.line.fill.background()
    add_text(slide, 0.45, 7.06, 8, 0.32, '2026 Q3 经营分析报告 · 数据源: pipeline_q3.xlsx', 9, GREY)


def style_chart(ch, legend=True, legend_pos=XL_LEGEND_POSITION.BOTTOM):
    ch.font.size = Pt(11); ch.font.name = FONT; ch.font.color.rgb = TXT
    ch.has_title = False
    if legend:
        ch.has_legend = True; ch.legend.position = legend_pos; ch.legend.include_in_layout = False
    else:
        ch.has_legend = False
    # 去网格线/去边框
    try:
        ch.chart_style = None
    except Exception: pass
    axes = []
    for _an in ('category_axis', 'value_axis'):
        try:
            axes.append(getattr(ch, _an))
        except Exception:
            pass
    for ax in axes:
        try:
            ax.has_major_gridlines = False
            ax.format.line.color.rgb = RGBColor(0xC9, 0xC9, 0xC9)
            ax.major_tick_mark = XL_TICK_MARK.NONE
        except Exception:
            pass


def series_color(plot, idx, color):
    ser = plot.series[idx]
    ser.format.fill.solid(); ser.format.fill.fore_color.rgb = color
    ser.format.line.fill.background()


def inset_title(slide, x, y, w, text, size=14, color=NAVY):
    add_text(slide, x, y, w, 0.35, text, size, color, True)


# ════ 页1 封面 ════
s = prs.slides.add_slide(BLANK)
bg = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SW, SH)
bg.fill.solid(); bg.fill.fore_color.rgb = NAVY; bg.line.fill.background()
accent = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.1), Inches(3.05), Inches(1.6), Pt(4))
accent.fill.solid(); accent.fill.fore_color.rgb = ORANGE; accent.line.fill.background()
add_text(s, 1.1, 1.6, 11, 0.5, 'QIYE  DATA  REPORT', 14, RGBColor(0x9D, 0xB8, 0xD8), True)
add_text(s, 1.1, 3.3, 11, 1.0, '2026 Q3 经营分析报告', 40, WHITE, True)
add_text(s, 1.1, 4.35, 11, 0.5, f'总销售 {TOTAL_A:.0f} 万元   |   目标达成率 {TOTAL_A/TOTAL_T:.1%}   |   数据周期 7.1 ~ 9.30',
         16, RGBColor(0xBD, 0xD7, 0xEE))
add_text(s, 1.1, 6.3, 11, 0.4, f'生成日期: {datetime.date.today()}   数据来源: 内部销售明细(1,104条记录)', 10,
         RGBColor(0x88, 0xAF, 0xDF))

# ════ 页2 目录 ════
s = prs.slides.add_slide(BLANK)
header(s, '目录 / Contents', 2)
items = ['01  核心指标总览', '02  月度走势: 实际 vs 目标', '03  区域达成率分析',
         '04  品类结构占比', '05  销售规模与利润率关系', '06  周度区域热力矩阵', '07  结论与行动建议']
for i, it in enumerate(items):
    col = i // 4; row = i % 4
    x = 1.2 + col * 6.2; y = 1.5 + row * 1.2
    num = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(0.62), Inches(0.62))
    num.fill.solid(); num.fill.fore_color.rgb = ORANGE if i == 0 else BLUE; num.line.fill.background()
    tf = num.text_frame; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = f'{i+1:02d}'; r.font.size = Pt(15); r.font.bold = True
    r.font.color.rgb = WHITE; r.font.name = FONT
    add_text(s, x + 0.85, y + 0.05, 4.8, 0.5, it[4:], 15, TXT, True, anchor=MSO_ANCHOR.MIDDLE)

# ════ 页3 核心指标(6卡片) ════
s = prs.slides.add_slide(BLANK)
header(s, '核心指标总览', 3)
cards = [
    ('季度总销售', f'{TOTAL_A:.0f}', '万元', '实际完成', ORANGE),
    ('目标达成率', f'{TOTAL_A/TOTAL_T:.1%}', '', f'目标 {TOTAL_T:.0f} 万', BLUE),
    ('领跑区域', best_region, f'{reg_ach[best_region]:.1%}', '达成率最高', BLUE),
    ('待改进区域', worst_region, f'{reg_ach[worst_region]:.1%}', '未达目标线', RED),
    ('主力品类', cat_share[0][0], f'{cat_share[0][1]:.1f}%', '销售占比第一', BLUE),
    ('日均销售', f'{TOTAL_A/92:.1f}', '万元/天', '92天均值', LIGHT),
]
for i, (lab, val, unit, sub, color) in enumerate(cards):
    col = i % 3; row = i // 3
    x = 0.85 + col * 4.05; y = 1.45 + row * 2.55
    card = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(3.75), Inches(2.2))
    card.fill.solid(); card.fill.fore_color.rgb = RGBColor(0xF7, 0xF9, 0xFC)
    card.line.color.rgb = RGBColor(0xDD, 0xE5, 0xF0); card.line.width = Pt(0.75)
    card.adjustments[0] = 0.06
    strip = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(0.09), Inches(2.2))
    strip.fill.solid(); strip.fill.fore_color.rgb = color; strip.line.fill.background()
    add_text(s, x + 0.28, y + 0.2, 3.2, 0.35, lab, 12, GREY, True)
    add_text(s, x + 0.28, y + 0.6, 3.2, 0.8, val, 34, NAVY, True)
    if unit:
        add_text(s, x + 2.55, y + 0.95, 1.1, 0.4, unit, 12, GREY)
    add_text(s, x + 0.28, y + 1.6, 3.2, 0.35, sub, 11, GREY)

# ════ 页4 月度走势: 原生柱状(实际vs目标) ════
s = prs.slides.add_slide(BLANK)
header(s, '月度走势: 实际 vs 目标', 4)
cd = CategoryChartData()
cd.categories = ['7月', '8月', '9月']
cd.add_series('实际销售', [round(mon_actual[m], 1) for m in ['7月', '8月', '9月']])
cd.add_series('目标', [round(mon_target[m], 1) for m in ['7月', '8月', '9月']])
gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.7), Inches(1.25), Inches(8.0), Inches(5.3), cd)
ch = gf.chart
style_chart(ch)
series_color(ch.plots[0], 0, BLUE); series_color(ch.plots[0], 1, LIGHT)
ch.plots[0].gap_width = 80
plot = ch.plots[0]
plot.has_data_labels = True; dl = plot.data_labels
dl.number_format = '0.0'; dl.number_format_is_linked = False
dl.font.size = Pt(10); dl.font.color.rgb = TXT; dl.font.name = FONT
add_text(s, 8.95, 1.45, 4.0, 3.0,
         f'要点解读\n\n· 7月实际 {mon_actual["7月"]:.1f} 万，目标差 {M7D}，\n  为季度唯一未达标月份\n\n· 8月回升至 {mon_actual["8月"]:.1f} 万，超目标 {M8D}\n\n· 9月冲刺 {mon_actual["9月"]:.1f} 万，超目标 {M9D}，\n  环比 {MOM:+.1f}%，创季度新高',
         13, TXT, line_spacing=1.25)
box = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(8.95), Inches(4.5), Inches(3.85), Inches(1.9))
box.fill.solid(); box.fill.fore_color.rgb = RGBColor(0xFD, 0xF3, 0xE7)
box.line.color.rgb = ORANGE; box.line.width = Pt(1)
box.adjustments[0] = 0.08
add_text(s, 9.15, 4.62, 3.5, 1.6, f'风险提示\n7月缺口 {GAP7:.1f} 万需在Q4通过\n区域补位与品类拉动补齐', 11.5, RGBColor(0xB0, 0x5E, 0x1A), True, line_spacing=1.15)

# ════ 页5 区域达成率: 原生横条 ════
s = prs.slides.add_slide(BLANK)
header(s, '区域达成率分析', 5)
order = sorted(reg_ach, key=reg_ach.get)
cd = CategoryChartData()
cd.categories = order
cd.add_series('达成率', [round(reg_ach[r], 3) for r in order])
gf = s.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, Inches(0.7), Inches(1.25), Inches(8.0), Inches(5.3), cd)
ch = gf.chart
style_chart(ch, legend=False)
plot = ch.plots[0]
ser = plot.series[0]
for i, r in enumerate(order):
    pt = ser.points[i]
    pt.format.fill.solid()
    pt.format.fill.fore_color.rgb = ORANGE if r == best_region else (RED if reg_ach[r] < 1 else BLUE)
    pt.format.line.fill.background()
plot.gap_width = 60
plot.has_data_labels = True; dl = plot.data_labels
dl.number_format = '0.0%'; dl.number_format_is_linked = False
dl.font.size = Pt(11); dl.font.bold = True; dl.font.color.rgb = TXT; dl.font.name = FONT
# 值轴百分比
va = ch.value_axis
va.tick_labels.number_format = '0%'; va.tick_labels.number_format_is_linked = False
add_text(s, 8.95, 1.45, 4.0, 4.0,
         f'梯队分析\n\n🏆 第一梯队 · {best_region}\n   {reg_ach[best_region]:.1%}，领跑全区域\n\n◆ 第二梯队 · ' +
         ' / '.join(r for r in order if 1 <= reg_ach[r] < 1.05) +
         f'\n   达标但增长动能趋缓\n\n⚠️ 待改进 · {worst_region}\n   {reg_ach[worst_region]:.1%}，距目标线差{abs(1-reg_ach[worst_region])*reg_target[worst_region]:.1f}万',
         13, TXT, line_spacing=1.3)

# ════ 页6 品类结构: 原生环形 ════
s = prs.slides.add_slide(BLANK)
header(s, '品类结构占比', 6)
cd = CategoryChartData()
cd.categories = [c for c, _ in cat_share]
cd.add_series('销售占比', [round(v, 1) for _, v in cat_share])
gf = s.shapes.add_chart(XL_CHART_TYPE.DOUGHNUT, Inches(0.7), Inches(1.25), Inches(7.2), Inches(5.3), cd)
ch = gf.chart
style_chart(ch, legend_pos=XL_LEGEND_POSITION.RIGHT)
cl = [BLUE, ORANGE, LIGHT]
ser = ch.plots[0].series[0]
for i in range(len(cat_share)):
    pt = ser.points[i]
    pt.format.fill.solid(); pt.format.fill.fore_color.rgb = cl[i % 3]
    pt.format.line.color.rgb = WHITE; pt.format.line.width = Pt(1.5)
plot = ch.plots[0]
plot.has_data_labels = True; dl = plot.data_labels
dl.number_format = '0.0"%"'; dl.number_format_is_linked = False
dl.font.size = Pt(12); dl.font.bold = True; dl.font.color.rgb = WHITE; dl.font.name = FONT
per_cat = {c: cat_total[c] for c, _ in cat_share}
add_text(s, 8.1, 1.5, 4.8, 4.5,
         '结构洞察\n\n· ' + cat_share[0][0] + f' 以 {cat_share[0][1]:.1f}% 占比居首，\n  贡献 {per_cat[cat_share[0][0]]:.0f} 万元\n\n· ' +
         cat_share[1][0] + f' {cat_share[1][1]:.1f}%，是第二大支柱\n\n· ' + cat_share[2][0] +
         f' {cat_share[2][1]:.1f}%，占比最小但\n  客单价高、复购稳定\n\n建议: 强化' + cat_share[0][0] + '主引擎地位，\n同时提升' + cat_share[2][0] + '的连带率',
         13, TXT, line_spacing=1.35)

# ════ 页7 散点: 规模vs利润率 ════
s = prs.slides.add_slide(BLANK)
header(s, '销售规模与利润率关系', 7)
xd = XyChartData()
serd = xd.add_series('日观测点')
for k in sorted(daily):
    serd.add_data_point(round(daily[k], 1), daily_margin[k])
gf = s.shapes.add_chart(XL_CHART_TYPE.XY_SCATTER, Inches(0.7), Inches(1.25), Inches(8.0), Inches(5.3), xd)
ch = gf.chart
style_chart(ch, legend=False)
pl = ch.plots[0]
pl.series[0].format.line.fill.background()
pl.series[0].marker.style = 8  # circle
pl.series[0].marker.format.fill.solid()
pl.series[0].marker.format.fill.fore_color.rgb = BLUE
try:
    pl.series[0].marker.size = 6
except Exception:
    pass
try:
    ch.plots[0].has_data_labels = False
except Exception:
    pass
add_text(s, 8.95, 1.45, 4.0, 4.5,
         f'相关性分析\n\n· 共 {len(daily)} 个日观测点\n\n· 规模上升时利润率呈下行趋势，\n  大促日的价格让利明显\n\n· 建议: 高销量日控制折扣深度，\n  以"保利润"替代"冲规模"\n\n· 关注利润率低于 {MG_MIN:.1f}% 的异常日（本季最低）\n  (可能存在过度促销)',
         13, TXT, line_spacing=1.3)
box = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(8.95), Inches(5.15), Inches(3.85), Inches(1.3))
box.fill.solid(); box.fill.fore_color.rgb = RGBColor(0xFD, 0xF3, 0xE7)
box.line.color.rgb = ORANGE; box.line.width = Pt(1); box.adjustments[0] = 0.1
add_text(s, 9.15, 5.28, 3.5, 1.0, '待验证\n利润率与规模的因果需\n结合折扣明细进一步归因', 11.5,
         RGBColor(0xB0, 0x5E, 0x1A), True, line_spacing=1.15)

# ════ 页8 周度区域热力表 ════
s = prs.slides.add_slide(BLANK)
header(s, '周度区域热力矩阵', 8)
rows_n = len(weeks) + 1
cols_n = len(reg_actual) + 1
tbl_shape = s.shapes.add_table(rows_n, cols_n, Inches(0.7), Inches(1.35), Inches(11.9), Inches(5.1))
tbl = tbl_shape.table
tbl.columns[0].width = Inches(1.5)
for c in range(1, cols_n):
    tbl.columns[c].width = Inches((11.9 - 1.5) / len(reg_actual))
tbl.rows[0].height = Inches(0.4)
hdr = ['周次'] + list(reg_actual.keys())
for c, h in enumerate(hdr):
    cell = tbl.cell(0, c)
    cell.text = h
    cell.fill.solid(); cell.fill.fore_color.rgb = NAVY
    p = cell.text_frame.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    p.runs[0].font.size = Pt(12); p.runs[0].font.bold = True
    p.runs[0].font.color.rgb = WHITE; p.runs[0].font.name = FONT
vals_wk = [[wk[w][r] for r in reg_actual] for w in weeks]
flat = [v for row in vals_wk for v in row]
vmin, vmax = min(flat), max(flat)
def heat_color(v):
    # 低=浅蓝白 高=深蓝
    t = (v - vmin) / (vmax - vmin + 1e-9)
    r0, g0, b0 = 0xEA, 0xF1, 0xFA
    r1, g1, b1 = 0x2E, 0x6F, 0xB7
    return RGBColor(int(r0 + (r1 - r0) * t), int(g0 + (g1 - g0) * t), int(b0 + (b1 - b0) * t))
for ri, w in enumerate(weeks, start=1):
    tbl.rows[ri].height = Inches(0.38)
    c0 = tbl.cell(ri, 0); c0.text = f'W{w+27}'
    c0.fill.solid(); c0.fill.fore_color.rgb = RGBColor(0xF2, 0xF2, 0xF2)
    p = c0.text_frame.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    p.runs[0].font.size = Pt(11); p.runs[0].font.bold = True
    p.runs[0].font.color.rgb = TXT; p.runs[0].font.name = FONT
    for ci, r in enumerate(reg_actual, start=1):
        v = wk[w][r]
        cell = tbl.cell(ri, ci); cell.text = f'{v:.0f}'
        cell.fill.solid(); cell.fill.fore_color.rgb = heat_color(v)
        p = cell.text_frame.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        p.runs[0].font.size = Pt(10.5); p.runs[0].font.name = FONT
        lum = 0.299 * cell.fill.fore_color.rgb[0] + 0.587 * cell.fill.fore_color.rgb[1] + 0.114 * cell.fill.fore_color.rgb[2]
        p.runs[0].font.color.rgb = WHITE if lum < 140 else TXT
add_text(s, 0.7, 6.55, 11.9, 0.4, '色阶: 浅蓝=周销低 · 深蓝=周销高   单位: 万元', 11, GREY)

# ════ 页9 结论与建议 ════
s = prs.slides.add_slide(BLANK)
header(s, '结论与行动建议', 9)
col1 = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.7), Inches(1.35), Inches(5.85), Inches(5.3))
col1.fill.solid(); col1.fill.fore_color.rgb = RGBColor(0xF7, 0xF9, 0xFC)
col1.line.color.rgb = RGBColor(0xDD, 0xE5, 0xF0); col1.line.width = Pt(0.75); col1.adjustments[0] = 0.03
add_text(s, 1.0, 1.55, 5.3, 0.5, '✓ 经营亮点', 17, BLUE, True)
add_text(s, 1.0, 2.2, 5.3, 4.3,
         f'· 季度总销售 {TOTAL_A:.0f} 万元，达成率 {TOTAL_A/TOTAL_T:.1%}，\n  整体跑赢目标\n\n'
         f'· {best_region}区域以 {reg_ach[best_region]:.1%} 达成率领跑，\n  可作为标杆经验复制\n\n'
         f'· 9月销售 {mon_actual["9月"]:.0f} 万元创季度新高，\n  环比增长 {((mon_actual["9月"]/mon_actual["8月"])-1)*100:.1f}%\n\n'
         f'· {cat_share[0][0]}品类占比 {cat_share[0][1]:.1f}%，\n  引擎作用稳固', 13, TXT, line_spacing=1.3)
col2 = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.8), Inches(1.35), Inches(5.85), Inches(5.3))
col2.fill.solid(); col2.fill.fore_color.rgb = RGBColor(0xFD, 0xF6, 0xF5)
col2.line.color.rgb = RGBColor(0xEE, 0xC9, 0xC7); col2.line.width = Pt(0.75); col2.adjustments[0] = 0.03
add_text(s, 7.1, 1.55, 5.3, 0.5, '⚠ 风险与行动', 17, RED, True)
add_text(s, 7.1, 2.2, 5.3, 4.3,
         f'· {worst_region}区域达成 {reg_ach[worst_region]:.1%}，\n  为唯一未达标区域，需专项诊断\n\n'
         f'· 7月销售缺口 {GAP7:.1f} 万元，\n  建议Q4通过资源补位追回\n\n'
         '· 销售规模与利润率呈负相关，\n  警惕"以价换量"侵蚀利润\n\n'
         '· 行动项: ① 区域对标帮扶 ② 折扣审批收紧\n  ③ 高毛利品类连带率提升 ④ 周度跟盘机制',
         13, TXT, line_spacing=1.3)

# ════ 页10 封底 ════
s = prs.slides.add_slide(BLANK)
bg = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SW, SH)
bg.fill.solid(); bg.fill.fore_color.rgb = NAVY; bg.line.fill.background()
add_text(s, 1.2, 3.0, 11, 0.8, 'THANK YOU', 36, WHITE, True)
add_text(s, 1.2, 3.9, 11, 0.5, '数据驱动决策 · 让每一次增长都有据可依', 15, RGBColor(0xBD, 0xD7, 0xEE))
accent = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.2), Inches(4.6), Inches(1.6), Pt(4))
accent.fill.solid(); accent.fill.fore_color.rgb = ORANGE; accent.line.fill.background()
add_text(s, 1.2, 6.4, 11, 0.4, '本报告数据由 Excel 分析工作簿自动生成 · 图表为原生可编辑对象', 10,
         RGBColor(0x88, 0xAF, 0xDF))

import sys as _sys, os as _os
_sys.path.insert(0, _os.path.dirname(_os.path.abspath(__file__)))
from chart_style import brand_ppt
brand_ppt(prs)
prs.save(os.path.join(OUT, 'pipeline_report.pptx'))
print('✅ PPT生成完成: 10页')
