# -*- coding: utf-8 -*-
"""复杂场景: Q3经营仪表盘(4个sheet) — KPI总览 / 双轴组合 / 瀑布利润桥 / 区域对比高亮"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import Reference
from chart_style import styled_combo, styled_waterfall, styled_bar, fit_to_page, BLUE, ORANGE, NAVY, LIGHT, GraphicalProperties, LineProperties

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'output')
os.makedirs(OUT, exist_ok=True)

MONTHS = ['4月', '5月', '6月', '7月', '8月', '9月']
SALES  = [120, 135, 158, 142, 176, 190]          # 万元
MARGIN = [37.5, 38.5, 38.6, 35.2, 39.8, 41.1]    # %
REGIONS = [('华东', 342), ('华南', 287), ('华北', 195), ('西南', 97)]
WF = [('Q2净利', 480, 'total'), ('营收增长', 186, 'up'), ('成本', -92, 'down'),
      ('费用', -58, 'down'), ('税', -24, 'down'), ('Q3净利', 492, 'total')]

wb = openpyxl.Workbook()

# ── Sheet1: KPI总览 ──
ws = wb.active; ws.title = 'KPI总览'
kpis = [('累计销售(万元)', '921', '环比 +12.6%'), ('净利率', '38.2%', '同比 +1.8pct'),
        ('活跃客户', '1,284', '本季新增 86'), ('交付准时率', '96.4%', '目标 95%')]
for i, (label, value, note) in enumerate(kpis):
    col = 2 + i * 3
    c1 = ws.cell(row=3, column=col, value=label)
    c1.font = Font(size=11, color='595959', name='微软雅黑')
    c2 = ws.cell(row=5, column=col, value=value)
    c2.font = Font(size=24, bold=True, color=NAVY, name='微软雅黑')
    c3 = ws.cell(row=7, column=col, value=note)
    c3.font = Font(size=10, color=ORANGE if i == 0 else '4C9F70', name='微软雅黑')
    for r in range(3, 8):
        for cc in range(col, col + 2):
            ws.cell(row=r, column=cc).fill = PatternFill('solid', fgColor='F4F8FC')
ws.cell(row=10, column=2, value='2026 Q3 经营仪表盘 · 数据截至9月30日 · 制表: office-visual-style').font = Font(size=9, color='B0B0B0', name='微软雅黑')
fit_to_page(ws)

# ── Sheet2: 双轴组合(销售额柱 + 利润率线) — 标准列向布局: 类别一列,系列各一列 ──
ws2 = wb.create_sheet('销售与利润率')
ws2.append(['月份', '销售额(万元)', '利润率(%)'])
for m, s, mgn in zip(MONTHS, SALES, MARGIN):
    ws2.append([m, s, mgn])
styled_combo(
    ws2,
    Reference(ws2, min_col=2, min_row=1, max_col=2, max_row=1 + len(MONTHS)),
    Reference(ws2, min_col=3, min_row=1, max_col=3, max_row=1 + len(MONTHS)),
    Reference(ws2, min_col=1, min_row=2, max_row=1 + len(MONTHS)),
    '销售额连续两季上涨，9月利润率创41.1%新高', 'B2')
fit_to_page(ws2)

# ── Sheet3: 瀑布利润桥 ──
ws3 = wb.create_sheet('利润桥')
ws3.append(['项目'] + [w[0] for w in WF])
base, delta, colors = [], [], []
run = 0
for i, (name, v, kind) in enumerate(WF):
    if kind == 'total':
        base.append(0); delta.append(v); colors.append(NAVY); run = v
    elif kind == 'up':
        base.append(run); delta.append(v); colors.append(BLUE); run += v
    else:
        base.append(run + v); delta.append(-v); colors.append(ORANGE); run += v
ws3.append(['基准(隐藏)'] + base)
ws3.append(['金额(万元)'] + delta)
for r in (2, 3):
    for c in range(2, 8):
        ws3.cell(row=r, column=c).number_format = '#,##0'
styled_waterfall(
    ws3,
    Reference(ws3, min_col=2, min_row=1, max_col=7, max_row=1),
    Reference(ws3, min_col=2, min_row=2, max_col=7, max_row=2),
    Reference(ws3, min_col=2, min_row=3, max_col=7, max_row=3),
    colors,
    f'Q3净利{WF[-1][1]}万：营收贡献+{WF[1][1]}万，费用端合计拖累{-(WF[2][1]+WF[3][1]+WF[4][1])}万', 'B5')
ws3.row_dimensions[2].hidden = True
ws3.cell(row=34, column=2, value='注: 蓝色=增利, 橙色=减利, 深蓝=起止净利 | 上方表格可改,图表自动更新').font = Font(size=9, color='595959', name='微软雅黑')
fit_to_page(ws3)

# ── Sheet4: 区域对比(高亮最优) — 注意: styled_bar按列取数,数据必须纵向排 ──
ws4 = wb.create_sheet('区域对比')
ws4.append(['区域', '销售额(万元)'])
for name, v in REGIONS:
    ws4.append([name, v])
best = max(range(len(REGIONS)), key=lambda i: REGIONS[i][1])
ch = styled_bar(
    ws4,
    Reference(ws4, min_col=2, min_row=1, max_col=2, max_row=1 + len(REGIONS)),
    Reference(ws4, min_col=1, min_row=2, max_row=1 + len(REGIONS)),
    f'华东贡献342万领跑，西南仍有{REGIONS[-1][1]}万空间', 'B2',
    colors=(BLUE,), labels=True, legend=False)
from openpyxl.chart.series import DataPoint
from chart_style import GraphicalProperties, LineProperties
ch.series[0].data_points = [DataPoint(idx=i, spPr=GraphicalProperties(
    solidFill=(ORANGE if i == best else BLUE), ln=LineProperties(noFill=True)))
    for i in range(len(REGIONS))]
fit_to_page(ws4)

import sys as _sys, os as _os
_sys.path.insert(0, _os.path.dirname(_os.path.abspath(__file__)))
from chart_style import set_brand
set_brand(wb, title='夏制 · 经营仪表盘', subject='夏制 · 经营仪表盘 · 盛夏蓝橙 · 照镜自检')
wb.save(os.path.join(OUT, 'dashboard_q3.xlsx'))
print('仪表盘生成 ✅ 4 sheets:', wb.sheetnames)
print('瀑布基准:', base, '| delta:', delta)
