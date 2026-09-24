# -*- coding: utf-8 -*-
"""复杂场景·完全体: Q3经营分析工作簿(1092行明细→透视→4图仪表盘)
数据: 91天×4区域×3品类 合成业务数据(周内季节性+区域权重+品类权重+上升趋势+噪声)"""
import sys, os, random, datetime
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import Reference
from openpyxl.formatting.rule import ColorScaleRule
from chart_style import (styled_combo, styled_waterfall, styled_bar, styled_stacked,
                         fit_to_page, BLUE, ORANGE, NAVY, LIGHT)
from openpyxl.drawing.spreadsheet_drawing import AbsoluteAnchor
from openpyxl.drawing.xdr import XDRPoint2D, XDRPositiveSize2D
from openpyxl.utils.units import cm_to_EMU as C2E
GRID = {'pareto': (0.9, 1.7), 'stacked': (13.4, 1.7), 'region': (0.9, 10.5), 'waterfall': (13.4, 10.5)}
def abs_anchor(key):
    x, y = GRID[key]
    return AbsoluteAnchor(pos=XDRPoint2D(C2E(x), C2E(y)), ext=XDRPositiveSize2D(C2E(12.4), C2E(8.2)))

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'output')
os.makedirs(OUT, exist_ok=True)

random.seed(20260923)
REGIONS = [('华东', 0.37), ('华南', 0.31), ('华北', 0.21), ('西南', 0.11)]
CATS = [('智能硬件', 0.42), ('大家电', 0.33), ('厨卫电器', 0.25)]
START = datetime.date(2026, 7, 1)
DAYS = 91  # 13个整周

# ── 生成明细 ──
rows = []  # (日期, 星期, 周, 月, 区域, 品类, 销售额)
for d in range(DAYS):
    date = START + datetime.timedelta(days=d)
    wd = date.weekday()  # 0=周一
    wf = {0: 0.85, 1: 0.9, 2: 0.95, 3: 1.0, 4: 1.15, 5: 1.3, 6: 1.2}[wd]
    trend = 1 + 0.15 * d / DAYS          # 季度内缓慢上行
    month = f'{date.month}月'
    week = f'W{d // 7 + 27}'             # W27..W39
    for rname, rw in REGIONS:
        for cname, cw in CATS:
            v = 10.1 * rw * cw * wf * trend * random.uniform(0.82, 1.18)
            rows.append((date.strftime('%m-%d'), '一二三四五六日'[wd], week, month, rname, cname, round(v, 2)))

TOTAL = sum(r[6] for r in rows)

def pivot(dim_idx, dims=None):
    """按(周/月, 区域/品类)透视: 返回行标签序列 + {行: {列: 值}}"""
    out_labels, out = [], {}
    for r in rows:
        k = r[dim_idx]
        if k not in out:
            out[k] = {d: 0 for d in dims}; out_labels.append(k)
        out[k][r[4] if dim_idx == 2 else r[5]] += r[6]
    return out_labels, out

weeks, wk_region = pivot(2, [r[0] for r in REGIONS])
months, mon_cat = pivot(3, [c[0] for c in CATS])

# 品类合计+帕累托
cat_total = {c[0]: 0 for c in CATS}
for r in rows: cat_total[r[5]] += r[6]
pareto = sorted(cat_total.items(), key=lambda x: -x[1])
cum, cums = 0, []
for _, v in pareto:
    cum += v; cums.append(cum / TOTAL)
region_total = {r[0]: 0 for r in REGIONS}
for r in rows: region_total[r[4]] += r[6]

wb = openpyxl.Workbook()

# ══ Sheet1: 明细数据(隐藏,11周数据源) ══
ws_d = wb.active; ws_d.title = '明细数据'
ws_d.append(['日期', '星期', '周次', '月份', '区域', '品类', '销售额(万元)'])
for r in rows: ws_d.append(list(r))
for c in 'G': ws_d.column_dimensions['G'].width = 12
ws_d.sheet_state = 'hidden'

# ══ Sheet2: 周×区域热力矩阵 + 色阶条件格式 ══
ws_h = wb.create_sheet('周度热力')
ws_h.append(['周次'] + [r[0] for r in REGIONS] + ['周合计'])
for w in weeks:
    row = [w] + [round(wk_region[w][r[0]], 1) for r in REGIONS] + [round(sum(wk_region[w].values()), 1)]
    ws_h.append(row)
ws_h.append(['季合计'] + [round(region_total[r[0]], 1) for r in REGIONS] + [round(TOTAL, 1)])
ws_h['A16'] = '色阶: 浅→深 = 低→高 (条件格式ColorScale,随数据自动更新)'
ws_h['A16'].font = Font(size=9, color='595959', name='微软雅黑')
ws_h.conditional_formatting.add(f'B2:E{1 + len(weeks)}',
    ColorScaleRule(start_type='min', start_color='FFFFFF', end_type='max', end_color='2E6FB7'))
for c in range(1, 7):
    ws_h.cell(row=1, column=c).font = Font(bold=True, color=NAVY, name='微软雅黑')
    ws_h.cell(row=1 + len(weeks) + 1, column=c).font = Font(bold=True, color=NAVY, name='微软雅黑')
    for r in range(2, 2 + len(weeks) + 1):
        ws_h.cell(row=r, column=c).number_format = '#,##0.0'
for col, wd in zip('BCDEF', [10, 10, 10, 10, 12]): ws_h.column_dimensions[col].width = wd
fit_to_page(ws_h)

# ══ Sheet3: 仪表盘(2×2四图) ══
ws_v = wb.create_sheet('仪表盘')
ws_v['A1'] = '2026 Q3 经营分析仪表盘'
ws_v['A1'].font = Font(size=16, bold=True, color=NAVY, name='微软雅黑')

# 图1: 帕累托(品类) — 复用双轴
ws_p = wb.create_sheet('帕累托源')
ws_p.append(['品类', '销售额(万元)', '累计占比'])
for i, (name, v) in enumerate(pareto):
    ws_p.append([name, round(v, 1), round(cums[i], 4)])
    ws_p.cell(row=2 + i, column=3).number_format = '0%'
# 帕累托(柱=销售额,线=累计占比,次轴%)
ch = styled_combo(
    ws_v,
    Reference(ws_p, min_col=2, min_row=1, max_col=2, max_row=1 + len(pareto)),
    Reference(ws_p, min_col=3, min_row=1, max_col=3, max_row=1 + len(pareto)),
    Reference(ws_p, min_col=1, min_row=2, max_row=1 + len(pareto)),
    f'品类帕累托：{pareto[0][0]}+{pareto[1][0]}贡献{cums[1]:.0%}销售额', 'B3', y2_fmt='0%')
ch.anchor = abs_anchor('pareto')
ws_p.sheet_state = 'hidden'

# 图2: 堆叠柱(月×品类构成)
ws_s = wb.create_sheet('堆叠源')
ws_s.append(['月份'] + [c[0] for c in CATS])
for m in months:
    ws_s.append([m] + [round(mon_cat[m][c[0]], 1) for c in CATS])
ch = styled_stacked(
    ws_v,
    Reference(ws_s, min_col=2, min_row=1, max_col=1 + len(CATS), max_row=1 + len(months)),
    Reference(ws_s, min_col=1, min_row=2, max_row=1 + len(months)),
    f'品类构成：{months[0]}–{months[-1]}逐月(万元)', 'I3')
ch.anchor = abs_anchor('stacked')
ws_s.sheet_state = 'hidden'

# 图3: 区域排名(横向条形,升序排→最大在顶)
ws_r = wb.create_sheet('区域源')
ws_r.append(['区域', '销售额(万元)'])
for name, v in sorted(region_total.items(), key=lambda x: x[1]):
    ws_r.append([name, round(v, 1)])
ch = styled_bar(
    ws_v,
    Reference(ws_r, min_col=2, min_row=1, max_col=2, max_row=1 + len(REGIONS)),
    Reference(ws_r, min_col=1, min_row=2, max_row=1 + len(REGIONS)),
    f'区域排名：华东{region_total["华东"]:.0f}万领跑', 'B20',
    colors=(BLUE,), labels=True, legend=False, horizontal=True, gap=60)
ch.anchor = abs_anchor('region')
from openpyxl.chart.series import DataPoint
best = max(range(len(REGIONS)), key=lambda i: region_total[REGIONS[i][0]])
order = [n for n, _ in sorted(region_total.items(), key=lambda x: x[1])]
ch.series[0].data_points = [DataPoint(idx=i, spPr=openpyxl.chart.shapes.GraphicalProperties(
    solidFill=(ORANGE if order[i] == '华东' else BLUE),
    ln=openpyxl.drawing.line.LineProperties(noFill=True))) for i in range(len(REGIONS))]
ws_r.sheet_state = 'hidden'

# 图4: 瀑布(季度利润桥) — 复用v1.2
ws_w = wb.create_sheet('瀑布源')
WF = [('Q2净利', 480, 'total'), ('营收增长', 186, 'up'), ('成本', -92, 'down'),
      ('费用', -58, 'down'), ('税', -24, 'down'), ('Q3净利', 492, 'total')]
ws_w.append(['项目'] + [w[0] for w in WF])
base, delta, colors = [], [], []
run = 0
for name, v, kind in WF:
    if kind == 'total': base.append(0.5); delta.append(v - 0.5); colors.append(NAVY); run = v
    elif kind == 'up': base.append(run); delta.append(v); colors.append(BLUE); run += v
    else: base.append(run + v); delta.append(-v); colors.append(ORANGE); run += v
ws_w.append(['基准(隐藏)'] + base)
ws_w.append(['金额(万元)'] + delta)
ch = styled_waterfall(
    ws_v,
    Reference(ws_w, min_col=2, min_row=1, max_col=1 + len(WF), max_row=1),
    Reference(ws_w, min_col=2, min_row=2, max_col=1 + len(WF), max_row=2),
    Reference(ws_w, min_col=2, min_row=3, max_col=1 + len(WF), max_row=3),
    colors,
    f'Q3净利492万：营收+186，费用端-82', 'I20')
ch.anchor = abs_anchor('waterfall')
ws_w.sheet_state = 'hidden'
ws_v.page_setup.orientation = 'landscape'
from openpyxl.worksheet.page import PageMargins
ws_v.page_margins = PageMargins(left=0.25, right=0.25, top=0.3, bottom=0.3, header=0.1, footer=0.1)

import sys as _sys, os as _os
_sys.path.insert(0, _os.path.dirname(_os.path.abspath(__file__)))
from chart_style import set_brand
set_brand(wb, title='夏制 · 经营仪表盘 v2', subject='夏制 · 经营仪表盘 v2 · 盛夏蓝橙 · 照镜自检')
wb.save(os.path.join(OUT, 'dashboard_v2.xlsx'))
print(f'✅ 明细{len(rows)}行 | 总销售{TOTAL:.0f}万 | 帕累托: {[(n, f"{c:.0%}") for (n, _), c in zip(pareto, cums)]}')
print(f'✅ 区域: {[(n, round(v)) for n, v in region_total.items()]}')
