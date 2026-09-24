# -*- coding: utf-8 -*-
"""管线压测·Excel端: 一套数据源→复杂分析工作簿(6种新视觉场景)
场景: 实际vs目标同轴 / 区域达成率横条 / 品类环形 / 日销售-利润率散点 / 月×区域三色热力 / 数据条"""
import sys, os, random, datetime
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.chart import Reference
from openpyxl.formatting.rule import ColorScaleRule, DataBarRule
from chart_style import (styled_target, styled_donut, styled_scatter, styled_bar,
                         fit_to_page, BLUE, ORANGE, NAVY, LIGHT)

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'output')
os.makedirs(OUT, exist_ok=True)
random.seed(20260924)

REGIONS = [('华东', 0.37, 1.06), ('华南', 0.31, 1.02), ('华北', 0.21, 1.01), ('西南', 0.11, 0.91)]
# (区域, 权重, 目标达成系数) — 华北/西南未达标→有趣的达成率分化
CATS = [('智能硬件', 0.42), ('大家电', 0.33), ('厨卫电器', 0.25)]
START = datetime.date(2026, 7, 1)
DAYS = 92
MONTH_OF = lambda d: f'{(START + datetime.timedelta(days=d)).month}月'

# 目标: 月度区域目标(万) — 月度爬坡 300/330/360 × 区域权重 × 达成系数
region_names = [r[0] for r in REGIONS]
target_mr = {(m, r[0]): round(mbase * r[1], 1)
             for mi, mbase in enumerate([300, 330, 360])
             for m in ['7月', '8月', '9月']
             for r in REGIONS
             for mbase in [mbase]}

# 重建干净的目标矩阵
target_mr = {}
for mi, mbase in enumerate([300, 330, 360]):
    m = ['7月', '8月', '9月'][mi]
    for name, w, ach in REGIONS:
        target_mr[(m, name)] = round(mbase * w, 1)

# ── 明细生成 ──
rows = []  # (日期, 星期, 周次, 月份, 区域, 品类, 实际, 目标按日均摊)
for d in range(DAYS):
    date = START + datetime.timedelta(days=d)
    wd = date.weekday()
    wf = {0: 0.85, 1: 0.9, 2: 0.95, 3: 1.0, 4: 1.15, 5: 1.3, 6: 1.2}[wd]
    trend = 1 + 0.15 * d / DAYS
    month = MONTH_OF(d)
    week = f'W{d // 7 + 27}'
    daily_target = sum(target_mr[(month, r[0])] for r in REGIONS) / 30.4
    for rname, rw, ach in REGIONS:
        r_target = target_mr[(month, rname)] / 30.4
        for cname, cw in CATS:
            base = daily_target * rw * cw
            actual = base * wf * trend * random.uniform(0.9, 1.1) * ach
            rows.append((date.strftime('%m-%d'), '一二三四五六日'[wd], week, month,
                         rname, cname, round(actual, 2), round(r_target * cw / sum(c[1] for c in CATS), 4)))

months = ['7月', '8月', '9月']
reg_names = [r[0] for r in REGIONS]
# 区域达成率精确校准(抵消季节性/趋势稀释,保证区域故事成立)
for name, w, ach in REGIONS:
    a = sum(r[6] for r in rows if r[4] == name)
    t = sum(target_mr[(m, name)] for m in months)
    scale = (ach * t) / a
    for i, r in enumerate(rows):
        if r[4] == name: rows[i] = r[:6] + (round(r[6] * scale, 2),) + r[7:]

TOTAL_A = sum(r[6] for r in rows)
TOTAL_T = sum(r[7] for r in rows)
ACH = TOTAL_A / TOTAL_T

# 聚合(全部从明细派生,保证Excel/PPT口径一致)
mon_reg = {}   # (月, 区域) -> [实际, 目标]
cat_total = {}
daily = {}     # 日期 -> [实际合计, 加权利润率]
for r in rows:
    key = (r[3], r[4])
    if key not in mon_reg: mon_reg[key] = [0, 0]
    mon_reg[key][0] += r[6]; mon_reg[key][1] += r[7]
    cat_total[r[5]] = cat_total.get(r[5], 0) + r[6]
    daily[r[0]] = daily.get(r[0], [0, 0])
    daily[r[0]][0] += r[6]
# 日利润率: 与销售额弱负相关+噪声(规模大了边际利润略降)
margin_base = {k: round(46 - 12 * (v[0] / 40) + random.uniform(-2.5, 2.5), 1) for k, v in daily.items()}
for k in daily: daily[k][1] = margin_base[k]

ach_region = {r: mon_reg[(m, r)][0] / mon_reg[(m, r)][1]
              for r in reg_names for m in months}
ach_region_all = {r: sum(mon_reg[(m, r)][0] for m in months) / sum(mon_reg[(m, r)][1] for m in months)
                  for r in reg_names}
best_region = max(reg_names, key=lambda r: ach_region_all[r])
worst_region = min(reg_names, key=lambda r: ach_region_all[r])
cat_share = sorted(((c, v / TOTAL_A * 100) for c, v in cat_total.items()), key=lambda x: -x[1])

wb = openpyxl.Workbook()

# ══ Sheet1: 明细(隐藏) ══
ws_d = wb.active; ws_d.title = '明细数据'
ws_d.append(['日期', '星期', '周次', '月份', '区域', '品类', '实际(万元)', '日均目标(万元)'])
for r in rows: ws_d.append(list(r))
ws_d.sheet_state = 'hidden'

# ══ Sheet2: 月度汇总(可见,含数据条+达成率) ══
ws_m = wb.create_sheet('月度汇总')
ws_m.append(['月份', '区域', '实际(万元)', '目标(万元)', '达成率'])
r_idx = 2
for m in months:
    for rn in reg_names:
        a, t = mon_reg[(m, rn)]
        ws_m.append([m, rn, round(a, 1), round(t, 1), round(a / t, 3)])
        ws_m.cell(row=r_idx, column=5).number_format = '0.0%'
        r_idx += 1
# 数据条: 达成率列
ws_m.conditional_formatting.add(f'E2:E{r_idx-1}',
    DataBarRule(start_type='num', start_value=0.8, end_type='num', end_value=1.1,
                color="2E6FB7", showValue=True, minLength=None, maxLength=None))
for c in range(1, 6):
    ws_m.cell(row=1, column=c).font = Font(bold=True, color=NAVY, name='微软雅黑')
for col, wd in zip('ABCDE', [8, 8, 12, 12, 10]): ws_m.column_dimensions[col].width = wd
fit_to_page(ws_m)

# ══ Sheet3: 仪表盘(2×2: 目标线/达成率/环形/散点) ══
ws_v = wb.create_sheet('仪表盘')
ws_v['A1'] = f'2026 Q3 经营分析: 总销售{TOTAL_A:.0f}万 / 目标{TOTAL_T:.0f}万 / 达成{ACH:.1%}'
ws_v['A1'].font = Font(size=15, bold=True, color=NAVY, name='微软雅黑')
ws_v.page_setup.orientation = 'landscape'
from openpyxl.worksheet.page import PageMargins
ws_v.page_margins = PageMargins(left=0.25, right=0.25, top=0.3, bottom=0.3, header=0.1, footer=0.1)

from openpyxl.drawing.spreadsheet_drawing import AbsoluteAnchor
from openpyxl.drawing.xdr import XDRPoint2D, XDRPositiveSize2D
from openpyxl.utils.units import cm_to_EMU as C2E
GRID = {'target': (0.85, 1.9), 'ach': (13.7, 1.9), 'donut': (0.85, 10.9), 'scatter': (13.7, 10.9)}
def abs_anchor(key):
    x, y = GRID[key]
    return AbsoluteAnchor(pos=XDRPoint2D(C2E(x), C2E(y)), ext=XDRPositiveSize2D(C2E(12.4), C2E(8.2)))

# ── 图1: 实际vs目标(同轴柱+虚线) ──
ws_t = wb.create_sheet('月度目标源')
ws_t.append(['月份', '实际(万元)', '目标(万元)'])
for m in months:
    a = sum(mon_reg[(m, r)][0] for r in reg_names)
    t = sum(mon_reg[(m, r)][1] for r in reg_names)
    ws_t.append([m, round(a, 1), round(t, 1)])
ws_t.sheet_state = 'hidden'
ch = styled_target(
    ws_v,
    Reference(ws_t, min_col=2, min_row=1, max_col=2, max_row=4),
    Reference(ws_t, min_col=3, min_row=1, max_col=3, max_row=4),
    Reference(ws_t, min_col=1, min_row=2, max_row=4),
    '月度实际 vs 目标（万元）', 'B3')
ch.anchor = abs_anchor('target')
ws_t.sheet_state = 'hidden'

# ── 图2: 区域达成率(横条+%) ──
ws_a = wb.create_sheet('达成率源')
ws_a.append(['区域', '达成率'])
for name in sorted(reg_names, key=lambda r: ach_region_all[r]):
    ws_a.append([name, round(ach_region_all[name], 3)])
    ws_a.cell(row=ws_a.max_row, column=2).number_format = '0.0%'
ws_a.sheet_state = 'hidden'
ch = styled_bar(
    ws_v,
    Reference(ws_a, min_col=2, min_row=1, max_col=2, max_row=1 + len(reg_names)),
    Reference(ws_a, min_col=1, min_row=2, max_row=1 + len(reg_names)),
    f'区域达成率：{best_region}领先({ach_region_all[best_region]:.0%})，{worst_region}未达标', 'B3',
    colors=(BLUE,), labels=True, legend=False, horizontal=True, gap=50, num_fmt="0%")
ch.anchor = abs_anchor('ach')
from openpyxl.chart.series import DataPoint
from chart_style import GraphicalProperties, LineProperties
order = sorted(reg_names, key=lambda r: ach_region_all[r])
ch.series[0].data_points = [DataPoint(idx=i, spPr=GraphicalProperties(
    solidFill=(ORANGE if order[i] == best_region else ('C0504D' if ach_region_all[order[i]] < 1 else BLUE)),
    ln=LineProperties(noFill=True))) for i in range(len(order))]
ch.series[0].dLbls.numFmt = '0.0%'

# ── 图3: 品类环形 ──
ws_c = wb.create_sheet('品类源')
ws_c.append(['品类', '占比(%)'])
for name, share in cat_share:
    cc = ws_c.cell(row=ws_c.max_row+1, column=2, value=round(share, 1))
    cc.number_format = '0.0"%"'
ws_c.sheet_state = 'hidden'
DONUT_COLORS = [BLUE, ORANGE, LIGHT]
ch = styled_donut(
    ws_v,
    Reference(ws_c, min_col=2, min_row=1, max_col=2, max_row=1 + len(cat_share)),
    Reference(ws_c, min_col=1, min_row=2, max_row=1 + len(cat_share)),
    DONUT_COLORS[:len(cat_share)],
    f'品类占比：{cat_share[0][0]}{cat_share[0][1]:.0f}%居首', 'B3')
ch.anchor = abs_anchor('donut')

# ── 图4: 日销售额 vs 利润率 散点 ──
ws_s = wb.create_sheet('散点源')
ws_s.append(['日销售额(万元)', '利润率(%)'])
for k in sorted(daily.keys()):
    ws_s.append([round(daily[k][0], 2), daily[k][1]])
ws_s.sheet_state = 'hidden'
ch = styled_scatter(
    ws_v,
    Reference(ws_s, min_col=1, min_row=2, max_row=1 + len(daily)),
    Reference(ws_s, min_col=2, min_row=2, max_row=1 + len(daily)),
    f'日销售规模 vs 利润率（{len(daily)}天）: 规模越大利润率越薄', 'B3',
    x_title='日销售额(万元)', y_title='利润率(%)')
ch.anchor = abs_anchor('scatter')

# ══ Sheet4: 月×区域三色热力(达成率) ══
ws_h = wb.create_sheet('达成热力')
ws_h.append(['月份'] + reg_names)
for m in months:
    row = [m]
    for rn in reg_names:
        a, t = mon_reg[(m, rn)]
        row.append(round(a / t, 3))
    ws_h.append(row)
for r in range(2, 2 + len(months)):
    for c in range(2, 2 + len(reg_names)):
        ws_h.cell(row=r, column=c).number_format = '0.0%'
ws_h.conditional_formatting.add(f'B2:{chr(65+len(reg_names))}{1+len(months)}',
    ColorScaleRule(start_type='num', start_value=0.9, start_color='F4CCCC',
                   mid_type='num', mid_value=1.0, mid_color='FFF2CC',
                   end_type='num', end_value=1.08, end_color='D9EAD3'))
for c in range(1, 6):
    ws_h.cell(row=1, column=c).font = Font(bold=True, color=NAVY, name='微软雅黑')
for col, wd in zip('ABCDE', [8, 10, 10, 10, 10]): ws_h.column_dimensions[col].width = wd
ws_h['A7'] = '三色阶: 红<90% 黄~100% 绿≥108% (达成率)'
ws_h['A7'].font = Font(size=9, color='595959', name='微软雅黑')
fit_to_page(ws_h)

import sys as _sys, os as _os
_sys.path.insert(0, _os.path.dirname(_os.path.abspath(__file__)))
from chart_style import set_brand
set_brand(wb, title='2026 Q3 经营分析', subject='2026 Q3 经营分析 · 盛夏蓝橙 · 照镜自检')
wb.save(os.path.join(OUT, 'pipeline_q3.xlsx'))

# ── 一致性断言(供PPT端核对) ──
checks = {
    '总销售': round(TOTAL_A), '总目标': round(TOTAL_T), '达成率%': round(ACH * 100, 1),
    '最佳区域': best_region, '最佳区域达成%': round(ach_region_all[best_region] * 100, 1),
    '未达标区域': worst_region,
    '品类Top1': cat_share[0][0], '品类Top1占比%': round(cat_share[0][1], 1),
    '月度实际': [round(sum(mon_reg[(m, r)][0] for r in reg_names)) for m in months],
    '月度目标': [round(sum(mon_reg[(m, r)][1] for r in reg_names)) for m in months],
}
import json
json.dump(checks, open(os.path.join(OUT, 'consistency.json'), 'w'), ensure_ascii=False, indent=1)
print('✅ Excel管线生成完成:', wb.sheetnames)
print('✅ 一致性基准:', json.dumps(checks, ensure_ascii=False))
