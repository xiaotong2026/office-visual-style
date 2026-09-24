# -*- coding: utf-8 -*-
"""图表美化demo: 生成 默认版(对比用) + 美化版(3张图)"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import openpyxl
from openpyxl.chart import BarChart, PieChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from chart_style import (styled_bar, styled_line, styled_pie, fit_to_page, BLUE)

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), '输出结果')
os.makedirs(OUT, exist_ok=True)

MONTHS = ['1月', '2月', '3月', '4月', '5月', '6月']
SALES  = [120, 135, 158, 142, 176, 190]
PROFIT = [45, 52, 61, 50, 70, 78]
PRODUCTS = [('智能硬件', 470), ('大家电', 200), ('厨卫电器', 140), ('配件耗材', 71), ('其他服务', 40)]

total_sales = sum(SALES)
peak = max(SALES)
top_prod, top_val = PRODUCTS[0]

def build_data(ws):
    ws.append(['月份', '销售额', '利润'])
    for m, s, p in zip(MONTHS, SALES, PROFIT):
        ws.append([m, s, p])
    ws.append([])
    ws.append(['产品线', '销售额'])
    for name, v in PRODUCTS:
        ws.append([name, v])

# ── 1. 默认版(对比基线) ──
wb1 = openpyxl.Workbook()
ws = wb1.active
build_data(ws)
bar = BarChart(); bar.type = 'col'; bar.title = '销售数据'
bar.add_data(Reference(ws, min_col=2, min_row=1, max_col=3, max_row=7), titles_from_data=True)
bar.set_categories(Reference(ws, min_col=1, min_row=2, max_row=7))
ws.add_chart(bar, 'E2')
wb1.save(os.path.join(OUT, '图表_默认版.xlsx'))

# ── 2. 美化版(3个sheet,各一张图) ──
wb2 = openpyxl.Workbook()
ws_data = wb2.active
ws_data.title = '数据'
build_data(ws_data)
ws_data.sheet_state = 'hidden'   # 隐藏数据页,渲染/打印只出图表

ws_bar = wb2.create_sheet('柱状图')
styled_bar(ws_bar,
           Reference(ws_data, min_col=2, min_row=1, max_col=3, max_row=7),
           Reference(ws_data, min_col=1, min_row=2, max_row=7),
           f'上半年累计销售{total_sales}万，6月创单月新高(万元)', 'B2',
           colors=(BLUE, '9AB8D8'))
fit_to_page(ws_bar)

ws_line = wb2.create_sheet('折线图')
styled_line(ws_line,
            Reference(ws_data, min_col=2, min_row=1, max_row=7),
            Reference(ws_data, min_col=1, min_row=2, max_row=7),
            '销售额逐月走高，6月达190万(万元)', 'B2')
fit_to_page(ws_line)

ws_pie = wb2.create_sheet('饼图')
styled_pie(ws_pie,
           Reference(ws_data, min_col=2, min_row=9, max_row=14),   # 行9=表头,10-14=5个产品线
           Reference(ws_data, min_col=1, min_row=10, max_row=14),
           f'{top_prod}贡献{top_val/total_sales:.0%}，是第一大收入来源', 'B2')
fit_to_page(ws_pie)

wb2.save(os.path.join(OUT, '图表_美化版.xlsx'))
print('生成完毕:', os.listdir(OUT))
print(f'校验: 累计={total_sales} 峰值={peak} 占比={top_val/total_sales:.1%}')
