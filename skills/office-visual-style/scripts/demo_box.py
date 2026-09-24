# -*- coding: utf-8 -*-
"""箱形图demo: 四门店日营业额分布(堆叠柱+误差线原生方案)"""
import sys, os, random, statistics
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import openpyxl
from openpyxl.chart import Reference
from chart_style import styled_box, fit_to_page

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), '输出结果')
os.makedirs(OUT, exist_ok=True)

GROUPS = ['A门店', 'B门店', 'C门店', 'D门店']
random.seed(20260923)
RAW = {
    'A门店': [round(random.gauss(12, 1.5), 1) for _ in range(20)],
    'B门店': [round(random.gauss(10, 2.5), 1) for _ in range(20)],
    'C门店': [round(random.gauss(14, 3.0), 1) for _ in range(20)],
    'D门店': [round(random.gauss(11, 1.0), 1) for _ in range(20)],
}

def stats(vals):
    q1, q2, q3 = statistics.quantiles(vals, n=4, method='inclusive')  # =QUARTILE.INC
    return dict(min=min(vals), q1=q1, med=q2, q3=q3, max=max(vals),
                mean=statistics.mean(vals))

S = {g: stats(v) for g, v in RAW.items()}
top_med = max(GROUPS, key=lambda g: S[g]['med'])
top_iqr = max(GROUPS, key=lambda g: S[g]['q3'] - S[g]['q1'])

wb = openpyxl.Workbook()
ws = wb.active
ws.title = '数据'

# 原始数据块
ws.append(['门店'] + GROUPS)
for i in range(20):
    ws.append([f'第{i+1}天'] + [RAW[g][i] for g in GROUPS])
ws.append([])
# 统计表(行号: 23起)
ws.append(['统计量(箱形图用)'] + [''] * 4)
row = {name: 24 + i for i, name in enumerate(
    ['最小值', 'Q1', '中位数', 'Q3', '最大值', '均值',
     '下须宽(Q1-最小)', '上须宽(最大-Q3)', '箱体下半(中位-Q1)', '箱体上半(Q3-中位)'])}
ws.append(['统计量'] + GROUPS)          # 行23
for name in ['最小值', 'Q1', '中位数', 'Q3', '最大值', '均值']:
    ws.append([name] + [round(S[g][{'最小值': 'min', 'Q1': 'q1', '中位数': 'med',
                                    'Q3': 'q3', '最大值': 'max', '均值': 'mean'}[name]], 2) for g in GROUPS])
ws.append(['下须宽(Q1-最小)'] + [round(S[g]['q1'] - S[g]['min'], 2) for g in GROUPS])
ws.append(['上须宽(最大-Q3)'] + [round(S[g]['max'] - S[g]['q3'], 2) for g in GROUPS])
ws.append(['箱体下半(中位-Q1)'] + [round(S[g]['med'] - S[g]['q1'], 2) for g in GROUPS])
ws.append(['箱体上半(Q3-中位)'] + [round(S[g]['q3'] - S[g]['med'], 2) for g in GROUPS])
for r in range(24, 34):
    for c in range(2, 6):
        ws.cell(row=r, column=c).number_format = '0.0'
ws.cell(row=34, column=1).value = '注: 统计值已按QUARTILE.INC口径预计算;如需随原始数据联动,可换成=QUARTILE.INC()公式'

# 图表页
ws_c = wb.create_sheet('箱形图')

# 类别标签带五数概括(openpyxl的DataLabel不支持自定义文本,箱体段标值会变成宽度 → 把真实四分位写在类别里)
from openpyxl.styles import Alignment
LABEL_ROW = 36
ws['A36'] = '图表类别标签(自动生成,含五数)'
for i, g in enumerate(GROUPS):
    c = ws.cell(row=LABEL_ROW, column=2 + i)
    c.value = f"{g}\n中位 {S[g]['med']:.1f}\nQ1 {S[g]['q1']:.1f} ~ Q3 {S[g]['q3']:.1f}"
    c.alignment = Alignment(wrapText=True, horizontal='center')

styled_box(
    ws_c,
    Reference(ws, min_col=2, min_row=LABEL_ROW, max_col=5, max_row=LABEL_ROW),   # 类别=组名+五数
    (Reference(ws, min_col=2, min_row=row['Q1'], max_col=5, max_row=row['Q1']),
     Reference(ws, min_col=2, min_row=row['箱体下半(中位-Q1)'], max_col=5, max_row=row['箱体下半(中位-Q1)']),
     Reference(ws, min_col=2, min_row=row['箱体上半(Q3-中位)'], max_col=5, max_row=row['箱体上半(Q3-中位)'])),
    f'四门店日营业额分布：{top_med}中位数最高({S[top_med]["med"]:.1f}万)，{top_iqr}波动最大',
    'B2',
    err_refs=(Reference(ws, min_col=2, min_row=row['下须宽(Q1-最小)'], max_col=5, max_row=row['下须宽(Q1-最小)']),
              Reference(ws, min_col=2, min_row=row['上须宽(最大-Q3)'], max_col=5, max_row=row['上须宽(最大-Q3)'])),
)
ws_c['B27'] = '解读：箱体分色边界=中位数 | 须=Min~Max | 每组下方=中位数和Q1~Q3区间(改数据自动更新)'
ws_c['B27'].font = openpyxl.styles.Font(color='595959', size=10)
fit_to_page(ws_c)

import sys as _sys, os as _os
_sys.path.insert(0, _os.path.dirname(_os.path.abspath(__file__)))
from chart_style import set_brand
set_brand(wb, title='夏制 · 箱形图', subject='夏制 · 箱形图 · 盛夏蓝橙 · 照镜自检')
wb.save(os.path.join(OUT, '箱形图_美化版.xlsx'))
print('行号映射:', row)
print(f'校验: {top_med}中位数={S[top_med]["med"]:.2f} | {top_iqr}IQR={S[top_iqr]["q3"]-S[top_iqr]["q1"]:.2f}')
