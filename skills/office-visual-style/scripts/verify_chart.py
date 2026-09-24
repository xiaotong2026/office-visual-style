#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""图表交付自检 · 三层验证 (v1.7, 2026-09-24)

血泪来源: 一次复合图表返工7轮 —— 根因是"只验渲染不验数据",
脚本里两段代码覆盖同一数据区把数据写成乱码, PNG看着正常, 用户打开必错。

用法:
  python3 verify_chart.py structure 文件.xlsx        # ①结构层: 轴crossAx闭合/系列数
  python3 verify_chart.py dump     文件.xlsx 表名    # ②数据层: 导出单元格值逐行核对
  python3 verify_chart.py render   文件.xlsx         # ③渲染层: 逐页文字重叠+%标签计数
  python3 verify_chart.py all      文件.xlsx         # 结构 + 渲染(全部页)

退出码: 0=通过, 1=发现高优先级问题(轴未闭合等), 2=渲染告警(重叠等)
"""
import sys, os, re, subprocess, zipfile, tempfile, glob

RED = "\033[31m"; YEL = "\033[33m"; GRN = "\033[32m"; RST = "\033[0m"
def ok(m):  print(f"{GRN}✓ {m}{RST}")
def warn(m):print(f"{YEL}⚠ {m}{RST}")
def bad(m): print(f"{RED}✗ {m}{RST}")


def check_structure(xlsx):
    """① 结构层: 每个 chartN.xml 的轴 axId<->crossAx 是否成对闭合 + 系列数"""
    print(f"\n=== ① 结构层: {os.path.basename(xlsx)} ===")
    issues = 0
    with zipfile.ZipFile(xlsx) as z:
        charts = sorted(n for n in z.namelist() if re.search(r'charts/chart\d+\.xml$', n))
        if not charts:
            warn("无内嵌图表(可能图表是图片, 或路径不同)")
            return 0
        for n in charts:
            x = z.read(n).decode('utf-8', 'ignore')
            axes = {}
            for m in re.finditer(r'<(?:c:)?(catAx|valAx)>(.*?)</(?:c:)?\1>', x, re.S):
                kind = m.group(1)
                ids = re.findall(r'<(?:c:)?axId val="(\d+)"', m.group(2))
                crs = re.findall(r'<(?:c:)?crossAx val="(\d+)"', m.group(2))
                if ids:
                    axes[ids[0]] = (crs[0] if crs else None, kind)
            charts_in = re.findall(r'<(?:c:)?(barChart|lineChart|scatterChart|pieChart|areaChart|doughnutChart|bubbleChart)\b', x)
            closed = all((axes.get(c) and axes[c][0] == a) for a, (c, k) in axes.items())
            grp = {}
            for m in re.finditer(r'<(?:c:)?(barChart|lineChart|scatterChart)>(.*?)</(?:c:)?\1>', x, re.S):
                aid = re.findall(r'<(?:c:)?axId val="(\d+)"', m.group(2))
                grp.setdefault(tuple(aid), []).append(m.group(1))
            print(f"  {n}: 图型={charts_in} 轴对={ {a:(c,k) for a,(c,k) in axes.items()} }"
                  f" 系列组={ {k:len(v) for k,v in grp.items()} }")
            if not axes:
                warn(f"    {n} 无轴定义(饼图/环形图正常)")
            elif not closed:
                bad(f"    {n} 轴未成对闭合! Excel/WPS 可能判定图表损坏→删除→空白")
                issues += 1
            _bars = len([g for g in grp.values() if 'barChart' in g])
            if _bars >= 2:
                bad(f"    {n} 有 {_bars} 个 barChart 组: 必须每组独立轴对, "
                    f"共享轴对象会被 Excel/WPS 合并堆叠!")
                issues += 1
    if issues == 0:
        ok("结构层通过(轴全部闭合)")
    return issues


def check_data(xlsx, sheet):
    """② 数据层: 导出指定sheet全部单元格值, 供逐行核对(防'两段代码覆盖写坏')"""
    print(f"\n=== ② 数据层: sheet「{sheet}」全部单元格 ===")
    import openpyxl
    wb = openpyxl.load_workbook(xlsx, data_only=True)
    if sheet not in wb.sheetnames:
        bad(f"没有 sheet「{sheet}」, 现有: {wb.sheetnames}")
        return 1
    ws = wb[sheet]
    for r in range(1, ws.max_row + 1):
        vals = [ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
        if any(v is not None for v in vals):
            print(f"  R{r}: {vals}")
    warn("请人工核对: 图表数据区每行的值是否与源数据一致、行列是否错位/覆盖")
    return 0


def _pdftotext_boxes(pdf, page):
    out = subprocess.run(['pdftotext', '-f', str(page), '-l', str(page), '-bbox', pdf, '-'],
                         capture_output=True, text=True).stdout
    words = re.findall(r'<word xMin="([\d.]+)" yMin="([\d.]+)" xMax="([\d.]+)" yMax="([\d.]+)">([^<]+)</word>', out)
    return [(float(a), float(b), float(c), float(d), t) for a, b, c, d, t in words]


def check_render(xlsx, pages=None):
    """③ 渲染层: soffice→pdf→逐页检测文字框重叠 / 统计%标签"""
    print(f"\n=== ③ 渲染层: {os.path.basename(xlsx)} ===")
    if not subprocess.run(['which', 'soffice'], capture_output=True).stdout:
        warn("没装 soffice, 跳过渲染层")
        return 0
    tmp = tempfile.mkdtemp(prefix='verify_')
    subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', tmp, xlsx],
                   capture_output=True)
    pdfs = glob.glob(os.path.join(tmp, '*.pdf'))
    if not pdfs:
        bad("转PDF失败")
        return 1
    pdf = pdfs[0]
    npages = int(subprocess.run(['pdfinfo', pdf], capture_output=True, text=True)
                 .stdout.split('Pages:')[1].split()[0])
    warns = 0
    for p in range(1, npages + 1):
        boxes = _pdftotext_boxes(pdf, p)
        if not boxes:
            continue
        bad_pairs = []
        for i in range(len(boxes)):
            for j in range(i + 1, len(boxes)):
                x1, y1, x2, y2, t1 = boxes[i]; X1, Y1, X2, Y2, t2 = boxes[j]
                ix = max(0, min(x2, X2) - max(x1, X1)); iy = max(0, min(y2, Y2) - max(y1, Y1))
                ov = ix * iy
                small = min((x2 - x1) * (y2 - y1), (X2 - X1) * (Y2 - Y1))
                # 相对判据: 交叠须占较小文字框40%以上(滤掉句子内相邻词/表格正常紧排的误报)
                if small > 0 and ov > 0.40 * small and ov > 3:
                    bad_pairs.append((t1, t2))
        pct = [t for *_, t in boxes if t.endswith('%')]
        flag = f"{RED}重叠{len(bad_pairs)}{RST}" if bad_pairs else "零重叠"
        print(f"  P{p}: 文字块{len(boxes)} | {flag} | %标签{len(pct)} {bad_pairs[:4]}")
        if bad_pairs:
            warns += 1
    print(f"  预览PNG目录: {tmp} (render-*.png 可喂 image 工具; /tmp 下部分读图工具读不到, 需拷到工作目录)")
    return 2 if warns else 0


def main():
    if len(sys.argv) < 3:
        print(__doc__); sys.exit(0)
    cmd, xlsx = sys.argv[1], sys.argv[2]
    if cmd == 'structure':
        sys.exit(1 if check_structure(xlsx) else 0)
    elif cmd == 'dump':
        sys.exit(check_data(xlsx, sys.argv[3] if len(sys.argv) > 3 else '数据源'))
    elif cmd == 'render':
        sys.exit(check_render(xlsx))
    elif cmd == 'all':
        s = check_structure(xlsx); check_render(xlsx)
        sys.exit(1 if s else 0)
    else:
        print(__doc__); sys.exit(0)


if __name__ == '__main__':
    main()
