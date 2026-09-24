---
name: office-visual-style
description: Excel图表/PPT配色美化标准与渲染自检流程。生成柱状图/折线图/饼图/箱形图时用chart_style.py一行美化；交付前用LibreOffice渲染逐页过目。触发词：图表美化、配色、柱状图、饼图、箱形图、PPT美化、渲染验证。
metadata: { "openclaw": { "emoji": "🎨" } }
---

# Office Visual Style — 图表/PPT美化标准 v1.7

> **夏制 · Xiazhi** — 让 AI 出的图表，像人做的
> 配色系统「盛夏蓝橙」 · 自检闭环「照镜自检」

> 2026-09-23 沉淀于图表美化实战迭代。

## 何时用

- 用 openpyxl 生成带图表的 Excel（柱状/折线/饼/箱形图）
- 用 python-pptx 生成 PPT
- 交付任何带视觉元素的文档前自检

## 快速开始

脚本在本 skill 的 `scripts/` 目录（与SKILL.md同级，用相对路径解析）：

```python
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__))))  # 指向scripts目录
from chart_style import styled_bar, styled_line, styled_pie, styled_box, styled_combo, styled_waterfall, styled_stacked, fit_to_page

styled_bar(ws, data_ref, cats_ref, '洞见式标题(万元)', 'E2')   # 直接美化
fit_to_page(ws)                                                # 渲染/打印前必设
```

demo 参考：`scripts/demo_bar.py`（柱/饼）、`demo_box.py`（箱形图）、`demo_ppt.py`（PPT）

## 设计5原则

1. **洞见式标题**：说结论不说"统计"（"上半年累计销售921万，6月创单月新高" ✅，"销售数据" ❌）
2. **去图表垃圾**：去外边框、去竖网格线、去多余刻度；浅灰横网格 E9E9E9
3. **直接标注数值**：不让眼睛来回对轴
4. **低饱和配色**：主蓝 2E6FB7 + 强调橙 E8873A（只给最重要的一条）；饼图首位高亮+白描边
5. **层级统一**：标题深蓝1F3864加粗 > 轴标签灰595959 > 数据标签小字；图例一律置底；字体微软雅黑(Win)/Noto(Linux自动替换)

## 配色表

| 用途 | 色值 |
|------|------|
| 标题 NAVY | 1F3864 |
| 主色 BLUE | 2E6FB7 |
| 强调 ORANGE | E8873A |
| 次要系列 LIGHT | BDD7EE |
| 文字深/浅 | 262626 / 595959 |
| 网格线/轴线 | E9E9E9 / C9C9C9 |
| 饼图色带 | E8873A→2E6FB7→5B8FD0→88AFDF→B5CDEF |

## 渲染自检流程（交付前必做）

```bash
# 1. 生成时设 fit_to_page(ws)（防分页切割）
# 2. 转PDF再转PNG
soffice --headless --convert-to pdf --outdir /tmp/xT 文件.xlsx   # pptx同理
pdftoppm -png -r 90 /tmp/xT/文件.pdf /tmp/xT/render
# 3. 复制到workspace再给image tool（/tmp下读不到！）
cp /tmp/xT/render-1.png ./tmp/    # 拷到工作目录(部分读图工具读不到 /tmp)
# 4. 交付：把 PNG 附给用户
```

环境依赖：libreoffice-calc + libreoffice-impress + pdftoppm + 中文字体（`sudo apt-get install -y libreoffice-calc libreoffice-impress poppler-utils`）

## 踩坑清单（血泪）

- ⚠️ DataLabelList 必须**显式 show*=False** 关闭多余项，否则LibreOffice渲染出"月份;系列;数值"
- ⚠️ Reference 行号对齐数据：表头行作 min_row 配 titles_from_data，错一行→类别全错
- ⚠️ from_rows 模式系列引用**别把A列标签包进去**（含首列→值错位一格、尾部类别丢失）
- ⚠️ 箱形图：类名是 `ErrorBars` 不是 ErrorBar（openpyxl 3.1.5）；Excel2016+原生箱形图是chartEx新格式openpyxl写不了→用堆叠柱+误差线；DataLabel不支持自定义文本tx→**五数概括写进类别单元格**（\n换行+wrapText），箱体段showVal会标成宽度（错值）
- ⚠️ LibreOffice 渲染公式单元格可能不重算 → 统计值用Python预计算写死
- ⚠️ QQ传大文件：媒体上限 图片30MB/视频100MB/**文件100MB**/语音20MB（skill: qqbot-media），超限用 `split -b 90M` 拆分+收方 `cat part-* > 文件` 合并

## 视觉验证备胎（image模型超时时）

1. `pdftotext -f N -l N 文件.pdf -`：检查类别/标题/数字标签是否齐全（抓引用错位）
2. PIL像素分析：目标色像素聚类数=系列组数；箱体最低点远离横轴=透明基准生效；须线色像素y跨度=上下须正常
3. xlsx结构验证：unzip 后 grep chart1.xml（grouping/overlap/errBars/noFill），注意openpyxl写的是无前缀命名空间

## 与 theme-factory 的关系

theme-factory（Anthropic官方）管"选主题"——10套现成配色主题供用户挑；本 skill 管"程序化落地"——把颜色真正写进 openpyxl/python-pptx 代码并渲染验收。做正式汇报可先用 theme-factory 选主题，再用本 skill 的原则落地图表。


---

## v1.4 新增 (2026-09-23 深夜·管线压测)

### 新增函数 (chart_style.py)
| 函数 | 用途 | 要点 |
|------|------|------|
| `styled_target(ws, actual_ref, target_ref, cats_ref, ...)` | 实际vs目标 | 柱+虚线**同轴**(非次轴),目标线 prstDash='dash' |
| `styled_donut(ws, data_ref, cats_ref, colors, ...)` | 环形占比 | holeSize+DataPoint配色+白描边;标签%靠**源单元格number_format**继承(numFmt参数传None) |
| `styled_scatter(ws, x_ref, y_ref, ...)` | 散点相关 | 必须 `series.line.noFill` 否则连成折线 |
| `styled_bar(..., horizontal=True, num_fmt='#,##0')` | 横向排名 | 达成率类数据传 `num_fmt='0%'`,否则0.9~1.1被'#,##0'四舍五入成"1" |

### Excel→PPT 素材管线模式 (实战验证)
```
单一数据源 → Excel明细sheet(4位精度) → 所有聚合从明细派生
         ↓
   consistency.json 基准 ← Excel端算
         ↓
   PPT端读同一xlsx(openpyxl当素材) → 断言数字一致 → 原生pptx图表出图
```
- **叙述文字必须 f-string 实时计算**：禁止硬编码"低于目标5.0%""缺口14.9万"——数据重算后必漂移(实测抓到2处)
- PPT端断言示例：`assert abs(PPT总销售 - Excel总销售) <= 1.5`，6项指标全绿才交付
- 本管线首次跑通即抓出4个真bug(见下)，全部由"渲染验收+断言"抓出

### 新踩坑清单 v1.4
1. **python-pptx轴对象在chart级**(`ch.category_axis`/`ch.value_axis`)，`plot.category_axis`不存在(BarPlot报错)
2. **环形图无category_axis**：访问抛`ValueError`，getattr吞不掉→必须`try: getattr(ch,name) except`
3. **散点图plot无`_remove_dLbls`**：`plot.has_data_labels=False`会AttributeError→try包裹
4. **明细每行round(2)累计漂移**：1104行×±0.005≈±5.5万(实测目标997 vs 990)。修法：明细存4位+聚合全从明细派生(单一口径)
5. **达成系数被季节性稀释**：`ach**0.15`≈无效；且周内因子(wf均值1.05)×趋势(1.075)会掩盖区域差异→后处理按区域精确校准(scale=目标×ach/实际)
6. **绝对锚点几何必须算可打印宽**：LO竖版A4可打印≈19.7cm、横向≈28.4cm；右列图表 x+width 超标→被页边**裁切**(非换页!)
7. **LibreOffice不渲染base=0的堆叠柱**：瀑布图总柱(base=0)消失→垫0.5隐形底
8. **系列引用别含A列文本**：min_col从2起，否则幻影空点让全柱错位一格
9. **pkill/pgrep -f 模式含在自身命令行里→自杀**(实测连崩2次)：改用 `fuser -k PORT/tcp`
10. **数据故事自洽性校验**：区域达成率系数导致"两个未达标"却写"唯一未达标"→断言+渲染双查

### 验证三层备胎(本轮实战用量)
1. `image` 视觉模型逐页看(最准,但今日超时5+次)
2. `pdftotext -bbox` 拿文字精确坐标→验标签槽位/数值对齐
3. `PIL` 像素分类(容差22-26)按行/列分区统计色块→验配色逻辑(如横向条形"顶橙中蓝底红")


---

## v1.5 配色主题引擎 + 五行五色 (2026-09-23)

### 主题引擎(一键切换整套配色)
```python
use_theme('商务')   # 默认: 盛夏蓝橙
use_theme('五行')   # 中式: 五行五色
```
切换范围: 标题色 / 网格线 / 轴线 / 图表底色(五行=宣纸底)。函数: `use_theme(name)` / `set_theme(**kw)`。

### 五行五色 — 品牌子色板「五行五色」(低饱和演绎)
| 五行 | 色名 | HEX | 用途 |
|------|------|-----|------|
| 木 | 竹青 | `#6F8F7C` | 主系列 · 生长 |
| 火 | 赭石 | `#A2594B` | 强调/高亮 · 热烈 |
| 土 | 秋香 | `#C7A96A` | 次要系列 · 承载 |
| 金 | 月白 | `#AEB9BE` | 参考/目标 · 收敛 |
| 水 | 玄青 | `#37474F` | 标题/主色 · 深沉 |

中性: 宣纸底 `#FAF8F4` · 墨字 `#2B2B2B` · 网格 `#E7E2D9` · 轴 `#C9C3B6`
取色: `wuxing(n)` 返回木→火→土→金→水前n色。

**要点**: ①五行正色(朱砂/藤黄等)太艳 → 全部做**低饱和演绎**,耐看不艳丽 ②金(白)在白底不可见 → 图上以"月白"银灰承载 ③适用 PPT 与部分表格(需中式沉稳调性时),默认仍用商务「盛夏蓝橙」。
demo: `demo_wuxing.py`(色卡表+环形+柱状)。


---

## v1.6 五行主题 · PPT模板 (2026-09-23)

**模板脚本**: `scripts/demo_wuxing_ppt.py`（3页样例: 封面 + 柱状一色一系 + 环形）

```python
from chart_style import wuxing_rgb, wuxing_ppt_theme
prs = Presentation(...)
# ... 建页与图表, 系列色用 wuxing_rgb() ...
wuxing_ppt_theme(prs)      # 全部slide刷宣纸底 + 右下"◆ 夏制 · 五行五色"署名
prs.save('xxx.pptx')
```

- `wuxing_rgb()` → pptx用五行色RGBColor列表(木火土金水)
- `wuxing_ppt_theme(prs)` → 一键刷底+署名(须在建完页后调用)
- 封面要点: 宣纸底 + 五行竖色条 + 玄青大标题 + 赭石细线; 赭石只用在标题线/高亮


---

## v1.7 交付前必过「三层验证」+ 图表排查铁律 (2026-09-24)

> **血泪背景**：一次复合图表返工 **7 轮**。根因：脚本里**两段代码覆盖同一数据区**，把图表数据写成乱码；但我只验渲染（PNG 看着"正常"），从不验数据，用户用 WPS 打开必然错。期间还一路猜"Excel/WPS 渲染器兼容"，改了好几版结构——**方向全错**。

### 🔴 铁律一：交付 xlsx 必过三层，缺一层不交付

| 层 | 查什么 | 工具 |
|----|--------|------|
| ① 数据层 | **重新 load 文件，逐格核对图表数据区**与源数据是否一致（防重复写入/行列错位） | `verify_chart.py dump 文件.xlsx 表名` |
| ② 结构层 | 每个 chartN.xml 轴 `axId↔crossAx` 成对闭合；多 barChart 组是否各自独立轴对 | `verify_chart.py structure 文件.xlsx` |
| ③ 渲染层 | 转 PDF→PNG，逐页文字框重叠检测 + % 标签计数 | `verify_chart.py render 文件.xlsx` |

```bash
python3 scripts/verify_chart.py all 文件.xlsx     # 结构+渲染一键
python3 scripts/verify_chart.py dump 文件.xlsx 数据源  # 数据层(值需人工核对)
```

### 🔴 铁律二：图表/文件出问题，先查"自己产出的文件内容"，禁止先猜渲染器

用户反馈异常 → **第一步是 `load_workbook` 反查数据区**，不是猜"是不是 Excel/WPS 不兼容"。今天就是顺序反了，白改 5 版结构。

### 🔴 铁律三：同一张 sheet 只允许一处代码写同一数据区

段代码里出现两个 `for ... ws.cell(row=R0+ri ...)` 写同一片区域 = 定时炸弹。改版时**删干净旧写入块**，并对 `str.replace` 之类改法**必须 assert 命中**（今天还栽过一次"替换没匹配上但静默跳过"）。

### 关键坑位 v1.7

1. **`dispBlanksAs='span'`**（Excel"空单元格用直线连接"）：让隔行空值(=本月才有的值)的两条线**各自连续成一整条**。openpyxl 属性名是 **`ch.display_blanks`**（不是 `displayBlanksAs`！赋值不报错但静默无效）
2. **双 barChart 组并排**：必须**每组独立轴对**（catAx+valAx 成对闭合，第二组轴 delete=True+刻度手动同步）。**共享轴对象 → Excel/WPS 把两组当同一系列组全部堆叠**；LO 却当独立组并排渲染（两边行为相反）
3. **多轴 crossAx 必须显式闭合**：openpyxl 改 axId 时 crossAx 不联动 → 指向不存在的轴 → Excel/WPS 判损坏→删图表→空白
4. **单 barChart + DataPoint 逐点染色**做"月份双色柱"：比双 barChart 组稳（Excel/WPS/LO 都认），槽交错 + 大 gap 即视觉并排
5. **% 号**靠源单元格 `number_format='0.0"%"'` 继承（`DataLabelList.numFmt` 直赋字符串在 LO/Excel 不生效）
6. **用户验收环境可能是 WPS**（不一定是 Microsoft Excel）→ 自检以 WPS 能开为准；LO 渲染正常 ≠ WPS/Excel 正常

### 验证脚本相对判据（防误报）

文字重叠检测用**相对判据**：交叠面积 > 较小文字框的 40% 才算重叠（避免句子内相邻词、表格正常紧排被误判）。
