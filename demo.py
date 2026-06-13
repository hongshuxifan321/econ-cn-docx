# -*- coding: utf-8 -*-
"""econ-cn-docx 完整演示 — Python DSL"""
import os, tempfile
from cn_docx import Paper

p = Paper()

# ── 元数据 ──
p.title("空气污染、远程工作可行性与城市经济产出")
p.abstract(
    "本文基于2000至2019年中国地级及以上城市面板数据，"
    "考察远程工作可行性（RWI）是否调节PM2.5污染对城市GDP的损害。"
    "结果表明：PM2.5浓度每上升1μg/m³，城市GDP下降约0.6%；"
    "但RWI每提高一个标准差，该边际损害减少约28%。"
)
p.keywords("PM2.5；远程工作可行性；职业结构；调节效应")

# ── 一、引言 ──
p.h1("一、引言")
p.para(
    "空气污染损害经济产出（Graff Zivin & Neidell, 2012）。"
    "但一个城市的产业结构决定了它受污染影响的程度——"
    "如果大量从业人员可以在家办公，污染对经济的冲击应当更小。"
)

# ── 二、研究设计 ──
p.h1("二、研究设计")
p.h2("（一）计量模型")
p.para("基准回归采用城市与年份双向固定效应模型：")
p.equation(
    r"lnGDP_{ct} = \beta_{1} PM_{ct} + \beta_{2} PM_{ct} \times RWI_{c}"
    r" + \gamma lnPop_{ct} + \alpha_{c} + \lambda_{t} + \varepsilon_{ct}"
)

# ── 三、实证结果 ──
p.h1("三、实证结果")
p.h2("（一）基准回归")
p.para(
    "表一报告了基准回归结果。PM2.5系数为-0.0060（p<0.001），"
    "PM2.5×RWI交互项系数为+0.0415（p=0.002），"
    "表明RWI越高的城市，PM2.5对GDP的边际损害越小。"
)
p.table("test_table")

p.h2("（二）公式压力测试")
p.para("分式+根号：")
p.equation(r"\sqrt{\frac{1}{n-1} \sum_{i=1}^{n} (x_{i} - \bar{x})^{2}}")
p.para("偏导数：")
p.equation(r"\frac{\partial y_{it}}{\partial PM_{ct}} = \beta_{1} + \beta_{2} \times RWI_{c}")
p.para("N次方根：")
p.equation(r"\sqrt[3]{x^{2} + y^{2}}")

# ── 表注 ──
p.note("注：标准误在城市层面聚类。*** p<0.01, ** p<0.05, * p<0.1。")

# ── 图表 ──
p.para("图一展示了测试图片。")
p.image("test_fig.png")
p.h3("图一：测试图片")

# ── 假说 ──
p.h2("（三）假说检验")
p.para("H1：PM2.5浓度对城市GDP存在显著的负向影响。")
p.para("H2：RWI对PM2.5的边际损害具有正向调节效应。")

# ── 参考文献 ──
p.h1("参考文献")
p.ref("[1] Dingel, J. I., Neiman, B. How many jobs can be done at home?[J]. Journal of Public Economics, 2020, 189: 104235.")
p.ref("[2] Graff Zivin, J., Neidell, M. The impact of pollution on worker productivity[J]. American Economic Review, 2012, 102(7): 3652-3673.")
p.ref("[3] He, J., Liu, H., Salvo, A. Severe air pollution and labor productivity[J]. American Economic Journal: Applied Economics, 2019, 11(1): 173-201.")

# ── 渲染 ──
out = os.path.join(tempfile.gettempdir(), "demo_output.docx")
p.render(out,
         tables_dir=os.path.join(os.path.dirname(__file__), "tests", "tables"),
         figures_dir=os.path.join(os.path.dirname(__file__), "tests", "figures"))
print(f"演示完成: {out}")
