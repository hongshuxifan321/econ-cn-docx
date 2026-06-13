# -*- coding: utf-8 -*-
"""End-to-end test for cn_docx Python DSL."""
import os, sys

ROOT = os.path.dirname(os.path.abspath(__file__))

from cn_docx import Paper

p = Paper()

# Metadata
p.title("测试论文 Python DSL 端到端验证")
p.abstract(
    "这是一段摘要，测试楷体渲染和弯引号——"
    "所谓“远程工作可行性”的“经济损害缓冲效应”。"
)
p.keywords("PM2.5；远程工作；调节效应")

# H1 + body
p.h1("一、引言")
p.para(
    "第一段正文。测试宋体两端对齐、首行缩进，以及内联公式 PM_{2.5} 的渲染效果。\n\n"
    "第二段正文。测试弯引号修复——"
    "“已有研究”发现“污染损害”存在异质性。"
)

# H2 + equation
p.h2("（一）基准回归")
p.equation(r"lnGDP_{ct} = \beta_{1} PM_{ct} + \beta_{2} PM_{ct} \times RWI_{c} + \alpha_{c} + \lambda_{t} + \varepsilon_{ct}")
p.para("回归结果如下。交互项系数显著为正，支持H2缓冲效应假说。")

# Table with RTF
p.para("表一报告了测试表格。")
p.table("test_table")

# Image
p.para("图一展示了测试图片。")
p.image("test_fig.png")
p.h3("图一：测试图片")

# Mechanism equation with frac
p.h2("（二）机制检验")
p.equation(r"lnFirms_{ct} = \beta_{1} PM_{ct} + \beta_{2} PM_{ct} \times RWI_{c} + \gamma X_{ct} + \alpha_{c} + \lambda_{t} + \varepsilon_{ct}")

# Bold hypothesis line
p.para("H1：PM2.5浓度对城市GDP存在显著的负向影响。")
p.para("H2：RWI对PM2.5的边际损害具有正向调节效应。")

# Note
p.note("注：标准误在城市层面聚类，括号内为t统计量。*** p<0.01，** p<0.05，* p<0.1。")

# References
p.h1("参考文献")
p.ref("[1] Dingel, J. I., Neiman, B. How many jobs can be done at home?[J]. Journal of Public Economics, 2020, 189: 104235.")
p.ref("[2] 陈诗一, 陈登科. 雾霾污染、政府治理与经济高质量发展[J]. 经济研究, 2018, 53(2): 20-34.")
p.ref("[3] Fu, S., Viard, V. B., Zhang, P. Air pollution and manufacturing firm productivity[J]. Economic Journal, 2021, 131(640): 3241-3273.")

# OMML stress test: frac + sqrt + sub + sup + Greek
p.h2("公式引擎压力测试")
p.para("行间分式+根号：")
p.equation(r"\sqrt{\frac{1}{n-1} \sum_{i=1}^{n} (x_{i} - \bar{x})^{2}}")
p.para("带上下标的分式：")
p.equation(r"\frac{\partial y_{it}}{\partial PM_{ct}} = \beta_{1} + \beta_{2} \times RWI_{c}")
p.para("N次方根：")
p.equation(r"\sqrt[3]{x^{2} + y^{2}}")

# Render
out = os.path.join(ROOT, "测试_output.docx")
p.render(out,
         tables_dir=os.path.join(ROOT, "tables"),
         figures_dir=os.path.join(ROOT, "figures"))

size_kb = os.path.getsize(out) / 1024
print(f"OK - {out} ({size_kb:.1f} KB)")
