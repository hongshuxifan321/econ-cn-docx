# -*- coding: utf-8 -*-
"""econ-cn-docx 完整演示 — Python DSL。

跑法：
    python demo.py                     # 通用版式（基于《经济研究》投稿体例）
    python demo.py path/to/other.yaml  # 换一份规格（首行 extends: base）
"""
import os
import sys
import tempfile

from cnpaper import Paper

spec = sys.argv[1] if len(sys.argv) > 1 else "base"
here = os.path.dirname(os.path.abspath(__file__))

p = Paper()

# ── 标题块 ──
p.title("空气污染、远程工作可行性与城市经济产出")
p.footnote(
    "张三，某某大学经济学院，邮政编码：100000，电子信箱：zhangsan@example.edu.cn。"
    "本文得到国家自然科学基金（编号 00000000）的资助。作者感谢匿名审稿人的建议。"
)
p.subtitle("——来自中国地级市面板数据的证据")
p.author("张　三　李　四")
p.abstract(
    "本文基于 2000—2019 年中国地级及以上城市面板数据，考察远程工作可行性（RWI）"
    "是否调节 PM2.5 污染对城市 GDP 的损害。结果表明：PM2.5 浓度每上升 1μg/m³，"
    "城市 GDP 下降约 0.6%；但 RWI 每提高一个标准差，该边际损害减少约 28%。"
)
p.keywords("PM2.5　远程工作可行性　职业结构　调节效应")

# ── 正文 ──
p.h1("引　言")
p.para(
    "空气污染损害经济产出{fn:Graff Zivin & Neidell (2012) 给出了较早的因果证据。}。"
    "但一个城市的产业结构决定了它受污染影响的程度——如果大量从业人员可以在家办公，"
    "污染对经济的冲击应当更小。"
)

p.h1("研究设计")
p.h2("计量模型")
p.para("基准回归采用城市与年份双向固定效应模型：")
p.equation(
    r"lnGDP_{ct} = \beta_{1} PM_{ct} + \beta_{2} PM_{ct} + \gamma lnPop_{ct}"
    r" + \alpha_{c} + \lambda_{t} + \varepsilon_{ct}"
)

p.h2("公式压力测试")
p.para("分式 + 根号：")
p.equation(r"\sqrt{\frac{1}{n-1} \sum_{i=1}^{n} (x_{i} - \bar{x})^{2}}")
p.para("偏导数：")
p.equation(r"\frac{\partial y_{it}}{\partial PM_{ct}} = \beta_{1} + \beta_{2} \times RWI_{c}")
p.para("N 次方根：")
p.equation(r"\sqrt[3]{x^{2} + y^{2}}")

p.para("行内公式 PM_{2.5} 与 RWI_{c} 会自动渲染为 OMML。")

# ── 结果 ──
p.h1("实证结果")
p.h2("基准回归")
p.para("表 1 报告了基准回归结果，交互项系数显著为正。")
p.table("test_table", caption="基准回归结果")
p.note("注：标准误在城市层面聚类。*** p<0.01，** p<0.05，* p<0.1。")

p.figure("test_fig.png", caption="PM2.5 与远程工作可行性的趋势（示意）", width=10)
p.para("图 1 显示两组城市的走势差异并不明显。")

p.h2("假说检验")
p.para("H1：PM2.5 浓度对城市 GDP 存在显著的负向影响。")
p.para("H2：RWI 对 PM2.5 的边际损害具有正向调节效应。")

# ── 参考文献 ──
p.h1("参考文献")
p.ref("Dingel, J. I., and B. Neiman, 2020, “How Many Jobs Can Be Done at Home?”, "
      "{i:Journal of Public Economics}, 189, 104235.")
p.ref("Graff Zivin, J., and M. Neidell, 2012, “The Impact of Pollution on Worker "
      "Productivity”, {i:American Economic Review}, 102(7), 3652-3673.")
p.ref("He, J., H. Liu, and A. Salvo, 2019, “Severe Air Pollution and Labor "
      "Productivity”, {i:American Economic Journal: Applied Economics}, 11(1), 173-201.")

# ── 渲染 ──
out = os.path.join(tempfile.gettempdir(), "demo_output_%s.docx" % spec)
p.render(
    out,
    spec=spec,
    tables_dir=os.path.join(here, "tests", "tables"),
    figures_dir=os.path.join(here, "tests", "figures"),
)
print("演示完成: %s（spec=%s）" % (out, spec))
