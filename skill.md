---
name: econ-cn-docx
description: 经管实证论文中文学术 Word 文档生成器。支持三种排版风格（毕业论文/期刊投稿/工作论文），封装 OMML 公式引擎、中文弯引号修复、字体三属性设置、三线表、GB/T 7714 参考文献格式等工具函数。参考了清华/上交/中科大论文模板、ElegantPaper 等项目。每次需要用 python-docx 生成符合中文学术排版标准的经管论文、开题报告或研究设计说明时使用。
argument-hint: "<task> e.g. '生成研究设计说明' or '把论文渲染成 Word' or '从 JSON 内容生成 docx'"
user-invocable: true
---
<!-- suppress-diagnostics -->

# 经管实证论文中文学术 Word 文档生成器

## 定位

经管实证论文的中文学术排版实现层。不做内容写作（`econ-write`），不做选题打磨（`thesis-idea`），只做一件事：**把内容变成符合标准的格式化 Word 文档**。

## 排版格式的依据

排版规范不是拍脑袋定的。三套预设各有依据：

| 风格 | 标识 | 依据 |
|------|------|------|
| **毕业论文** | `thesis` | GB/T 7713.1-2006《学位论文编写规则》+ 中南财/中财/海大等多校经管类毕业论文格式要求 |
| **期刊投稿** | `journal` | 《经济研究》《经济学（季刊）》《中国工业经济》等中文顶刊的投稿格式指南 |
| **工作论文** | `working_paper` | NBER 工作论文系列 + ElegantPaper LaTeX 模板的排版惯例 |

用户也可以传自定义 dict 覆盖参数。支持的键：`body_size` (Pt), `body_line_spacing` (倍数), `body_indent` (Pt), `heading1_size`/`heading2_size` (Pt), `heading1_align`/`heading2_align`, `ref_size` (Pt), `ref_line_spacing` (倍数), `table_size` (Pt), `margin_top`/`bottom`/`left`/`right` (Cm)。示例：`setup_page(doc, style={'body_size': Pt(11), 'margin_left': Cm(3)})`。

## 三种预设的差异

| 参数 | `thesis`（毕业论文） | `journal`（期刊投稿） | `working_paper`（工作论文） |
|------|---------------------|----------------------|---------------------------|
| 页边距 | 左 3.18, 其余 2.54 cm | 四边 2.54 cm | 四边 2.54 cm |
| 正文字号 | 12pt（小四） | 10.5pt（五号） | 11pt |
| 正文行距 | 1.5 倍 | 1.5 倍 | 1.5 倍 |
| 首行缩进 | Pt(24) ≈ 两字符 | Pt(21) ≈ 两字符 | Pt(22) |
| 一级标题 | 16pt 黑体居中 | 14pt 黑体居中 | 14pt 黑体左对齐 |
| 二级标题 | 14pt 黑体左对齐 | 12pt 黑体左对齐 | 12pt 黑体左对齐 |
| 参考文献字号 | 10pt | 9pt | 10pt |
| 参考文献行距 | 1.25 倍 | 1.25 倍 | 1.25 倍 |
| 表格字号 | 9pt | 8pt | 9pt |

## 与其他技能的关系

| 技能 | 负责 | 与本技能的关系 |
|------|------|---------------|
| `econ-write` | 论文内容写作（摘要/引言/结论/结果） | 上游——产出内容 |
| `academic-paper-writer` | 论文结构搭建、章节模板 | 上游——产出结构 |
| `lit-review-assistant` | 文献搜索、汇总、综述 | 上游——产出文献部分 |
| `latex-tables` | LaTeX 回归表格代码生成 | 互补——LaTeX 路径 |
| **`econ-cn-docx`** | **python-docx 中文学术 Word 排版** | **本技能——排版实现** |

## 涉及的社区参考

| 来源 | 类型 | 参考内容 |
|------|------|----------|
| `tuna/thuthesis` (5376★) | 清华大学 LaTeX 论文模板 | 页面边距、标题层级、图表编号规范 |
| `sjtug/SJTUThesis` (3787★) | 上海交大 LaTeX 论文模板 | 章节结构、参考文献编号体系 |
| `ustctug/ustcthesis` (2103★) | 中科大 LaTeX 论文模板 | 字体体系、正文排版密度 |
| `ElegantLaTeX/ElegantPaper` (1450★) | 工作论文 LaTeX 模板 | NBER 风格布局 |
| `darksider-9/master-thesis-studio-skill` (104★) | 东南大学 Word 自动生成 Skill | Codex Skill 模式参考 |
| `zouchenzhen/docx-template-translator-skill` (36★) | LaTeX→Word 模板转换 Skill | 格式映射思路 |
| GB/T 7714-2015 | 国家标准 | 参考文献著录规则 |
| GB/T 7713.1-2006 | 国家标准 | 学位论文编写规则 |
| 《经济研究》《经济学（季刊）》等 | 期刊规范 | 投稿格式要求 |
| 中南财经政法大学 / 中央财经大学等 | 高校规范 | 毕业论文具体格式 |

## 解决的问题

1. **中文弯引号**：源码中的中文双引号被存为 ASCII 直引号（U+0022），在 Word 中显示为 Times New Roman 而非中文字体。`fix_chinese_quotes()` 配对替换。
2. **LaTeX 公式**：行内和行间方程需转为 Word 原生 OMML 数学方程（双击可编辑）。引擎支持下标、上标、上下标组合、希腊字母、数学符号、求和/积分。
3. **中文字体三属性**：`set_cn_font` 必须同时设置 `w:ascii=TNR`、`w:hAnsi=中文字体`、`w:eastAsia=中文字体`。
4. **参考文献格式**：GB/T 7714-2015 要求作者 3 人以下全列、3 人以上加"等"/"et al."，文献类型标识（[J]/[M]/[D]等），中文在前英文在后。

## 核心工具函数

所有函数在 `references/toolkit.py`。

### 风格选择

```python
from cn_docx_toolkit import *

doc = Document()
setup_page(doc, style='thesis')   # 'thesis' | 'journal' | 'working_paper'
# 或自定义
setup_page(doc, style='custom', body_size=Pt(11), body_line_spacing=1.5)
```

所有高级排版函数（add_title、add_body 等）自动使用 `setup_page` 设置的风格参数，无需每个函数单独传。

### 高级排版 API

| 函数 | `thesis` 规格 |
|------|--------------|
| `add_title(doc, text)` | 18pt 黑体居中加粗 |
| `add_heading_1(doc, text)` | 16pt 黑体居中加粗 |
| `add_heading_2(doc, text)` | 14pt 黑体左对齐加粗 |
| `add_heading_3(doc, text)` | 10.5pt 黑体居中加粗 |
| `add_body(doc, text, indent=True)` | 12pt 宋体两端对齐，首行缩进，1.5 倍行距 |
| `add_bullet(doc, text)` | 12pt 宋体，悬挂缩进 |
| `add_abstract(doc, body)` | 10.5pt 黑体加粗标签 + 楷体正文 |
| `add_keywords(doc, keywords)` | 10.5pt 黑体加粗标签 + 楷体内容 |
| `add_reference(doc, text)` | 10pt 宋体左对齐，1.25 倍行距 |
| `add_table_note(doc, text)` | 9pt 宋体左对齐 |
| `add_display_equation(doc, latex_str)` | 居中 OMML 方程 |
| `add_three_line_table(doc, headers, rows)` | 三线表（顶线12/表头下线8/底线12） |
| `load_template(path)` | 从预设 docx 模板创建文档（继承样式和多级编号） |
| `add_figure(doc, image, caption, width)` | 图片嵌入，居中 + 图题 |
| `add_page_number(doc)` | 页脚居中页码（PAGE 域代码） |
| `add_numbered_heading(doc, text, level)` | 多级标题自动编号（1, 1.1, 1.1.1） |
| `reset_heading_counters()` | 重置标题计数器 |
| `add_chart(doc, fig, caption, width)` | matplotlib 图表嵌入（自动转 PNG） |
| `render_template(template, context, output)` | docxtpl 模板填充（Jinja2 语法） |
| `add_regression_table(doc, models, var_names, notes)` | 回归结果专用表（系数/标准误/星号/N/R²） |
| `add_cover_page(doc, title, ...)` | 毕业论文封面（学校/标题/作者/日期） |
| `add_abstract_and_keywords(doc, body, kw)` | 摘要 + 关键词一键生成 |
| `reset_heading_counters(doc)` | 重置该文档的标题计数器 |

### 底层引擎

| 函数 | 作用 |
|------|------|
| `set_cn_font(run, font_name, size, bold)` | 中文字体三属性设置 |
| `fix_chinese_quotes(text)` | ASCII 直引号→中文弯引号 |
| `parse_inline_text(text, paragraph, size)` | 解析行内 `$...$` 公式和 `**粗体**` |
| `insert_inline_omath(paragraph, latex_str)` | 插入行内 OMML 公式 |

## 典型用法

```python
from docx import Document
from cn_docx_toolkit import *

doc = Document()
setup_page(doc, style='journal')  # 按《经济研究》格式

add_title(doc, '论文标题')
add_abstract(doc, '本文基于...')
add_keywords(doc, '关键词1；关键词2')
add_heading_1(doc, '一、引言')
add_body(doc, '正文段落，包含 $\\beta_1$ 行内公式。')
add_display_equation(doc, r'y_{it} = \\beta_1 x_{it} + \\alpha_i + \\lambda_t + \\varepsilon_{it}')
add_heading_2(doc, '（一）基准回归')
add_three_line_table(doc,
    ['变量', '(1) OLS', '(2) FE'],
    [['$x$', '0.123***', '0.098**'],
     ['$z$', '0.045', '0.067*']])
add_table_note(doc, '注：括号内为聚类标准误。*** p<0.01, ** p<0.05, * p<0.1')
add_reference(doc, '[1] 作者. 题名[J]. 期刊名, 年, 卷(期): 页码.')

doc.save('output.docx')
```

## 内容排版分离模式

适合论文正文。JSON 定义内容，排版脚本遍历渲染：

```json
[
  {"title": "论文标题", "body": "", "level": 0},
  {"title": "摘要", "body": "本文基于...", "level": 1},
  {"title": "一、引言", "body": "第一段\\n\\n第二段", "level": 1},
  {"title": "", "body": "EQUATION: y = \\beta x + \\varepsilon", "level": 0},
  {"title": "表1: 描述性统计", "body": "TABLE: desc_stats", "level": 3}
]
```

- `level` 0 = 特殊元素（标题、公式、表格标记）
- `level` 1 = 一级标题
- `level` 2 = 二级标题 + 正文段落
- `level` 3 = 表/图标题
- `body` 以 `EQUATION:` / `TABLE:` 开头触发特殊渲染

## 已知局限与升级路径

| 局限 | 当前 | 升级 |
|------|------|------|
| 不支持 \frac, \sqrt, 矩阵 | 手动 OMML XML | 安装 pandoc，texmath 批量转换 |
| fix_chinese_quotes 基于启发式 | 引号内含中文才替换 | 纯英文引号手动查 |
| \ln/\log 函数名斜体 | italic | 加 m:sty 直立 |
| 仅 Word，不支持 LaTeX | python-docx | 用 thuthesis/SJTUThesis |
| 高级页眉/目录 | 暂不支持 TOC 域代码和复杂页眉 | 页脚页码已支持（add_page_number），页眉用 section.header |
