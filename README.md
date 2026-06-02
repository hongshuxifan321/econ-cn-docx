# econ-cn-docx

用 python-docx 生成符合中文学术排版标准的经管实证论文 Word 文档。

## 特性

- **双路径**：纯 Python 代码构建 + docxtpl Jinja2 模板填充，按需选用
- **三种预设风格**：毕业论文（GB/T 7713）、期刊投稿（《经济研究》格式）、工作论文（NBER 风格）
- **Word 原生公式**：LaTeX 语法 → OMML 数学方程，支持上下标、希腊字母、求和积分
- **自动引号修复**：ASCII 直引号 → 中文弯引号（U+201C/U+201D）
- **三线表 + 回归表**：经管论文标准三线表，专用回归结果格式化
- **多级编号**：1 → 1.1 → 1.1.1 自动递增，计数器按文档隔离
- **封面 / 摘要 / 关键词 / 参考文献 / 页码**：开箱即用

## 安装

```bash
pip install -r requirements.txt
```

依赖：`python-docx`、`lxml`、`docxtpl`（可选，仅模板填充路径需要）。

## 快速开始

```python
from docx import Document
from references.toolkit import *

doc = Document()
setup_page(doc, style='thesis')  # 'thesis' | 'journal' | 'working_paper'

add_title(doc, '论文标题')
add_abstract_and_keywords(doc, '本文基于...', '关键词1；关键词2')

add_numbered_heading(doc, '引言', 1)
add_body(doc, '正文段落，支持 $\\beta_1$ 行内公式和 **粗体**。')
add_display_equation(doc, r'y_{it} = \\beta_1 x_{it} + \\alpha_i + \\lambda_t')

add_numbered_heading(doc, '基准回归', 2)
add_regression_table(doc, [
    {'name': '(1) FE', 'N': 5000, 'R2': 0.123,
     'coef': {'x': 0.123, 'z': -0.045},
     'se':   {'x': 0.023, 'z': 0.018},
     'pvalues': {'x': 0.000, 'z': 0.031}},
])
add_page_number(doc)
doc.save('output.docx')
```

运行 `python demo.py` 查看完整示例。

## 完整 API

### 页面与风格

| 函数 | 说明 |
|------|------|
| `setup_page(doc, style)` | 三种预设 + 自定义 dict |
| `set_style(name)` | 运行时切换风格 |
| `load_template(path)` | 从 Word 模板创建文档 |

### 标题与正文

| 函数 | 排版 |
|------|------|
| `add_title(doc, text)` | 18pt 黑体居中 |
| `add_heading_1(doc, text)` | 风格决定字号对齐 |
| `add_heading_2(doc, text)` | 风格决定字号对齐 |
| `add_heading_3(doc, text)` | 10.5pt 黑体居中 |
| `add_numbered_heading(doc, text, level)` | 自动编号 |
| `reset_heading_counters(doc)` | 重置计数器 |
| `add_body(doc, text, indent=True)` | 宋体两端对齐 |
| `add_bullet(doc, text)` | 悬挂缩进 |

### 摘要与参考文献

| 函数 | 说明 |
|------|------|
| `add_abstract(doc, body)` | 黑体标签 + 楷体正文 |
| `add_keywords(doc, keywords)` | 黑体标签 + 楷体内容 |
| `add_abstract_and_keywords(doc, body, kw)` | 两步合一 |
| `add_reference(doc, text)` | 宋体 1.25 倍行距 |

### 封面与页码

| 函数 | 说明 |
|------|------|
| `add_cover_page(doc, title, school, author, ...)` | 毕业论文封面 |
| `add_page_number(doc, show_total=True)` | "第 X 页 / 共 Y 页" |

### 公式

| 函数 | 说明 |
|------|------|
| `add_display_equation(doc, latex_str)` | 居中 OMML 行间方程 |
| `parse_inline_text(text, paragraph, size)` | 解析行内 `$...$` 和 `**粗体**` |

### 表格

| 函数 | 说明 |
|------|------|
| `add_three_line_table(doc, headers, rows)` | 三线表（顶线12/下线8/底线12） |
| `add_regression_table(doc, models, ...)` | 回归结果专用（系数/标准误/星号/N/R²） |
| `add_table_note(doc, text)` | 表注 |

### 图表

| 函数 | 说明 |
|------|------|
| `add_figure(doc, image, caption, width)` | 图片嵌入 + 图题 |
| `add_chart(doc, fig, caption, width)` | matplotlib 图表嵌入 |

### 模板

| 函数 | 说明 |
|------|------|
| `render_template(template, context, output)` | docxtpl 模板填充 |

### 底层工具

| 函数 | 说明 |
|------|------|
| `set_cn_font(run, font_name, size, bold)` | 中文字体三属性设置 |
| `fix_chinese_quotes(text)` | ASCII 直引号 → 中文弯引号 |
| `insert_inline_omath(paragraph, latex_str)` | 插入行内 OMML 元素 |

## 风格预设对照

| 参数 | `thesis` | `journal` | `working_paper` |
|------|----------|-----------|-----------------|
| 页边距 | 左 3.18, 其余 2.54 cm | 四边 2.54 cm | 四边 2.54 cm |
| 正文字号 | 12pt | 10.5pt | 11pt |
| 正文行距 | 1.5 | 1.5 | 1.5 |
| 一级标题 | 16pt 居中 | 14pt 居中 | 14pt 左对齐 |
| 二级标题 | 14pt 左对齐 | 12pt 左对齐 | 12pt 左对齐 |
| 参考文献 | 10pt | 9pt | 10pt |
| 表格 | 9pt | 8pt | 9pt |

## 设计依据

- GB/T 7713.1-2006《学位论文编写规则》
- GB/T 7714-2015《信息与文献 参考文献著录规则》
- 《经济研究》《经济学（季刊）》投稿格式指南
- thuthesis / SJTUThesis / ustcthesis 等 LaTeX 模板的排版惯例
- ElegantPaper LaTeX 工作论文模板

## 局限

- 不支持 `\frac`、`\sqrt`、矩阵等复杂 LaTeX 公式（可用 pandoc 批量转换或切换到 LaTeX 工作流）
- `\ln`、`\log` 等函数名在 OMML 中显示为斜体而非直立
- 无 TOC 目录自动生成、无图表交叉引用

## 许可证

MIT
