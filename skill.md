---
name: econ-cn-docx
description: 经管实证论文中文学术 Word 排版。Python DSL (Paper 类) 一行命令生成 .docx。封装 OMML 公式引擎（分式、根号、上下标、希腊字母）、中文弯引号修复、字体三属性设置、esttab RTF 三线表、参考文献格式。
argument-hint: "<task> e.g. '生成 docx' / '排版论文' / '用 Paper DSL 写论文'"
user-invocable: true
---

# econ-cn-docx

## 快速开始

```bash
pip install -e ~/.dsh/skills/econ-cn-docx
```

```python
from cn_docx import Paper

p = Paper()
p.title("论文标题")
p.abstract("本文基于……")
p.keywords("关键词1；关键词2")

p.h1("一、引言")
p.para("正文第一段。\n\n正文第二段。")

p.h2("（一）基准回归")
p.equation(r"y_{it} = \beta x_{it} + \alpha_i + \lambda_t + \varepsilon_{it}")
p.para("回归结果如下。")
p.table("table_baseline_reg")
p.image("fig_trend.png")
p.h3("图一：趋势图")

p.h1("五、结论")
p.para("研究发现……")

p.h1("参考文献")
p.ref("[1] Author. Title[J]. Journal, Year, Vol(No): Pages.")

p.render("论文.docx")
```

## API 参考

### 元数据

| 方法 | 排版 |
|------|------|
| `title(text)` | 18pt 黑体居中加粗 |
| `abstract(text)` | 标签"摘要："黑体 10.5pt 加粗 + 楷体正文，顶格单倍行距 |
| `keywords(text)` | 标签"关键词："黑体 10.5pt 加粗 + 楷体正文 |

### 标题

| 方法 | 排版 |
|------|------|
| `h1(title, body="")` | 16pt 黑体居中加粗（一、引言 等） |
| `h2(title, body="")` | 14pt 黑体左对齐加粗（（一）基准回归 等） |
| `h3(title)` | 10.5pt 黑体居中加粗（图/表标题） |

### 正文

| 方法 | 排版 |
|------|------|
| `para(text)` | 12pt 宋体两端对齐，首行缩进 24pt，1.5 倍行距。`\n\n` 分多段 |
| `equation(latex)` | OMML 行间公式，居中 |
| `note(text)` | 9pt 宋体斜体（表注） |

### 浮动体

| 方法 | 排版 |
|------|------|
| `table(key)` | 解析 RTF → 三线表（顶线 12/底线 12/表头下线 8，单位 1/8 pt；自动采用 RTF 中的列宽） |
| `image(filename)` | 嵌入图片，居中，宽 10 cm |
| `image("fig.png:14")` | 嵌入图片，居中，宽 14 cm（`文件名:宽度cm` 语法） |

### 参考文献

| 方法 | 排版 |
|------|------|
| `ref(text)` | 10pt 宋体左对齐，1.25 倍行距 |

### 渲染

```python
p.render(output="论文.docx",
         tables_dir="tables", figures_dir="figures",
         table_map={"key": "file.rtf"},
         margin_left=3.18, body_size=12, heading1_size=16,
         # ... 全部 render() kwargs 可用
)
```

## OMML 公式引擎

支持以下 LaTeX 语法：

| 语法 | 效果 |
|------|------|
| `x_{sub}` | 下标 |
| `x^{sup}` | 上标 |
| `x_{sub}^{sup}` | 同时上下标 |
| `\frac{a}{b}` | 分式 |
| `\sqrt{x}` | 平方根 |
| `\sqrt[n]{x}` | n 次方根 |
| `\sum_{i=1}^{n}`、`\prod`、`\int` | 求和/连乘/积分，上下限显示在符号正上/正下方（m:nary） |
| `\bar{x}、\hat{x}、\tilde{x}、\dot{x}` | 组合变音符 |
| `\alpha、\beta、\gamma、\delta` 等 | 小写希腊字母 |
| `\Delta、\Sigma、\Phi` 等 | 大写希腊字母 |
| `\times、\cdot、\pm、\leq、\geq、\neq` 等 | 运算符 |
| `\rightarrow、\Rightarrow` | 箭头 |
| `\text{GDP}` | 公式内文本（正体，非斜体） |

### 公式能力边界（已知限制）

- 正文内联公式仅支持「字母 + 上下标」模式（如 `PM_{2.5}`、`RWI_{c}`）；含 `\times`、`\beta` 等命令的公式请用 `equation()`，写在正文里会打警告且显示为反斜杠原文
- `\text{...}` 内部不再递归解析（如 `\text{x_{i}}` 按字面渲染）
- 下标/上标内容不递归（如 `x^{10^3}` 按字面渲染）
- `\sum_{...}` 等命令参数内不支持嵌套花括号（如 `\{1,2\}` 会被截断）
- 连续单字符下标 `a_b_c` 有歧义，请用花括号形式 `a_{b}_{c}`

## 行内格式

- `H1`/`H2`/`H3`/`H4`/`H1a`/`H1b` 开头且后接标点或空格的段落 → 加粗正文（`H1N1`、`H2O` 等不误判）
- `word_{subscript}` 模式自动检测并渲染为内联 OMML 公式

## 图表标题约定（避免双标题）

- RTF 表文件自带标题时（esttab 输出第一行），`TABLE:` 渲染会自动插入居中黑体标题——**此时不要再为该表格配 `h3()`**，否则表格上方出现两个标题
- 图片没有内置标题，`IMAGE:` 后必须配 `h3()` 作为图题

## 字体体系

中文排版三属性设置（`w:eastAsia` / `w:hAnsi` / `w:ascii`）：

| 元素 | 字体 |
|------|------|
| 正文 | 宋体 / Times New Roman |
| 标题 | 黑体 / Times New Roman |
| 摘要/关键词 | 楷体 / Times New Roman |
| 表格 | 宋体 / Times New Roman（9pt） |

## 默认排版参数（毕业论文）

| 参数 | 默认值 |
|------|--------|
| 页边距 上/下/左/右 | 2.54 / 2.54 / 3.18 / 2.54 cm |
| 正文字号/行距 | 12pt / 1.5× |
| 首行缩进 | 24pt（≈两字符） |
| 一级标题 | 16pt |
| 二级标题 | 14pt |
| 图表标题 | 10.5pt |
| 参考文献 | 10pt / 1.25× |

所有排版参数通过 `render()` kwargs 覆盖。

## 排版预设（preset）

三套预设，`preset=` 一行切换；显式传入的参数覆盖预设值：

| 预设 | 场景 | 差异 |
|------|------|------|
| `preset="thesis"`（默认） | 毕业论文 | 页边距 2.54/2.54/3.18/2.54，正文 12pt/1.5×，一级标题 16pt |
| `preset="journal"` | 期刊投稿 | 左页边距 2.54，正文 10.5pt/缩进 21pt，一级标题 14pt，参考文献 9pt |
| `preset="working"` | 工作论文 | 左页边距 2.54，正文 12pt/1.5×，一级标题 15pt，二级 13pt |

```python
p.render("论文.docx", preset="journal")

# 预设基础上覆盖个别参数
p.render("论文.docx", preset="journal", body_size=11)
```

## 旧版兼容：JSON 入口（仍可用）

```python
from cn_docx import render
render("paper_content.json", output="论文.docx")
```

JSON 格式参见原文档，所有 `TABLE:key` / `IMAGE:file.png` / `EQUATION:...` / `REF:...` / `NOTE:...` 前缀仍然有效。

**注意**：JSON 入口不再做 LaTeX 转义恢复（`//cmd` 写法与损坏控制字符修复已移除，2026-08 清理）。公式请直接写 Unicode 字符（`β`、`×`、`α`）或双反斜杠（`\\beta`）；写单反斜杠 `\beta` 会被 JSON 解析成控制字符。

## 依赖

- python-docx >= 1.0
- lxml >= 4.9
