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
| `keywords(text)` | 标签"关键词："（字体/字号见 spec）+ 正文 |
| `subtitle(text)` | 副标题「——…」，居中（见 `subtitle_font/size`） |
| `author(text)` | 作者姓名行，居中（见 `author_font/size`） |
| `footnote(text)` | 给**上一段**挂脚注：页下注、每页重新编号（首页作者信息就挂在标题上）；标记样式见 spec 的 `footnote_num_format`（`chicago` → `*`、`†`、`‡`） |

### 标题

| 方法 | 排版 |
|------|------|
| `h1(title, body="")` | 16pt 黑体居中加粗（一、引言 等） |
| `h2(title, body="")` | 14pt 黑体左对齐加粗（（一）基准回归 等） |
| `h3(title)` | 10.5pt 黑体居中加粗（图/表标题） |

### 正文

| 方法 | 排版 |
|------|------|
| `para(text)` | 正文段落，两端对齐、首行缩进（样式见 spec） |
| `para_no_indent(text)` | 正文段落，不缩进 |
| `equation(latex)` | OMML 行间公式，居中；`equation_number: true` 时右端自动编号 (1)(2)… |
| `note(text)` | 表注行（默认 9pt 宋体斜体，见 spec） |

### 浮动体

| 方法 | 排版 |
|------|------|
| `table(key, caption="")` | 解析 RTF → 三线表；给 caption 时自动编号「表1　…」并忽略 RTF 自带标题 |
| `figure(file, caption="", width=10)` | 插图 + 自动编号图题「图1　…」（图题在图下方，图与题不分页） |
| `image(filename)` | 只管插图，不编号（需要手写 `h3()` 做图题时才用） |
| `image("fig.png:14")` | 嵌入图片，宽 14 cm（`文件名:宽度cm` 语法） |

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
- 行内标记（`para()` 与 `ref()` 都支持，花括号内不能再嵌套花括号）：

| 写法 | 效果 |
|------|------|
| `{i:American Economic Review}` | 斜体（参考文献里的外文期刊名） |
| `{b:重点}` | 加粗 |
| `{fn:脚注内容}` | 就地插脚注：正文中间也能挂，不必另起一段 |

```python
p.para("其中，x_{it} 为产出{fn:控制变量见附录 A。}，均取对数。")
p.ref("Koopman, R., et al., 2010, “Give Credit…”, {i:NBER Working Paper}, 16426.")
```

## 图表标题约定（避免双标题）

- 用 `table(key, caption=...)` / `figure(file, caption=...)` 时编号与标题由渲染器生成，**不要再手写 `h3()`**
- 只有用底层的 `image()` 才需要自己配 `h3()` 当图题
- RTF 自带标题的表格，不传 caption 就沿用 RTF 标题；传了 caption 则以 caption 为准（避免双标题）

## 字体体系

中文排版三属性设置（`w:eastAsia` / `w:hAnsi` / `w:ascii`）。
各元素字体由 spec 决定（`body_font`、`heading1_font`、`abstract_font`、`table_font` …），
默认值见 `specs/base.yaml`；西文一律 Times New Roman。

## 默认值

以下为 `specs/base.yaml` 的中性默认值，任何期刊要求都用 `spec` 覆盖：

| 参数 | 默认值 |
|------|--------|
| 页边距 上/下/左/右 | 2.54 / 2.54 / 3.18 / 2.54 cm |
| 正文字号/行距 | 12pt / 1.5× |
| 首行缩进 | 24pt（≈两字符） |
| 一级标题 | 16pt |
| 二级标题 | 14pt |
| 图表标题 | 10.5pt |
| 参考文献 | 10pt / 1.25× |

全部字段都可在 spec 或 `render()` kwargs 里覆盖。

## 版式规格（spec）

只有一套通用版式 `specs/base.yaml`；具体刊/学校要求另写一个 yaml，`extends: base` 只列差异。

```
specs/base.yaml                    # 通用版式：中性默认值，不代表任何期刊
specs/erj.yaml                     # 例：《经济研究》投稿体例
~/.dsh/econ-cn-docx/specs/*.yaml   # 用户自定义规格（skill 升级不会丢）
```

```python
p.render("论文.docx")                            # 不指定 → specs/base.yaml
p.render("论文.docx", spec="erj")                # 按规格名
p.render("论文.docx", spec="erj", body_size=11)  # 规格 + 单点覆盖
p.render("论文.docx", spec="erj", report=True)   # 打印生效值 + 每项来源
```

- 字段清单与含义见 `specs/base.yaml`（带注释，可直接当填空模板）
- 行距写数字 = 倍数；写 `"16pt"` = 固定值
- 自动编号：`h1("引言")` → 「一、引言」，`h2("基准结果")` → 「（一）基准结果」；已有编号则跳过（`auto_number: false` 关闭）
- 参考文献标题走 `ref_heading_*`（左对齐小标题），不套一级标题
- 脚注：`footnote_per_page`（每页重编号）、`footnote_num_format`（`decimal` / `decimalEnclosedCircle` ①②③ / `chicago` `*†‡`）、`footnote_font/size`
- 注意：脚注编号格式是**整篇统一**的，做不出「首页 `*`、正文 ①②③」这种混排；`chicago` 下首页第一条就是 `*`，同页第二条是 `†`

```python
from references.style import Style, list_specs, save, load_spec
list_specs()                                             # 可用规格
Style.load(spec="erj").describe()                        # 逐项看值与来源
save("my-erj", Style.load(spec="erj", overrides={"body_size": 9}))  # 存为用户规格
```

预览（docx → pdf → PNG，用于逐轮比对）：

```python
from references.preview import preview
preview("论文.docx")        # → preview/p1.png, p2.png …
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
