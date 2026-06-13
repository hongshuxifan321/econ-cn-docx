# econ-cn-docx

经管实证论文中文学术 Word 排版。Python DSL 写论文，一个 `render()` 出 docx。

```bash
pip install -e ~/.claude/skills/econ-cn-docx
```

## 快速开始

```python
from cn_docx import Paper

p = Paper()
p.title("论文标题")
p.abstract("摘要内容……")
p.keywords("关键词1；关键词2")

p.h1("一、引言")
p.para("正文段落。支持内联公式如 PM_{2.5} 自动渲染为 OMML。\n\n第二段。")

p.h2("（一）基准回归")
p.equation(r"lnGDP_{ct} = \beta_{1} PM_{ct} + \beta_{2} PM_{ct} \times RWI_{c} + \alpha_{c} + \lambda_{t} + \varepsilon_{ct}")

p.para("回归结果如下。")
p.table("table_baseline_reg")

p.image("fig_trend.png")
p.h3("图一：PM2.5 趋势")

p.h1("参考文献")
p.ref("[1] Dingel, J. I., Neiman, B. How many jobs can be done at home?[J]. Journal of Public Economics, 2020, 189: 104235.")

p.render("论文.docx", tables_dir="tables", figures_dir="figures")
```

## API

### 元数据
| 方法 | 排版 |
|------|------|
| `title(text)` | 18pt 黑体居中加粗 |
| `abstract(text)` | "摘要："黑体 10.5pt 加粗 + 楷体正文 |
| `keywords(text)` | "关键词："黑体 10.5pt 加粗 + 楷体正文 |

### 标题
| 方法 | 排版 |
|------|------|
| `h1(title, body="")` | 16pt 黑体居中加粗 |
| `h2(title, body="")` | 14pt 黑体左对齐加粗 |
| `h3(title)` | 10.5pt 黑体居中加粗（图/表标题） |

### 正文
| 方法 | 排版 |
|------|------|
| `para(text)` | 12pt 宋体两端对齐，首行缩进 24pt，1.5 倍行距 |
| `equation(latex)` | OMML 行间公式，居中 |
| `note(text)` | 9pt 宋体斜体（表注） |

### 浮动体
| 方法 | 排版 |
|------|------|
| `table(key)` | RTF → 三线表（顶线 12/底线 12/表头下线 8） |
| `image(filename)` | 嵌入 PNG，居中，宽 10 cm |

### 参考文献
| 方法 | 排版 |
|------|------|
| `ref(text)` | 10pt 宋体，1.25 倍行距 |

### 渲染
```python
p.render("论文.docx",
         tables_dir="tables", figures_dir="figures",
         table_map={"key": "file.rtf"},  # TABLE key → RTF 文件名
         margin_left=3.18, body_size=12, heading1_size=16,
         # 全部参数可覆盖，默认值为毕业论文格式
)
```

## OMML 公式引擎

| 语法 | 效果 |
|------|------|
| `x_{sub}` / `x^{sup}` | 下标 / 上标 |
| `x_{sub}^{sup}` | 同时上下标 |
| `\frac{a}{b}` | 分式 |
| `\sqrt{x}` / `\sqrt[n]{x}` | 平方根 / n 次方根 |
| `\bar{x}` `\hat{x}` `\tilde{x}` `\dot{x}` | 组合变音符 |
| `\alpha` `\beta` `\gamma` `\delta` … | 小写希腊字母 |
| `\Gamma` `\Delta` `\Sigma` `\Phi` … | 大写希腊字母 |
| `\times` `\cdot` `\pm` `\leq` `\geq` `\neq` | 运算符 |
| `\partial` `\infty` `\nabla` `\forall` `\exists` | 数学符号 |
| `\rightarrow` / `\to` / `\Rightarrow` | 箭头 |
| `\text{GDP}` | 公式内文本 |

## 行内格式

- `H1`/`H2`/`H3` 开头段落 → 加粗（假说陈述行）
- `word_{subscript}` 模式自动检测 → 内联 OMML

## 字体体系

| 元素 | 中文字体 | 西文字体 |
|------|---------|---------|
| 正文 | 宋体 | Times New Roman |
| 标题 | 黑体 | Times New Roman |
| 摘要/关键词 | 楷体 | Times New Roman |
| 表格 | 宋体 | Times New Roman |

## 排版参数默认值

| 参数 | 默认值 |
|------|--------|
| 页边距（上/下/左/右） | 2.54 / 2.54 / 3.18 / 2.54 cm |
| 正文字号 / 行距 | 12pt / 1.5× |
| 首行缩进 | 24pt（≈两字符） |
| 一级标题 | 16pt |
| 二级标题 | 14pt |
| 图表标题 | 10.5pt |
| 参考文献 | 10pt / 1.25× |

所有参数通过 `render()` kwargs 按需覆盖。期刊投稿示例：
```python
p.render("论文.docx", margin_left=2.54, body_size=10.5, body_indent=21,
         heading1_size=14, heading2_size=12, ref_size=9)
```

## 旧版兼容：JSON 入口

```python
from cn_docx import render
render("paper_content.json", output="论文.docx")
```

JSON 格式仍支持所有 `TABLE:key` / `IMAGE:file.png` / `EQUATION:...` / `REF:...` / `NOTE:...` 前缀。

## 依赖

- python-docx >= 1.0
- lxml >= 4.9

## 局限

- 无页码 / 页眉 / 目录（需在 Word 中手动添加）
- 无图表交叉引用、公式编号
- RTF 仅解析 esttab/estout 输出
- 弯引号修复为启发式，极少数嵌套引号可能误判

## 许可证

MIT
