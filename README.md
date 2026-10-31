# econ-cn-docx

经管实证论文中文学术 Word 排版。Python DSL 写论文，指定一份「版式规格」，`render()` 出 docx。

```bash
pip install -e ~/.dsh/skills/econ-cn-docx
```

依赖：`python-docx`、`lxml`、`PyYAML`（跟随安装自动装）。

## 快速开始

```python
from cn_docx import Paper

p = Paper()
p.title("空气污染、远程工作可行性与城市经济产出")
p.footnote("作者信息、基金项目、致谢都放在首页脚注。")   # 页下注、每页重编号
p.subtitle("——来自中国地级市面板的证据")
p.author("张　三　李　四")
p.abstract("本文基于 2000—2019 年地级市面板数据……")
p.keywords("PM2.5　远程工作　调节效应")

p.h1("引言")                       # 自动编号成「一、引言」
p.para("空气污染损害经济产出{fn:Graff Zivin & Neidell (2012) 给出较早的证据。}，"
       "但产业结构决定了受损程度。")

p.h2("计量模型")                   # 自动编号成「（一）计量模型」
p.para("基准回归设定如下：")
p.equation(r"lnGDP_{ct} = \beta_{1} PM_{ct} + \beta_{2} PM_{ct} + \varepsilon_{ct}")

p.table("table_baseline", caption="基准回归结果")      # 自动「表1　基准回归结果」
p.figure("fig_trend.png", caption="PM2.5 与 RWI 的趋势", width=11)  # 自动「图1　…」

p.h1("参考文献")
p.ref("Dingel, J. I., and B. Neiman, 2020, “How Many Jobs Can Be Done at Home?”, "
      "{i:Journal of Public Economics}, 189, 104235.")

p.render("论文.docx", tables_dir="tables", figures_dir="figures")   # 默认即经研投稿体例
```

## 版式规格（spec）

**只有一套通用版式，就是《经济研究》投稿体例（A 版）。** 其他期刊 / 学校另写一个 yaml，`extends: base` 只列差异。

```
specs/base.yaml                    # 通用版式 = 《经济研究》投稿体例（默认出稿版式）
specs/erj.yaml                     # 同 base 的别名，为兼容旧调用保留
~/.dsh/econ-cn-docx/specs/*.yaml   # 你自己的规格（skill 升级不会丢）
```

```python
p.render("论文.docx")                            # 不指定 → 通用版式（经研体例）
p.render("论文.docx", spec="erj")                # 等价，别名
p.render("论文.docx", body_size=11)              # 通用版式 + 单点覆盖
p.render("论文.docx", report=True)               # 打印生效值 + 每项来源
```

- 字段清单、含义与**逐项出处**见 [`specs/base.yaml`](specs/base.yaml)：`sources` 标明每项是「官方样张」「刊登版实测」还是「推断」
- 行距写数字 = 倍数；写 `"16pt"` = 固定值
- 未支持的字段或拼错的参数会**直接报错**，不会静默忽略

管理规格：

```python
from references.style import Style, list_specs, save

list_specs()                                                  # 可用规格
Style.load().describe()                                       # 逐项看生效值 + 来源
save("my-erj", Style.load(overrides={"body_size": 9}))        # 存为用户规格（只写差异）
```

## API

### 元数据

| 方法 | 说明 |
|------|------|
| `title(text)` | 论文标题（居中、加粗，字号见 `title_size`） |
| `subtitle(text)` | 副标题「——…」 |
| `author(text)` | 作者姓名行 |
| `abstract(text)` | 摘要；标签文字由 `abstract_label` 决定（如「内容摘要」） |
| `keywords(text)` | 关键词行 |
| `footnote(text)` | 给**上一段**挂脚注（首页作者信息就挂在 `title` 上） |

### 标题

| 方法 | 说明 |
|------|------|
| `h1(title)` | 一级标题，自动编号「一、」（已有编号则跳过） |
| `h2(title)` | 二级标题，自动编号「（一）」 |
| `h3(title)` | 三级标题／自绘图题，不编号 |

### 正文

| 方法 | 说明 |
|------|------|
| `para(text)` | 正文段落（缩进、行距、段距见 spec） |
| `para_no_indent(text)` | 正文段落，不缩进 |
| `equation(latex)` | OMML 行间公式，居中；`equation_number: true` 时右端自动编号 (1)(2)… |
| `note(text)` | 表注行 |

### 图表

| 方法 | 说明 |
|------|------|
| `table(key, caption="")` | RTF → 三线表；给 caption 时自动编号「表1　…」 |
| `figure(file, caption="", width=10)` | 插图 + 自动编号图题「图1　…」（图题在图下方，图与题不分页） |
| `image(filename)` | 只插图不编号（需要自己 `h3()` 当图题时才用） |

### 参考文献

| 方法 | 说明 |
|------|------|
| `ref(text)` | 单条文献；`p.h1("参考文献")` 会自动识别为左对齐小标题，不参与编号 |

### 渲染

```python
p.render("论文.docx",            # 不写 spec 即通用版式（经研投稿体例）
         tables_dir="tables", figures_dir="figures",
         table_map={"key": "file.rtf"},
         report=True)             # 打印生效参数 + 来源
```

## 行内标记

`para()` 和 `ref()` 都支持，花括号内不能再嵌套花括号：

| 写法 | 效果 |
|------|------|
| `{i:American Economic Review}` | 斜体（参考文献里的外文期刊名） |
| `{b:重点}` | 加粗 |
| `{fn:脚注内容}` | 就地插脚注，正文中间也能挂 |

另外 `word_{subscript}` 会被自动识别为内联 OMML 公式（如 `PM_{2.5}`）。

## 自动编号

- 一级标题：一、二、三……；二级标题：（一）（二）……（每级自动重置）
- 已手写编号的标题不会被重复编号
- 图表：`table(key, caption=…)` / `figure(file, caption=…)` 生成「表1　…」「图1　…」
- 公式：`equation_number: true` 时右端出 (1)(2)，居中与右对齐靠制表位实现

## 脚注

- `footnote(text)` 挂到上一段（首页作者信息用这个）
- `{fn:…}` 就地插，适合正文中间
- 每页重新编号：`footnote_per_page: true`
- 标记格式：`footnote_num_format` = `decimal`（1,2,3）/ `decimalEnclosedCircle`（①②③）/ `chicago`（`*`、`†`、`‡`）
- 注意：编号格式是**整篇统一**的，做不出「首页 `*`、正文 ①②③」混排；`chicago` 下首页第一条就是 `*`

## 预览

逐轮比对排版结果，不用开 Word：

```python
from references.preview import preview
preview("论文.docx")        # → preview/p1.png, p2.png …
```

需要 LibreOffice（Windows 上自动用 `soffice.com`；配置目录写在 docx 同级的 `.lo_profile`）。

## OMML 公式引擎

| 语法 | 效果 |
|------|------|
| `x_{sub}` / `x^{sup}` / `x_{sub}^{sup}` | 上下标 |
| `\frac{a}{b}` / `\sqrt{x}` / `\sqrt[n]{x}` | 分式、根号 |
| `\sum_{i=1}^{n}` `\prod` `\int` | 大型运算符，上下限在符号正上/正下方 |
| `\bar{x}` `\hat{x}` `\tilde{x}` `\dot{x}` | 组合变音符 |
| `\alpha` … `\Omega` | 希腊字母 |
| `\times` `\cdot` `\pm` `\leq` `\geq` `\neq` | 运算符 |
| `\partial` `\infty` `\nabla` `\forall` `\exists` | 数学符号 |
| `\rightarrow` `\to` `\Rightarrow` | 箭头 |
| `\text{GDP}` | 公式内文本 |

已知边界：`\text{}` 与上下标内容不递归解析；`\sum_{…}` 参数内不支持嵌套花括号；连续单字符下标 `a_b_c` 请写成 `a_{b}_{c}`。

## 兼容：JSON 入口

```python
from cn_docx import render
render("paper_content.json", output="论文.docx")   # 默认即通用版式
```

`TABLE:key` / `IMAGE:file.png` / `EQUATION:...` / `REF:...` / `NOTE:...` 前缀仍然有效。

## 字体与默认值

各元素字体由 spec 决定（`body_font`、`heading1_font`、`abstract_font`、`table_font` …），
西文一律 Times New Roman。默认值 = 《经济研究》投稿体例，见 [`specs/base.yaml`](specs/base.yaml)。

## 局限

- 无页眉／目录／交叉引用（图表编号是顺序号，不是域）
- 脚注标记格式整篇统一，不能按脚注单独指定
- RTF 仅解析 esttab/estout 输出
- 弯引号修复为启发式，极少数嵌套引号可能误判
- 预览依赖 LibreOffice

## 许可证

MIT
