---
name: econ-cn-docx
description: 经管实证论文中文学术 Word 排版。Python DSL（Paper 类）+ 版式规格（specs/*.yaml）生成 .docx，含 OMML 公式、标题/图表/公式自动编号、页下注、三线表、参考文献体例、docx→PNG 预览。
argument-hint: "<task> e.g. '生成 docx' / '按经济研究体例排版' / '给这段加脚注'"
user-invocable: true
---

# econ-cn-docx

论文排版器：`Paper` 累积内容 → 套**通用版式**（= 《经济研究》投稿体例）→ `render()` 出 docx。

- 完整用法与 API：[`README.md`](README.md)
- 通用版式（字段全集 + 逐项出处，可当填空模板）：[`specs/base.yaml`](specs/base.yaml)
- `specs/erj.yaml` 是 base 的别名，`spec="erj"` 与不写 spec 等价

## 命令

```bash
pip install -e ~/.dsh/skills/econ-cn-docx
python demo.py            # 通用版式（经研投稿体例）
python demo.py erj        # 等价
```

## 最小示例

```python
from cn_docx import Paper

p = Paper()
p.title("……")
p.footnote("作者信息、基金项目、致谢。")      # 挂在标题上的页下注
p.subtitle("——来自……的证据")
p.author("张　三　李　四")
p.abstract("……")
p.keywords("……　……")

p.h1("引言")                      # 自动编号「一、引言」
p.para("正文{fn:这句的脚注。}继续。")          # 行内脚注
p.h2("计量模型")                  # 自动编号「（一）计量模型」
p.equation(r"y_{it} = \beta x_{it} + \varepsilon_{it}")

p.table("table_baseline", caption="基准回归结果")     # 「表1　基准回归结果」
p.figure("fig_trend.png", caption="趋势", width=11)   # 「图1　趋势」

p.h1("参考文献")                  # 识别为小标题，不编号
p.ref("Dingel, J. I., and B. Neiman, 2020, “…”, "
      "{i:Journal of Public Economics}, 189, 104235.")

p.render("论文.docx", spec="erj")
```

## 方法速查

| 方法 | 一句话 |
|------|--------|
| `title(text)` | 论文标题 |
| `subtitle(text)` | 副标题「——…」 |
| `author(text)` | 作者姓名行 |
| `abstract(text)` | 摘要；标签字由 `abstract_label` 决定（如「内容摘要」） |
| `keywords(text)` | 关键词行 |
| `footnote(text)` | 给**上一段**挂页下注（首页作者信息就挂在 `title` 上） |
| `h1(title)` / `h2(title)` | 一二三级标题；h1/h2 自动编号，已写编号则跳过 |
| `h3(title)` | 三级标题／自绘图题，不编号 |
| `para(text)` | 正文段落（缩进/行距/段距见 spec） |
| `para_no_indent(text)` | 正文段落，不缩进 |
| `equation(latex)` | OMML 行间公式；`equation_number` 开启时右端自动 (n) |
| `note(text)` | 表注行 |
| `table(key, caption="")` | RTF → 三线表；给 caption 自动编号「表N」 |
| `figure(file, caption="", width=10)` | 插图 + 自动编号图题（图与题不分页） |
| `image(filename)` | 只插图不编号（自己写 `h3()` 当图题时才用） |
| `ref(text)` | 单条参考文献 |
| `render(out, spec=…)` | 渲染；`report=True` 打印生效值与来源 |

行内标记（`para()` 与 `ref()` 通用）：`{i:斜体}`、`{b:加粗}`、`{fn:脚注}`；`PM_{2.5}` 自动转内联公式。

## 选择规则

- 表格：要编号就传 `caption`；不传则沿用 RTF 自带标题
- 图：优先 `figure()`；只有需要自定义图题样式时才用 `image()` + `h3()`
- 脚注：首页信息用 `footnote()`，正文中间用 `{fn:…}`
- 参考文献：`h1("参考文献")` 会自动走小标题样式

## 易错点（都踩过）

1. **拼错参数会直接报错**（KeyError），不静默忽略——字段名以 `specs/base.yaml` 为准
2. 摘要/关键词/参考文献**不参与**自动编号；标题里已写「一、」不会再加
3. 表格传了 `caption` 就忽略 RTF 自带标题，避免双标题
4. 脚注编号格式**整篇统一**（`decimal` / `decimalEnclosedCircle` ①②③ / `chicago` `*†‡`），做不出「首页 `*` + 正文 ①」混排
5. 插图段落必须单倍行距，否则被固定行距裁成一条——渲染器已处理，别在外层覆盖段落行距
6. 预览要 LibreOffice：Windows 用 `soffice.com`（`.exe` 拿不到退出码），配置文件写在 docx 同级的 `.lo_profile`（放 %TEMP% 会被沙箱拒绝）

## 扩展版式

```python
# 1) 新期刊：新建 yaml，首行写 extends: base，只列与经研体例不同的键
# 2) 用户自定义规格放 ~/.dsh/econ-cn-docx/specs/（skill 升级不会丢）
from references.style import Style, list_specs, save
list_specs()
Style.load().describe()                              # 逐项看生效值 + 来源
save("my-erj", Style.load(overrides={"body_size": 9}))
```

`render(..., report=True)` 会打印生效参数和每项来源（默认/规格/调用参数）。

## 边界

- 无页眉 / 目录 / 交叉引用；图表编号是顺序号，不是 Word 域
- 只解析 esttab/estout 的 RTF 表格
- 弯引号修复是启发式，极端嵌套可能误判
