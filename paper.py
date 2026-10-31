# -*- coding: utf-8 -*-
"""Paper — Python DSL for econ-cn-docx.

   Usage:
       from cn_docx import Paper

       p = Paper()
       p.title("论文标题")
       p.abstract("摘要内容...")
       p.keywords("关键词1；关键词2")
       p.h1("一、引言")
       p.para("正文...")
       p.h2("（一）基准回归")
       p.equation(r"y_{it} = \\beta x_{it} + \\alpha_i + \\varepsilon_{it}")
       p.table("table_baseline_reg")
       p.image("fig_trend.png")
       p.h3("图一：标题")
       p.ref("[1] Author. Title[J]. Journal, Year.")
       p.render("论文.docx")
"""

from references.pipeline import render_from_dicts


_EXPECTED_RENDER_KWARGS = {
    "output", "tables_dir", "figures_dir", "table_map",
    "margin_top", "margin_bottom", "margin_left", "margin_right",
    "body_size", "body_line_spacing", "body_indent",
    "heading1_size", "heading2_size", "heading3_size",
    "ref_size", "ref_line_spacing", "table_font_size",
    "spec", "style", "report",
}


class Paper:
    """Accumulate sections and render to .docx via econ-cn-docx pipeline."""

    def __init__(self):
        self.sections = []

    # ── metadata ──────────────────────────────────────────────

    def title(self, text: str) -> None:
        """Main paper title (level 0, 18pt bold centred)."""
        self.sections.append({"title": text, "body": "", "level": 0})

    def abstract(self, text: str) -> None:
        """Chinese abstract (level 1, 楷体)。"""
        self.sections.append({"title": "摘要", "body": text, "level": 1})

    def keywords(self, text: str) -> None:
        """Keywords line (level 1, 楷体)。"""
        self.sections.append({"title": "关键词", "body": text, "level": 1})

    def subtitle(self, text: str) -> None:
        """副标题「——…」，居中，字体见 subtitle_font/size。"""
        self.sections.append({"title": "", "body": f"SUBTITLE:{text}", "level": 0})

    def author(self, text: str) -> None:
        """作者姓名行，居中，字体见 author_font/size。"""
        self.sections.append({"title": "", "body": f"AUTHOR:{text}", "level": 0})

    # ── headings ──────────────────────────────────────────────

    def h1(self, title: str, body: str = "") -> None:
        """Level-1 heading (一、引言 etc.)。"""
        self.sections.append({"title": title, "body": body, "level": 1})

    def h2(self, title: str, body: str = "") -> None:
        """Level-2 heading (（一）基准回归 etc.)。"""
        self.sections.append({"title": title, "body": body, "level": 2})

    def h3(self, title: str) -> None:
        """Level-3 heading (图/表 caption, 10.5pt bold centred)。"""
        self.sections.append({"title": title, "body": "", "level": 3})

    # ── body ──────────────────────────────────────────────────

    def para(self, text: str) -> None:
        """Body paragraph (12pt 宋体 justified, first-line indent)。"""
        self.sections.append({"title": "", "body": text, "level": 0})

    def para_no_indent(self, text: str) -> None:
        """Body paragraph without first-line indent (顶格)。"""
        self.sections.append({"title": "", "body": f"NOINDENT:{text}", "level": 0})

    def equation(self, latex: str) -> None:
        """Display equation (OMML, centred). Prefix EQUATION: added."""
        self.sections.append({"title": "", "body": f"EQUATION:{latex}", "level": 0})

    def note(self, text: str) -> None:
        """Table note (9pt 宋体 italic). Prefix NOTE: added."""
        self.sections.append({"title": "", "body": f"NOTE:{text}", "level": 0})

    # ── float ─────────────────────────────────────────────────

    def table(self, key: str, caption: str = "") -> None:
        """三线表。caption 非空时自动编号为「表N　caption」。"""
        self.sections.append({"kind": "table", "key": key, "caption": caption})

    def figure(self, filename: str, caption: str = "", width: float = 10) -> None:
        """插图 + 自动编号图题「图N　caption」（图题在图下方）。"""
        self.sections.append({"kind": "figure", "file": filename,
                              "caption": caption, "width": width})

    def footnote(self, text: str) -> None:
        """给上一段（通常是标题）挂脚注，页下注、每页重新编号。

        标记样式由 spec 的 footnote_num_format 决定（chicago → *、†、‡）。
        """
        self.sections.append({"kind": "footnote", "text": text})

    def image(self, filename: str) -> None:
        """Insert image (centred, 10 cm wide). Prefix IMAGE: added."""
        self.sections.append({"title": "", "body": f"IMAGE:{filename}", "level": 0})

    # ── back matter ───────────────────────────────────────────

    def ref(self, text: str) -> None:
        """Reference entry (10pt 宋体, 1.25× line spacing). Prefix REF: added."""
        self.sections.append({"title": "", "body": f"REF:{text}", "level": 1})

    # ── render ────────────────────────────────────────────────

    def render(self, output: str = "论文.docx", **kwargs) -> None:
        """Render accumulated sections to .docx.

           Accepts all render() kwargs (tables_dir, figures_dir, table_map,
           margin_*, body_size, heading1_size, heading2_size, spec, etc.).

           spec:   规格名或 yaml 路径（"base"、"/path/to/x.yaml"），见 specs/；
           style:  直接传入 Style 对象（优先级最高）；
           report: True 时打印生效参数与来源。
        """
        unknown = set(kwargs) - _EXPECTED_RENDER_KWARGS
        if unknown:
            print(f"[cn_docx] WARNING: 未知参数被忽略: {sorted(unknown)}")
        render_kwargs = {k: v for k, v in kwargs.items()
                         if k in _EXPECTED_RENDER_KWARGS}
        # Merge output into kwargs — render_from_dicts takes output as positional arg2
        render_from_dicts(self.sections, output, **render_kwargs)
