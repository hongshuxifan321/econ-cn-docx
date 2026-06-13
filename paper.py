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

    def equation(self, latex: str) -> None:
        """Display equation (OMML, centred). Prefix EQUATION: added."""
        self.sections.append({"title": "", "body": f"EQUATION:{latex}", "level": 0})

    def note(self, text: str) -> None:
        """Table note (9pt 宋体 italic). Prefix NOTE: added."""
        self.sections.append({"title": "", "body": f"NOTE:{text}", "level": 0})

    # ── float ─────────────────────────────────────────────────

    def table(self, key: str) -> None:
        """Insert three-line table from RTF. Prefix TABLE: added."""
        self.sections.append({"title": "", "body": f"TABLE:{key}", "level": 0})

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
           margin_*, body_size, heading1_size, heading2_size, etc.).
        """
        render_kwargs = {k: v for k, v in kwargs.items()
                         if k in _EXPECTED_RENDER_KWARGS}
        # Merge output into kwargs — render_from_dicts takes output as positional arg2
        render_from_dicts(self.sections, output, **render_kwargs)
