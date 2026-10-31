# -*- coding: utf-8 -*-
"""cn_docx — 经管实证论文中文学术 Word 排版。

   Usage:
       from cn_docx import Paper

       p = Paper()
       p.title("论文标题")
       ...
       p.render("论文.docx")          # 不写 spec = 通用版式（基于经研投稿体例）

   Or use JSON path:
       from cn_docx import render
       render("paper_content.json", output="论文.docx")

   版式管理::
       from cn_docx import Style, save, list_specs
       save("my-style", Style.load(overrides={"body_size": 9}))
"""

from references.pipeline import render, render_from_dicts, set_cn_font
from references.preview import preview
from references.style import Style, list_specs, resolve, save
from paper import Paper

__all__ = [
    "Paper", "render", "render_from_dicts", "set_cn_font", "preview",
    "Style", "resolve", "save", "list_specs",
]
