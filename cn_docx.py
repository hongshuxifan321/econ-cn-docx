# -*- coding: utf-8 -*-
"""cn_docx — 经管实证论文中文学术 Word 排版。

   Usage:
       from cn_docx import Paper

       p = Paper()
       p.title("论文标题")
       ...
       p.render("论文.docx")

   Or use JSON path:
       from cn_docx import render
       render("paper_content.json", output="论文.docx")
"""

from references.pipeline import render, render_from_dicts, set_cn_font
from paper import Paper
