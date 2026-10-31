# -*- coding: utf-8 -*-
"""脚注支持（python-docx 原生没有）。

原理
----
Word 的脚注不在 document.xml 里，而在独立的 ``word/footnotes.xml`` 部件中；
正文处只放一个引用（``w:footnoteReference w:id=N``），脚注正文按 id 放在该部件里。
python-docx 的默认模板不含这个部件，所以这里手工创建并挂关系。

对外只暴露 :func:`add_footnote` 与 :func:`set_footnote_props`。
"""

from docx.opc.packuri import PackURI
from docx.opc.part import XmlPart
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.oxml import OxmlElement

__all__ = ["add_footnote", "set_footnote_props", "footnote_count"]

FOOTNOTES_CT = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"
)
RT_FOOTNOTES = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
)

_EMPTY = (
    '<w:footnotes %s>'
    '<w:footnote w:type="separator" w:id="-1">'
    '<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
    '<w:r><w:separator/></w:r></w:p></w:footnote>'
    '<w:footnote w:type="continuationSeparator" w:id="0">'
    '<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
    '<w:r><w:continuationSeparator/></w:r></w:p></w:footnote>'
    "</w:footnotes>" % nsdecls("w")
)


def _root(doc):
    """取（必要时创建）footnotes 部件。"""
    for rel in doc.part.rels.values():
        if rel.reltype == RT_FOOTNOTES:
            return rel.target_part.element
    part = XmlPart(PackURI("/word/footnotes.xml"), FOOTNOTES_CT,
                   parse_xml(_EMPTY), doc.part.package)
    doc.part.relate_to(part, RT_FOOTNOTES)
    return part.element


def footnote_count(doc):
    root = _root(doc)
    return len([f for f in root.findall(qn("w:footnote"))
                if f.get(qn("w:type")) is None])


def add_footnote(paragraph, doc, text, font="宋体", size=7.5, first_indent_chars=2):
    """在 paragraph 末尾插脚注引用，并写入脚注正文。

    标记样式由 :func:`set_footnote_props` 决定（数字 / ①②③ / *†‡）。
    """
    root = _root(doc)
    fid = 2 + footnote_count(doc)          # -1/0 是分隔符，正文脚注从 2 起

    # 1) 脚注正文
    fn = OxmlElement("w:footnote")
    fn.set(qn("w:id"), str(fid))
    p = OxmlElement("w:p")
    ppr = OxmlElement("w:pPr")
    spacing = OxmlElement("w:spacing")
    spacing.set(qn("w:line"), "240")
    spacing.set(qn("w:lineRule"), "auto")
    spacing.set(qn("w:after"), "0")
    ppr.append(spacing)
    if first_indent_chars:
        ind = OxmlElement("w:ind")
        ind.set(qn("w:firstLineChars"), str(int(first_indent_chars * 100)))
        ind.set(qn("w:firstLine"), str(int(size * 20 * first_indent_chars)))
        ppr.append(ind)
    p.append(ppr)

    r1 = OxmlElement("w:r")
    r1.append(OxmlElement("w:footnoteRef"))
    p.append(r1)

    r2 = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    rfonts = OxmlElement("w:rFonts")
    rfonts.set(qn("w:ascii"), "Times New Roman")
    rfonts.set(qn("w:hAnsi"), "Times New Roman")
    rfonts.set(qn("w:eastAsia"), font)
    rpr.append(rfonts)
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), str(int(size * 2)))       # half-points
    rpr.append(sz)
    r2.append(rpr)
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    r2.append(t)
    p.append(r2)
    fn.append(p)
    root.append(fn)

    # 2) 正文里的引用标记
    run = paragraph.add_run()
    rpr = OxmlElement("w:rPr")
    rstyle = OxmlElement("w:rStyle")
    rstyle.set(qn("w:val"), "FootnoteReference")
    rpr.append(rstyle)
    va = OxmlElement("w:vertAlign")
    va.set(qn("w:val"), "superscript")
    rpr.append(va)
    run._r.append(rpr)
    ref = OxmlElement("w:footnoteReference")
    ref.set(qn("w:id"), str(fid))
    run._r.append(ref)
    return fid - 1


def set_footnote_props(doc, num_fmt="decimal", per_page=True):
    """设置脚注编号格式与是否每页重编。

    num_fmt: decimal（1,2,3）/ decimalEnclosedCircle（①②③）/ chicago（*†‡）
    """
    for sect in doc.sections:
        pr = sect._sectPr
        fpr = pr.find(qn("w:footnotePr"))
        if fpr is None:
            fpr = OxmlElement("w:footnotePr")
            pr.insert(0, fpr)
        for tag, val in (("w:numFmt", num_fmt),
                         ("w:numRestart", "eachPage" if per_page else "continuous")):
            el = fpr.find(qn(tag))
            if el is None:
                el = OxmlElement(tag)
                fpr.append(el)
            el.set(qn("w:val"), val)
