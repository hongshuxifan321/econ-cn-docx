# -*- coding: utf-8 -*-
"""cnpaper — 经管实证论文中文学术 Word 排版（单文件模块）。

用法::

    from cnpaper import Paper

    p = Paper()
    p.title("标题")
    p.author("张　三")
    p.abstract("……")
    p.render("论文.docx")          # 不写 spec = 通用版式（基于经研投稿体例）

版式：通用版式写在 specs/base.yaml；用户自定义规格放 ~/.dsh/econ-cn-docx/specs/
"""

__all__ = [
    "Paper", "render", "render_from_dicts", "set_cn_font", "preview",
    "Style", "DEFAULTS", "resolve", "load_spec", "save", "list_specs",
    "USER_SPEC_DIR", "SKILL_SPEC_DIR",
]

import json
import os
import re
import subprocess

import yaml
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.opc.packuri import PackURI
from docx.opc.part import XmlPart
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Pt
from lxml import etree


# ==============================================================================
# 来自 style.py
# ==============================================================================
import os

import yaml


USER_SPEC_DIR = os.path.expanduser("~/.dsh/econ-cn-docx/specs")
SKILL_SPEC_DIR = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "specs"
)

# ═══════════════════════════════════════════════════════════════
# 参数全集（单一事实来源：pipeline 从这里取默认值）
# 值 = 《经济研究》投稿体例（A 版），与 specs/base.yaml 一致，改一边必须改另一边
# ═══════════════════════════════════════════════════════════════

DEFAULTS = dict(
    # 目录与表格映射
    tables_dir="tables", figures_dir="figures", table_map=None,
    # 页面
    margin_top=2.54, margin_bottom=2.54, margin_left=2.54, margin_right=2.54,
    # 标题块
    title_size=15, title_font="宋体", title_bold=True,
    subtitle_size=14, subtitle_font="宋体", subtitle_bold=False,
    author_size=12, author_font="宋体", author_bold=False,
    # 摘要 / 关键词
    abstract_label="内容摘要", abstract_label_font="黑体", abstract_label_size=10.5,
    abstract_font="仿宋", abstract_size=10.5, abstract_indent=21,
    keyword_label="关键词", keyword_label_font="黑体", keyword_label_size=10.5,
    keyword_font="仿宋", keyword_size=10.5,
    # 正文
    body_size=10.5, body_font="宋体", body_line_spacing="16pt", body_indent=21,
    body_space_before=0, body_space_after=0,
    # 各级标题
    heading1_size=14, heading1_font="宋体", heading1_bold=False, heading1_align="center",
    heading2_size=10.5, heading2_font="宋体", heading2_bold=False, heading2_align="left",
    heading3_size=9, heading3_font="黑体", heading3_bold=False, heading3_align="center",
    # 表格
    table_font_size=9, table_font="仿宋", table_header_font="仿宋", table_header_bold=False,
    table_rule_top=12, table_rule_mid=8, table_rule_bottom=12,
    table_title_size=10.5, table_title_font="仿宋", table_title_bold=False,
    table_note_size=7.5, table_note_font="宋体",
    # 表注 / 脚注样式（NOTE: 行）
    note_size=7.5, note_font="宋体", note_italic=False,
    # 参考文献
    ref_size=7.5, ref_font="宋体", ref_line_spacing=1.0, ref_indent=21,
    # 页码
    page_number=True, page_number_size=9, page_number_font="宋体",
    # 自动编号（一、/（一）/图1/表1/公式(1)）
    auto_number=True,
    equation_number=True,
    # 参考文献标题
    ref_heading_font="黑体", ref_heading_size=10.5, ref_heading_bold=True,
    ref_heading_align="left",
    # 脚注
    footnote_font="宋体", footnote_size=7.5, footnote_per_page=True,
    footnote_num_format="chicago",   # decimal / decimalEnclosedCircle / chicago
)


# ═══════════════════════════════════════════════════════════════
# 规格文件加载
# ═══════════════════════════════════════════════════════════════

def _find_spec_file(name):
    """名字 → 路径。顺序：显式路径 → 用户目录 → skill 目录。"""
    if name.endswith((".yaml", ".yml")):
        if os.path.exists(name):
            return name
        if os.path.exists(os.path.join(USER_SPEC_DIR, name)):
            return os.path.join(USER_SPEC_DIR, name)
        if os.path.exists(os.path.join(SKILL_SPEC_DIR, name)):
            return os.path.join(SKILL_SPEC_DIR, name)
        raise FileNotFoundError("找不到规格文件: %s" % name)
    for d in (USER_SPEC_DIR, SKILL_SPEC_DIR):
        p = os.path.join(d, name + ".yaml")
        if os.path.exists(p):
            return p
    avail = list_specs()
    raise FileNotFoundError(
        "找不到规格 '%s'；现有可用规格: %s（默认版式 base 不需要指定，"
        "自定义规格放 %s）"
        % (name, ", ".join(avail) if avail else "无，只有通用版式 base",
           USER_SPEC_DIR)
    )


def list_specs():
    """列出可用规格名（用户目录在前）。"""
    names = []
    for d in (USER_SPEC_DIR, SKILL_SPEC_DIR):
        if not os.path.isdir(d):
            continue
        for f in sorted(os.listdir(d)):
            if f.endswith((".yaml", ".yml")) and f != "base.yaml":
                n = f.rsplit(".", 1)[0]
                if n not in names:
                    names.append(n)
    return names


def load_spec(name, _seen=None):
    """加载规格（支持 extends 继承），返回 dict（不含 extends/sources 键）。"""
    _seen = _seen or []
    path = _find_spec_file(name)
    if path in _seen:
        raise ValueError("规格循环继承: %s" % path)
    with open(path, "r", encoding="utf-8") as f:
        data = yaml.safe_load(f) or {}
    if not isinstance(data, dict):
        raise ValueError("规格文件必须是键值表: %s" % path)
    sources = data.pop("sources", {}) or {}
    data.pop("note", None)
    parent = data.pop("extends", None)
    if parent:
        base = load_spec(parent, _seen + [path])
        base_sources = base.pop("_sources", {})
        merged = dict(base)
        merged.update(data)
        merged_sources = dict(base_sources)
        merged_sources.update(sources)
        merged["_sources"] = merged_sources
        return merged
    data["_sources"] = sources
    return data


def _spec_payload(name):
    """(params, sources, label) — 供 resolve 使用。"""
    raw = load_spec(name)
    sources = raw.pop("_sources", {}) or {}
    label = os.path.splitext(os.path.basename(name))[0]
    return raw, sources, label


# ═══════════════════════════════════════════════════════════════
# Style
# ═══════════════════════════════════════════════════════════════

class Style:
    """生效的排版参数 + 每项来源。"""

    def __init__(self, params, sources=None):
        self.params = dict(params)
        self.sources = dict(sources or {})

    # ---- 构造 ----
    @classmethod
    def load(cls, spec=None, overrides=None):
        return resolve(spec=spec, overrides=overrides)

    # ---- 访问 ----
    def get(self, key, default=None):
        return self.params.get(key, DEFAULTS.get(key, default))

    def __getitem__(self, key):
        return self.params[key]

    def to_params(self):
        return dict(self.params)

    # ---- 报告 ----
    def describe(self, keys=None, file=None):
        """打印生效值 + 来源（默认只打印排版相关项）。"""
        keys = keys or [k for k in DEFAULTS if not k.endswith("_dir")
                        and k not in ("table_map",)]
        lines = ["%-24s %-22s %s" % ("参数", "生效值", "来源"),
                 "-" * 64]
        for k in keys:
            lines.append("%-24s %-22s %s" % (k, str(self.get(k)), self.sources.get(k, "")))
        text = "\n".join(lines)
        if file:
            with open(file, "w", encoding="utf-8") as f:
                f.write(text + "\n")
        return text


def resolve(spec=None, overrides=None):
    """合并顺序：DEFAULTS → spec → 调用参数；同时记录来源。

    不指定 spec 就用 base.yaml（= 通用版式 = 《经济研究》投稿体例）。
    """
    params = dict(DEFAULTS)
    sources = {k: "默认" for k in DEFAULTS}
    spec = spec or "base"

    if spec:
        data, spec_sources, label = _spec_payload(spec)
        unknown = [k for k in data if k not in DEFAULTS]
        if unknown:
            raise KeyError(
                "规格 %s 含未知字段 %s（不是全部拼错，也可能是渲染器还不支持的能力）"
                % (label, sorted(unknown))
            )
        for k, v in data.items():
            params[k] = v
            sources[k] = spec_sources.get(k, "规格:%s" % label)

    unknown = [k for k in (overrides or {}) if k not in DEFAULTS]
    if unknown:
        raise KeyError(
            "未知参数 %s（拼错了，或者渲染器还不支持这个能力）" % sorted(unknown))
    for k, v in (overrides or {}).items():
        params[k] = v
        sources[k] = "调用参数"

    # 兼容旧名：body_line_spacing 可以是倍数(1.5)或固定值字符串("16pt")
    return Style(params, sources)


def save(name, style, extends="base", note=""):
    """把 Style 存进用户目录，与 base 并行存在。"""
    os.makedirs(USER_SPEC_DIR, exist_ok=True)
    if not name.endswith((".yaml", ".yml")):
        name += ".yaml"
    base = Style.load(spec=extends) if extends else Style({})
    diffs = {}
    for k, v in style.params.items():
        if base.params.get(k) != v:
            diffs[k] = v
    srcs = {k: style.sources.get(k, "") for k in diffs}
    payload = {"extends": extends}
    if note:
        payload["note"] = note
    payload.update(diffs)
    payload["sources"] = srcs
    path = os.path.join(USER_SPEC_DIR, name)
    with open(path, "w", encoding="utf-8") as f:
        yaml.safe_dump(payload, f, allow_unicode=True, sort_keys=False)
    spec_name = os.path.splitext(os.path.basename(path))[0]
    print("[cnpaper] 已存为用户规格: %s（与 %s 的差异 %d 项）\n"
          "          下次调用: p.render(..., spec=\"%s\")"
          % (path, extends or "（空）", len(diffs), spec_name))
    return path


# ==============================================================================
# 来自 footnotes.py
# ==============================================================================
from docx.opc.packuri import PackURI
from docx.opc.part import XmlPart
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.oxml import OxmlElement


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


# ==============================================================================
# 来自 pipeline.py
# ==============================================================================
import json, os, re
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from lxml import etree

def _cn_num(n):
    """1 → 一、11 → 十一（到九十九够用）。"""
    d = "零一二三四五六七八九"
    if n <= 10:
        return "十" if n == 10 else d[n]
    if n < 20:
        return "十" + d[n % 10]
    if n < 100:
        return d[n // 10] + "十" + (d[n % 10] if n % 10 else "")
    return str(n)


def _spacing(v):
    """行距：数字 = 倍数；"16pt" / "16磅" / "固定16" = 固定值。"""
    if isinstance(v, str):
        m = re.match(r"^\s*(?:固定|exact)?\s*([\d.]+)\s*(?:pt|磅)?\s*$", v)
        if m:
            return Pt(float(m.group(1)))
    return v


def _align(name):
    return {
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
        "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
    }.get(name, WD_ALIGN_PARAGRAPH.LEFT)


def _add_page_number(doc, font="宋体", size=9):
    """页脚居中页码（PAGE 域）。"""
    sec = doc.sections[0]
    p = sec.footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    set_cn_font(run, font, Pt(size))
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    fld_sep = OxmlElement("w:fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")
    txt = OxmlElement("w:t")
    txt.text = "1"
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    for el in (fld_begin, instr, fld_sep, txt, fld_end):
        run._r.append(el)

# ═══════════════════════════════════════════════════════════════
# Chinese quote repair
# ═══════════════════════════════════════════════════════════════

def _fix_quotes(text):
    """ASCII straight quotes -> Chinese curly quotes (paired).

    Only converts when the text contains CJK characters; pure ASCII/English
    text is left untouched, avoiding false positives like
    'He said "hello"' where a space precedes the quote.
    """
    if not text:
        return text
    # Quick pre-scan: if no CJK, leave ASCII quotes alone
    if not any(ord(ch) > 127 for ch in text):
        return text
    result = []
    in_quote = False
    for i, ch in enumerate(text):
        if ch == '"':
            prev = text[i-1] if i > 0 else ' '
            nxt = text[i+1] if i+1 < len(text) else ' '
            if not in_quote:
                # 开引号:前面不是 ASCII 字母数字(空格、CJK、标点、行首均可开)
                if i == 0 or not (prev.isalnum() and ord(prev) < 128):
                    result.append('“')
                    in_quote = True
                    continue
            else:
                # 闭引号:后面不是 ASCII 字母数字(空格、CJK、标点、行尾均可闭)
                if i+1 >= len(text) or not (nxt.isalnum() and ord(nxt) < 128):
                    result.append('”')
                    in_quote = False
                    continue
        result.append(ch)
    if in_quote:
        result.append('”')
    return _remove_extra_spaces(''.join(result))


# ═══════════════════════════════════════════════════════════════
# CJK spacing cleanup
# ═══════════════════════════════════════════════════════════════

# 汉字 + 全角标点(弯引号/书名号/顿号句号/全角符号等)
_CJK_FULLWIDTH = (
    r'一-鿿'                 # 汉字
    r'‘’“”'      # 弯单/双引号
    r'、。《》「」『』'  # 顿句书名号
    r'！-～'                 # 全角标点
    r'—…'                  # 破折号/省略号
)
# 西文内容字符: 字母数字 + % + 半角括号
# (% 是数字后缀, 与汉字间不留空格「1%的」; 括号同理「他(说)」)
_LATIN_CONTENT = r'A-Za-z0-9%()'
# 西文标点: 其后空格一律保留(参考文献 GB/T 7714 的 `, ` `. ` 及序号 `[2] ` 格式空格)
_LATIN_PUNCT = r',.;:!?\[\]{}<>"\x27/\\@#$&*+=|~^`'


def _remove_extra_spaces(text):
    """删除中文语境中的多余半角空格, 输出全紧凑混排。

    删: 汉字/全角标点 与 汉字/全角标点 之间、汉字/全角标点 与 西文字母数字 之间。
    留: 西文标点之后的空格(参考文献 `陈诗一, 陈登科. 雾霾污染` 的格式空格),
        纯英文内部空格(`Dingel, J. I.`)。
    例: "GDP" 与 "CPI" → “GDP”与“CPI”;  GDP 与 CPI → GDP与CPI;  PM2.5 浓度 → PM2.5浓度
    """
    text = re.sub(rf'(?<=[{_CJK_FULLWIDTH}]) (?=[{_CJK_FULLWIDTH}])', '', text)
    text = re.sub(rf'(?<=[{_LATIN_CONTENT}]) (?=[{_CJK_FULLWIDTH}])', '', text)
    text = re.sub(rf'(?<=[{_CJK_FULLWIDTH}]) (?=[{_LATIN_CONTENT}])', '', text)
    return text


# ═══════════════════════════════════════════════════════════════
# Chinese font helper
# ═══════════════════════════════════════════════════════════════

def set_cn_font(run, font_name="宋体", size=Pt(12), bold=False):
    """Set w:ascii=TNR, w:hAnsi/w:eastAsia=font_name so Chinese
       punctuation (U+201C/D etc.) renders in the correct font."""
    run.font.name = 'Times New Roman'
    rPr = run._element.rPr
    if rPr is None:
        rPr = OxmlElement('w:rPr')
        run._element.insert(0, rPr)
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:eastAsia'), font_name)
    rFonts.set(qn('w:hAnsi'), font_name)
    run.font.size = size
    run.font.bold = bold


# ═══════════════════════════════════════════════════════════════
# OMML equation engine (LaTeX → Word native math)
# ═══════════════════════════════════════════════════════════════

MATH_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
XML_NS  = 'http://www.w3.org/XML/1998/namespace'

# LaTeX command → Unicode (rendered directly in OMML)
_LATEX_UNICODE = {
    '\\alpha': 'α', '\\beta': 'β', '\\gamma': 'γ', '\\delta': 'δ',
    '\\varepsilon': 'ε', '\\zeta': 'ζ', '\\eta': 'η', '\\theta': 'θ',
    '\\iota': 'ι', '\\kappa': 'κ', '\\lambda': 'λ', '\\mu': 'μ',
    '\\nu': 'ν', '\\xi': 'ξ', '\\pi': 'π', '\\rho': 'ρ',
    '\\sigma': 'σ', '\\tau': 'τ', '\\phi': 'φ', '\\chi': 'χ',
    '\\psi': 'ψ', '\\omega': 'ω',
    '\\Gamma': 'Γ', '\\Delta': 'Δ', '\\Theta': 'Θ', '\\Lambda': 'Λ',
    '\\Pi': 'Π', '\\Sigma': 'Σ', '\\Phi': 'Φ', '\\Psi': 'Ψ', '\\Omega': 'Ω',
    '\\cdot': '·', '\\times': '×', '\\div': '÷', '\\pm': '±',
    '\\leq': '≤', '\\geq': '≥', '\\neq': '≠', '\\approx': '≈',
    '\\to': '→', '\\rightarrow': '→', '\\Rightarrow': '⇒',
    '\\mid': '∣', '\\sum': '∑', '\\prod': '∏', '\\int': '∫',
    '\\partial': '∂', '\\infty': '∞', '\\nabla': '∇',
    '\\forall': '∀', '\\exists': '∃', '\\in': '∈', '\\notin': '∉',
    '\\subset': '⊂', '\\subseteq': '⊆', '\\sim': '∼', '\\equiv': '≡',
    '\\ln': 'ln', '\\log': 'log', '\\exp': 'exp',
    '\\max': 'max', '\\min': 'min',
}
# Combining diacritics
_DIACRITICS = {
    '\\bar':  '\N{COMBINING MACRON}',
    '\\hat':  '\N{COMBINING CIRCUMFLEX ACCENT}',
    '\\tilde': '\N{COMBINING TILDE}',
    '\\dot':  '\N{COMBINING DOT ABOVE}',
}


def _m(tag):
    return etree.Element(f'{{{MATH_NS}}}{tag}')


def _m_run(parent, text):
    r = etree.SubElement(parent, f'{{{MATH_NS}}}r')
    t = etree.SubElement(r, f'{{{MATH_NS}}}t')
    t.set(f'{{{XML_NS}}}space', 'preserve')
    t.text = text
    return r


def _m_run_text(parent, text):
    """m:r with m:rPr/m:sty val='p' — upright (non-italic) math text."""
    r = etree.SubElement(parent, f'{{{MATH_NS}}}r')
    rPr = etree.SubElement(r, f'{{{MATH_NS}}}rPr')
    sty = etree.SubElement(rPr, f'{{{MATH_NS}}}sty')
    sty.set(f'{{{MATH_NS}}}val', 'p')
    t = etree.SubElement(r, f'{{{MATH_NS}}}t')
    t.set(f'{{{XML_NS}}}space', 'preserve')
    t.text = text
    return r


def _m_box(parent, text):
    """Wrap text in a math box (m:box) inside an m:e for frac/rad args."""
    box = _m('box')
    e = etree.SubElement(box, f'{{{MATH_NS}}}e')
    # Recursively parse the inner text
    _render_math_content(e, text)
    return box


def _m_rad(parent, radicand, deg=None):
    """m:rad — square root. If deg is given, it's \\sqrt[deg]{radicand}."""
    rad = etree.SubElement(parent, f'{{{MATH_NS}}}rad')
    if deg is not None:
        deg_elem = etree.SubElement(rad, f'{{{MATH_NS}}}deg')
        _m_run(deg_elem, str(deg))
    e = etree.SubElement(rad, f'{{{MATH_NS}}}e')
    _render_math_content(e, radicand)


def _m_frac(parent, num_text, den_text):
    """m:f — fraction numerator/denominator."""
    f = etree.SubElement(parent, f'{{{MATH_NS}}}f')
    # numerator
    num = etree.SubElement(f, f'{{{MATH_NS}}}num')
    _render_math_content(num, num_text)
    # denominator
    den = etree.SubElement(f, f'{{{MATH_NS}}}den')
    _render_math_content(den, den_text)


def _m_sub(parent, base_text, sub_text, plain_base=False):
    sSub = etree.SubElement(parent, f'{{{MATH_NS}}}sSub')
    e = etree.SubElement(sSub, f'{{{MATH_NS}}}e')
    (_m_run_text if plain_base else _m_run)(e, base_text)
    sub = etree.SubElement(sSub, f'{{{MATH_NS}}}sub')
    _m_run(sub, sub_text)


def _m_sup(parent, base_text, sup_text, plain_base=False):
    sSup = etree.SubElement(parent, f'{{{MATH_NS}}}sSup')
    e = etree.SubElement(sSup, f'{{{MATH_NS}}}e')
    (_m_run_text if plain_base else _m_run)(e, base_text)
    sup = etree.SubElement(sSup, f'{{{MATH_NS}}}sup')
    _m_run(sup, sup_text)


def _m_sub_sup(parent, base_text, sub_text, sup_text, plain_base=False):
    sSubSup = etree.SubElement(parent, f'{{{MATH_NS}}}sSubSup')
    e = etree.SubElement(sSubSup, f'{{{MATH_NS}}}e')
    (_m_run_text if plain_base else _m_run)(e, base_text)
    sub = etree.SubElement(sSubSup, f'{{{MATH_NS}}}sub')
    _m_run(sub, sub_text)
    sup = etree.SubElement(sSubSup, f'{{{MATH_NS}}}sup')
    _m_run(sup, sup_text)


def _m_nary(parent, ch, sub_text, sup_text, e_text=""):
    """m:nary — ∑/∏/∫ 的上下限显示在符号正上/正下方（学术规范）。
    e_text 是被求和/连乘/积分表达式, 为空时 Word 会显示点状占位框。"""
    nary = etree.SubElement(parent, f'{{{MATH_NS}}}nary')
    naryPr = etree.SubElement(nary, f'{{{MATH_NS}}}naryPr')
    chr_el = etree.SubElement(naryPr, f'{{{MATH_NS}}}chr')
    chr_el.set(f'{{{MATH_NS}}}val', ch)
    # 上下限显式置于符号正上/正下方, 不依赖 Word 版本默认
    lim_loc = etree.SubElement(naryPr, f'{{{MATH_NS}}}limLoc')
    lim_loc.set(f'{{{MATH_NS}}}val', 'undOvr')
    grow = etree.SubElement(naryPr, f'{{{MATH_NS}}}grow')
    grow.set(f'{{{MATH_NS}}}val', '1')
    if sub_text:
        sub = etree.SubElement(nary, f'{{{MATH_NS}}}sub')
        _render_math_content(sub, sub_text)
    if sup_text:
        sup = etree.SubElement(nary, f'{{{MATH_NS}}}sup')
        _render_math_content(sup, sup_text)
    e = etree.SubElement(nary, f'{{{MATH_NS}}}e')
    if e_text:
        _render_math_content(e, e_text)


def _latex_to_unicode(s):
    """LaTeX commands → Unicode for direct OMML rendering."""
    for cmd, uni in _LATEX_UNICODE.items():
        s = s.replace(cmd, uni)
    for cmd, diac in _DIACRITICS.items():
        # \\bar{x} → x + combining macron
        s = re.sub(rf'{re.escape(cmd)}\{{(\w)\}}', rf'\1{diac}', s)
    # 注意: \\text{...} 不在此剥壳, 由 _render_math_content 生成正体 run
    return s


def _find_matching_brace(s, start):
    """Given s[start] == '{', return index of matching '}'."""
    depth = 0
    for i in range(start, len(s)):
        if s[i] == '{':
            depth += 1
        elif s[i] == '}':
            depth -= 1
            if depth == 0:
                return i
    return -1


def _render_math_content(parent, text):
    """Render LaTeX-like math text into an OMML parent element.
    Handles: word_{sub}, word^{sup}, word_{sub}^{sup},
             \\frac{a}{b}, \\sqrt{x}, \\sqrt[n]{x},
             LaTeX Unicode commands, combining diacritics.
    """
    s = _latex_to_unicode(text)
    i = 0
    while i < len(s):
        # ── \\text{...} → 正体 run, 支持紧跟的 _{sub}/^{sup}（须在 \\frac 之前）──
        m = re.match(r'\\text\{([^}]*)\}', s[i:])
        if m:
            base = m.group(1)
            j = i + len(m.group(0))
            # 正体 base 上的上下标: \text{GDP}_{t}、\text{RWI}^{c}
            ms = re.match(r'(_\{[^}]+\}|\^\{[^}]+\})+', s[j:])
            if ms:
                seg = ms.group(0)
                subs = re.findall(r'_\{([^}]+)\}', seg)
                sups = re.findall(r'\^\{([^}]+)\}', seg)
                if subs and sups:
                    _m_sub_sup(parent, base, subs[0], sups[0], plain_base=True)
                elif subs:
                    _m_sub(parent, base, subs[0], plain_base=True)
                else:
                    _m_sup(parent, base, sups[0], plain_base=True)
                j += len(seg)
            else:
                _m_run_text(parent, base)
            i = j
            continue

        # ── \\frac{num}{den} ──
        if s[i:].startswith('\\frac{') and i+6 < len(s):
            brace_start = i + 5  # after \\frac
            num_end = _find_matching_brace(s, brace_start)
            if num_end >= 0 and num_end + 1 < len(s) and s[num_end + 1] == '{':
                den_start = num_end + 1
                den_end = _find_matching_brace(s, den_start)
                if den_end >= 0:
                    num_text = s[brace_start+1:num_end]
                    den_text = s[den_start+1:den_end]
                    _m_frac(parent, num_text, den_text)
                    i = den_end + 1
                    continue

        # ── \\sqrt[n]{x} or \\sqrt{x} ──
        if s[i:].startswith('\\sqrt'):
            j = i + 5
            deg = None
            if j < len(s) and s[j] == '[':
                bracket_end = s.find(']', j)
                if bracket_end >= 0:
                    deg = s[j+1:bracket_end]
                    j = bracket_end + 1
            if j < len(s) and s[j] == '{':
                rad_end = _find_matching_brace(s, j)
                if rad_end >= 0:
                    _m_rad(parent, s[j+1:rad_end], deg)
                    i = rad_end + 1
                    continue

        # ── ∑_{sub}^{sup} / ∏ / ∫ → m:nary（须在 sSubSup 之前匹配）──
        m = re.match(r'([∑∏∫])_\{\s*([^}]+)\s*\}\^\{([^}]+)\}', s[i:])
        if not m:
            m = re.match(r'([∑∏∫])\^\{([^}]+)\}_\{\s*([^}]+)\s*\}', s[i:])
        if m:
            # 被求和/连乘/积分表达式作为 e 传入（到下一个顶层 + - = 为止），
            # 否则 e 为空, Word 显示点状占位方框
            j = i + len(m.group(0))
            depth = 0
            k = j
            while k < len(s):
                ch = s[k]
                if ch in '([{':
                    depth += 1
                elif ch in ')]}':
                    depth = max(0, depth - 1)
                elif depth == 0 and ch in '+-=':
                    break
                k += 1
            _m_nary(parent, m.group(1), m.group(2), m.group(3),
                    s[j:k].strip())
            i = k
            continue

        # ── word_{sub}^{sup} or word^{sup}_{sub} ──
        m = re.match(r'(\S+?)_\{([^}]+)\}\^\{([^}]+)\}', s[i:])
        if not m:
            m = re.match(r'(\S+?)\^\{([^}]+)\}_\{([^}]+)\}', s[i:])
        if m:
            # Determine sub/sup order: which brace comes first in the source
            src = s[i:i+len(m.group(0))]
            underscore_pos = src.index('_') if '_' in src else 999
            caret_pos = src.index('^') if '^' in src else 999
            if underscore_pos < caret_pos:
                _m_sub_sup(parent, m.group(1), m.group(2), m.group(3))
            else:
                _m_sub_sup(parent, m.group(1), m.group(3), m.group(2))
            i += len(m.group(0))
            continue

        # ── word_{sub} ──
        m = re.match(r'(\S+?)_\{([^}]+)\}', s[i:])
        if m:
            _m_sub(parent, m.group(1), m.group(2))
            i += len(m.group(0))
            continue

        # ── word^{sup} ──
        m = re.match(r'(\S+?)\^\{([^}]+)\}', s[i:])
        if m:
            _m_sup(parent, m.group(1), m.group(2))
            i += len(m.group(0))
            continue

        # ── single-char sub/sup: x_2, x^2 ──
        m = re.match(r'(\S+)_(\S)', s[i:])
        if m:
            _m_sub(parent, m.group(1), m.group(2))
            i += len(m.group(0))
            continue
        m = re.match(r'(\S+)\^(\S)', s[i:])
        if m:
            _m_sup(parent, m.group(1), m.group(2))
            i += len(m.group(0))
            continue

        # ── plain text run ──
        _m_run(parent, s[i])
        i += 1


def _add_display_equation(doc, latex_str, number=None, tab_pos=None):
    """Insert centred OMML display equation; 可选右端 (n) 编号。"""
    from docx.enum.text import WD_TAB_ALIGNMENT

    omath = _m('oMath')
    _render_math_content(omath, latex_str)
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(4)

    if number is not None and tab_pos is not None:
        # 带编号：左端 TAB 到居中制表位，公式用行内 OMML，末尾 TAB 到右对齐位放 (n)
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.tab_stops.add_tab_stop(int(tab_pos / 2),
                                                  WD_TAB_ALIGNMENT.CENTER)
        p.paragraph_format.tab_stops.add_tab_stop(tab_pos, WD_TAB_ALIGNMENT.RIGHT)
        p.add_run("\t")
        p._p.append(omath)
        r = p.add_run("\t(%d)" % number)
        set_cn_font(r, "宋体", Pt(10.5))
        return

    # 无编号：oMathPara 整段居中
    omath_para = etree.Element(f'{{{MATH_NS}}}oMathPara')
    para_pr = etree.SubElement(omath_para, f'{{{MATH_NS}}}oMathParaPr')
    jc = etree.SubElement(para_pr, f'{{{MATH_NS}}}jc')
    jc.set(f'{{{MATH_NS}}}val', 'center')
    omath_para.append(omath)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p._p.append(omath_para)


def _add_inline_omath(paragraph, math_text):
    """Add inline OMML into a paragraph.

    m:oMath 必须是 w:p 的直接子元素；塞进 w:r 会让 Word 静默丢弃（表现为空白）。
    """
    omath = _m('oMath')
    _render_math_content(omath, math_text)
    paragraph._p.append(omath)


# ═══════════════════════════════════════════════════════════════
# RTF table parser (esttab output → python-docx table data)
# ═══════════════════════════════════════════════════════════════

def _rtf_u2uni(match):
    n = int(match.group(1))
    if n < 0:
        n += 65536
    return chr(n)


def _clean_rtf_text(text):
    text = re.sub(r'\\u(-?\d{2,5})\?', _rtf_u2uni, text)
    text = re.sub(r'\{\\super\s+([^}]+)\}', r'\1', text)
    text = re.sub(r'\{\\[a-z]+\d*\s+([^}]*)\}', r'\1', text)
    text = text.replace('{', '').replace('}', '')
    text = re.sub(r'\\[a-z]+\-?\d*\s?', ' ', text)
    text = re.sub(r'\s+', ' ', text)
    return text.strip()


def _parse_rtf_table(filepath):
    """Parse esttab RTF → (headers, data_rows, notes, title, col_widths)."""
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()

    # title
    title = ""
    m = re.search(r'\{\\pard\\keepn\\ql\s(.+?)\\par\}', content)
    if not m:
        m = re.search(r'\{\\pard\\qc\\fs\d+\s(.+?)\\par\}', content)
    if m:
        title = _clean_rtf_text(m.group(1))

    # notes — find \pard\ql block containing 注 after last table row
    # esttab 输出用 RTF 转义 \\u27880 表示「注」; 手写夹具可能直接写 UTF-8「注」
    # 多行表注(「注：…\par 数据来源：…」)可在一个 \pard 块内(\par 分隔)或跨多个块:
    # 从含「注」的块起收集到最后一个块, 原正则到第一个 \par 就停导致后续行丢失
    notes = ""
    last_row = content.rfind('\\row')
    if last_row > 0:
        note_parts = []
        started = False
        for m in re.finditer(r'\{\\pard\\ql\\fs\d+\s(.+?)(?=\{\\pard|\Z)',
                             content[last_row:], re.S):
            block = m.group(1)
            if '\\u27880' in block or '注' in block:
                started = True
            if started:
                note_parts.append(block)
        if note_parts:
            notes = _clean_rtf_text(' '.join(note_parts))

    # rows + 列宽（cellx 是 twips 边界位置，累计值）
    raw_rows = []
    cell_edges = []
    BS = chr(92)
    segments = re.split('(' + re.escape(BS + 'trowd') + ')', content)
    for k in range(1, len(segments), 2):
        segment = segments[k] + (segments[k+1] if k+1 < len(segments) else '')
        row_end = segment.find(BS + 'row')
        if row_end < 0:
            continue
        row_text = segment[:row_end]
        ncols = len(re.findall(re.escape(BS + 'cellx'), row_text))
        if ncols == 0:
            continue
        widths = [int(x) for x in re.findall(re.escape(BS + 'cellx') + r'(\d+)', row_text)]
        if widths:
            cell_edges.append(widths)
        cell_parts = re.split(re.escape(BS + 'cell') + r'(?=\s|\\|}|$)', row_text)
        cells = []
        for part in cell_parts:
            part = part.strip()
            if not part:
                continue
            cm = re.search(re.escape(BS+'pard') + re.escape(BS+'intbl') +
                           re.escape(BS+'q') + r'[lrc]\s*(.+)$', part)
            cells.append(_clean_rtf_text(cm.group(1)) if cm else '')
        while len(cells) < ncols:
            cells.append('')
        if cells:
            raw_rows.append(cells[:ncols])

    if not raw_rows:
        return None, [], notes, title, []

    # 每列取最大右边界（表头行通常定义完整），相邻边界之差即列宽（twips）
    col_widths = []
    if cell_edges:
        ncol = max(len(r) for r in cell_edges)
        edges = []
        for j in range(ncol):
            vals = [r[j] for r in cell_edges if j < len(r)]
            edges.append(max(vals) if vals else 0)
        for j in range(ncol):
            w = edges[j] - (edges[j-1] if j > 0 else 0)
            col_widths.append(max(w, 0))
        if not any(col_widths):
            col_widths = []

    ncols = max(len(r) for r in raw_rows)

    # detect header
    headers = []
    data_start = 0
    for i, row in enumerate(raw_rows):
        if len(row) >= 2 and row[0]:
            headers = raw_rows[:i]
            data_start = i
            break
    if not headers:
        headers = raw_rows[:1]
        data_start = 1

    # merge multi-row headers
    if len(headers) >= 2:
        merged = []
        for j in range(ncols):
            parts = [h[j] for h in headers if j < len(h) and h[j]]
            merged.append('\n'.join(parts) if parts else '')
        header_row = merged
    elif len(headers) == 1:
        header_row = list(headers[0]) + [''] * (ncols - len(headers[0]))
    else:
        header_row = [''] * ncols

    # merge SE rows (empty first cell → standard errors)
    merged_rows = []
    for i in range(data_start, len(raw_rows)):
        row = list(raw_rows[i]) + [''] * (ncols - len(raw_rows[i]))
        if all(c == '' for c in row):
            continue
        if row[0] == '' and merged_rows:
            prev = merged_rows[-1]
            for j in range(1, ncols):
                if j < len(row) and row[j]:
                    se = row[j]
                    if not se.startswith('('):
                        se = '(' + se + ')'
                    prev[j] = prev[j] + '\n' + se if prev[j] else se
            continue
        merged_rows.append(row)

    return [header_row], merged_rows, notes, title, col_widths


# ═══════════════════════════════════════════════════════════════
# Table rendering (RTF → Word 三线表)
# ═══════════════════════════════════════════════════════════════

def _set_cell_font(cell, text, font_name="宋体", size=Pt(9), bold=False,
                   alignment=WD_ALIGN_PARAGRAPH.CENTER):
    """Set cell text with font; newlines → separate paragraphs."""
    p = cell.paragraphs[0]
    p.clear()
    p.alignment = alignment
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.0
    lines = text.split('\n') if text else ['']
    for k, line in enumerate(lines):
        if k > 0:
            p = cell.add_paragraph()
            p.alignment = alignment
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.0
        run = p.add_run(_fix_quotes(line))
        run.font.name = 'Times New Roman'
        run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
        run._element.rPr.rFonts.set(qn('w:hAnsi'), font_name)
        run.font.size = size
        run.font.bold = bold


def _set_cell_border(cell, **kwargs):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for edge, attrs in kwargs.items():
        if attrs:
            el = OxmlElement(f'w:{edge}')
            for key, val in attrs.items():
                el.set(qn(f'w:{key}'), str(val))
            tcBorders.append(el)
    tcPr.append(tcBorders)


def _add_rtf_table(doc, filepath, font_size=Pt(9), opts=None, with_title=True):
    """Parse esttab RTF and insert as 三线表（with_title=False 时忽略 RTF 自带标题）。"""
    opts = opts or {}
    headers, rows, notes, title, col_widths = _parse_rtf_table(filepath)
    if headers is None:
        return

    # table title from RTF
    if title and with_title:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(4)
        r = p.add_run(_fix_quotes(title))
        set_cn_font(r, opts.get("title_font", "黑体"),
                    Pt(opts.get("title_size", 10.5)), bold=opts.get("title_bold", True))

    all_rows = headers + rows
    ncols = len(headers[0])
    nrows = len(all_rows)

    table = doc.add_table(rows=nrows, cols=ncols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    if col_widths:
        # 用 RTF 里的 cellx 列宽（twips → cm），避免 Word 自动分配挤压
        for j, w in enumerate(col_widths):
            if w > 0 and j < len(table.columns):
                table.columns[j].width = Cm(w / 1440 * 2.54)

    for i, row_data in enumerate(all_rows):
        for j, val in enumerate(row_data):
            cell = table.cell(i, j)
            is_header = (i == 0)
            _set_cell_font(cell, val if val else "",
                           font_name=(opts.get("header_font", "黑体") if is_header
                                      else opts.get("font", "宋体")),
                           size=font_size,
                           bold=(opts.get("header_bold", True) if is_header else False),
                           alignment=WD_ALIGN_PARAGRAPH.CENTER)

    # 三线 borders
    for j in range(ncols):
        _set_cell_border(table.cell(0, j),
                         top={"sz": str(opts.get("rule_top", 12)), "val": "single",
                              "color": "000000"},
                         bottom={"sz": str(opts.get("rule_mid", 8)), "val": "single",
                                 "color": "000000"})
    for j in range(ncols):
        _set_cell_border(table.cell(nrows - 1, j),
                         bottom={"sz": str(opts.get("rule_bottom", 12)), "val": "single",
                                 "color": "000000"})

    if notes:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.line_spacing = 1.0
        r = p.add_run(_fix_quotes(notes))
        set_cn_font(r, opts.get("note_font", "宋体"), Pt(opts.get("note_size", 9)))


# ═══════════════════════════════════════════════════════════════
# Body paragraph (with inline math detection)
# ═══════════════════════════════════════════════════════════════

_MATH_PATTERN = re.compile(
    r'([A-Za-zα-ωΑ-Ωβγδε'
    r'λμσρπφθ∂ΣΠ'
    r'ΔΩ]+(?:_\{[^}]+\}|\^\{[^}]+\})+(?:_\{[^}]+\}|\^\{[^}]+\})*)'
)


_INLINE_MARK = re.compile(r"\{(i|b|fn):([^{}]*)\}")


def _iter_inline(text):
    """把文本切成 [('text'|'i'|'b'|'fn', 内容)]。

    `{i:…}` 斜体、`{b:…}` 加粗、`{fn:…}` 就地插脚注。
    花括号内不支持再嵌套花括号。
    """
    pos = 0
    for m in _INLINE_MARK.finditer(text):
        if m.start() > pos:
            yield ("text", text[pos:m.start()])
        yield (m.group(1), m.group(2))
        pos = m.end()
    if pos < len(text):
        yield ("text", text[pos:])


def _add_runs_with_math(p, text, font, size):
    """普通文本段：识别内联公式后写入 runs。"""
    parts = _MATH_PATTERN.split(text)
    for i, part in enumerate(parts):
        if not part:
            continue
        if i % 2 == 0:
            r = p.add_run(_fix_quotes(part))
            set_cn_font(r, font, size)
        else:
            if '\\' in part:
                print(f"[cnpaper] WARNING: 内联公式含未支持命令: {part!r} — 请改用 equation()")
            _add_inline_omath(p, part)


def _add_body_paragraph(doc, text, size=Pt(12), indent=Pt(24),
                        line_spacing=1.5, space_before=Pt(6), space_after=Pt(6),
                        font="宋体", fn_font="宋体", fn_size=7.5):
    """Body paragraph: justified, first-line indent, inline math + 行内标记。"""
    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = indent
    p.paragraph_format.line_spacing = line_spacing
    p.paragraph_format.space_before = space_before
    p.paragraph_format.space_after = space_after
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    for kind, seg in _iter_inline(text):
        if kind == "text":
            _add_runs_with_math(p, seg, font, size)
        elif kind == "fn":
            add_footnote(p, doc, seg, font=fn_font, size=fn_size)
        else:
            r = p.add_run(_fix_quotes(seg))
            set_cn_font(r, font, size, bold=(kind == "b"))
            if kind == "i":
                r.italic = True


# ═══════════════════════════════════════════════════════════════
# Main render loop
# ═══════════════════════════════════════════════════════════════

def render(json_path, output="论文.docx", **kwargs):
    """Render paper_content.json → formatted .docx.

       Convenience wrapper: reads JSON then calls render_from_dicts().
       For programmatic use, call render_from_dicts(sections, output, ...) directly.
    """
    with open(json_path, 'r', encoding='utf-8') as f:
        sections = json.load(f)
    render_from_dicts(sections, output, **kwargs)


# 参数默认值与预设已移至 references/style.py（单一事实来源）


def _stamp_doc_properties(doc):
    """写入正确的文档属性：生成时间 + 清空 python-docx 模板署名。

    python-docx 内置空白模板把 created/modified 写死为
    2013-12-23T23:15:00Z、author 写为 "python-docx"、app.xml 的
    Application 写为 "Microsoft Macintosh Word"；不覆盖的话，文件属性
    （Windows 资源管理器 → 属性 → 详细信息）会显示错误的时间与来源。
    """
    from datetime import datetime, timezone
    now = datetime.now(timezone.utc).replace(microsecond=0)
    cp = doc.core_properties
    cp.author = ""
    cp.last_modified_by = ""
    cp.comments = ""
    cp.created = now
    cp.modified = now

    # app.xml（extended properties）无公开 API，直接改 part 字节
    for part in doc.part.package.iter_parts():
        if part.content_type == (
            "application/vnd.openxmlformats-officedocument.extended-properties+xml"
        ):
            part._blob = part.blob.replace(
                b"<Application>Microsoft Macintosh Word</Application>",
                b"<Application>Microsoft Office Word</Application>",
            )
            break


def render_from_dicts(sections, output="论文.docx", *, spec=None,
                      style=None, report=False, **kwargs):
    """Render a list of section dicts → formatted .docx.

       Each section dict: {"title": str, "body": str, "level": int}
       Same format as paper_content.json, but passed as a Python list.

       spec: 规格名或 yaml 路径（默认 "base" = 通用版式 = 《经济研究》投稿体例，见 specs/base.yaml）；
             显式传入的 kwargs 覆盖 spec 里的同名项。
       可用 kwargs：tables_dir, figures_dir, table_map, 以及 specs/base.yaml
       中的任意字段（margin_*, body_size, heading1_size, table_font …）。
    """
    if style is not None:
        st = style if isinstance(style, Style) else Style(style)
    else:
        st = resolve(spec=spec, overrides=kwargs)
    if report:
        print(st.describe())
    p = st.to_params()
    tables_dir, figures_dir, table_map = p["tables_dir"], p["figures_dir"], p["table_map"]
    margin_top, margin_bottom = p["margin_top"], p["margin_bottom"]
    margin_left, margin_right = p["margin_left"], p["margin_right"]
    body_size, body_line_spacing, body_indent = (
        p["body_size"], _spacing(p["body_line_spacing"]), p["body_indent"])
    heading1_size, heading2_size, heading3_size = (p["heading1_size"], p["heading2_size"], p["heading3_size"])
    ref_size, ref_line_spacing = p["ref_size"], _spacing(p["ref_line_spacing"])
    table_font_size = p["table_font_size"]

    # ── 字体与元素级字号（Style 层）──
    f_title, size_title = p["title_font"], Pt(p["title_size"])
    f_sub, size_sub = p["subtitle_font"], Pt(p["subtitle_size"])
    f_abs_lbl, size_abs_lbl = p["abstract_label_font"], Pt(p["abstract_label_size"])
    f_abs, size_abs = p["abstract_font"], Pt(p["abstract_size"])
    f_kw_lbl, size_kw_lbl = p["keyword_label_font"], Pt(p["keyword_label_size"])
    f_kw, size_kw = p["keyword_font"], Pt(p["keyword_size"])
    f_body = p["body_font"]
    h1 = (p["heading1_font"], Pt(p["heading1_size"]), p["heading1_bold"], _align(p["heading1_align"]))
    h2 = (p["heading2_font"], Pt(p["heading2_size"]), p["heading2_bold"], _align(p["heading2_align"]))
    h3 = (p["heading3_font"], Pt(p["heading3_size"]), p["heading3_bold"], _align(p["heading3_align"]))
    f_ref = p["ref_font"]
    f_note, size_note, note_italic = p["note_font"], Pt(p["note_size"]), p["note_italic"]
    title_bold = p["title_bold"]
    subtitle_bold = p["subtitle_bold"]
    f_author, size_author, author_bold = (
        p["author_font"], Pt(p["author_size"]), p["author_bold"])
    abstract_label, keyword_label = p["abstract_label"], p["keyword_label"]
    abstract_indent = Pt(p["abstract_indent"])
    abstract_ls = _spacing(p["abstract_line_spacing"]) if p.get("abstract_line_spacing") else 1.0
    ref_indent = p["ref_indent"]
    auto_number = p["auto_number"]
    eq_number = p["equation_number"]
    ref_head = dict(font=p["ref_heading_font"], size=Pt(p["ref_heading_size"]),
                    bold=p["ref_heading_bold"], align=_align(p["ref_heading_align"]))
    fn_font, fn_size, fn_per_page = (
        p["footnote_font"], p["footnote_size"], p["footnote_per_page"])
    content_w = Cm(21.0 - margin_left - margin_right)
    body_sp_before = Pt(p["body_space_before"])
    body_sp_after = Pt(p["body_space_after"])

    h1_no = h2_no = eq_no = fig_no = tab_no = 0

    table_opts = dict(
        font=p["table_font"], header_font=p["table_header_font"],
        header_bold=p["table_header_bold"],
        rule_top=p["table_rule_top"], rule_mid=p["table_rule_mid"],
        rule_bottom=p["table_rule_bottom"],
        title_font=p["table_title_font"], title_size=p["table_title_size"],
        title_bold=p["table_title_bold"],
        note_font=p["table_note_font"], note_size=p["table_note_size"],
    )

    doc = Document()
    _stamp_doc_properties(doc)

    if p["page_number"]:
        _add_page_number(doc, p["page_number_font"], p["page_number_size"])
    set_footnote_props(doc, p["footnote_num_format"], fn_per_page)

    # page setup
    for section in doc.sections:
        section.page_width = Cm(21.0)
        section.page_height = Cm(29.7)
        section.top_margin = Cm(margin_top)
        section.bottom_margin = Cm(margin_bottom)
        section.left_margin = Cm(margin_left)
        section.right_margin = Cm(margin_right)

    # default style
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(body_size)
    style.element.rPr.rFonts.set(qn('w:eastAsia'), p["body_font"])
    style.paragraph_format.line_spacing = body_line_spacing

    # resolve table paths
    if table_map is None:
        table_map = {}

    def _resolve_table(key):
        if key in table_map:
            return os.path.join(tables_dir, table_map[key])
        return os.path.join(tables_dir, key + '.rtf')

    def _resolve_image(name):
        return os.path.join(figures_dir, name)

    # ---- loop ----
    for sec in sections:
        title = _fix_quotes(sec.get("title", ""))
        raw_title = title          # 编号前的原标题，用于分支判断
        body = _fix_quotes(sec.get("body", ""))
        level = sec.get("level", 0)
        kind = sec.get("kind")

        # ---- 结构化内容（脚注 / 图 / 表）----
        if kind == "footnote":
            if doc.paragraphs:
                add_footnote(doc.paragraphs[-1], doc, sec.get("text", ""),
                             font=fn_font, size=fn_size)
            continue

        if kind == "figure":
            fig_no += 1
            fp = os.path.join(figures_dir, sec.get("file", ""))
            if os.path.exists(fp):
                pp = doc.add_paragraph()
                pp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                pp.paragraph_format.line_spacing = 1.0
                pp.paragraph_format.keep_with_next = True
                pp.add_run().add_picture(fp, width=Cm(float(sec.get("width") or 10)))
            else:
                print(f"[cnpaper] WARNING: image file not found: {fp}")
            cap = sec.get("caption")
            if cap:
                cp = doc.add_paragraph()
                cp.alignment = h3[3]
                cp.paragraph_format.space_before = Pt(6)
                cp.paragraph_format.space_after = Pt(4)
                cp.paragraph_format.line_spacing = 1.0
                cr = cp.add_run("图%d　%s" % (fig_no, _fix_quotes(cap)))
                set_cn_font(cr, h3[0], h3[1], bold=h3[2])
            continue

        if kind == "table":
            tab_no += 1
            cap = sec.get("caption")
            if cap:
                cp = doc.add_paragraph()
                cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                cp.paragraph_format.space_before = Pt(10)
                cp.paragraph_format.space_after = Pt(4)
                cp.paragraph_format.line_spacing = 1.0
                cr = cp.add_run("表%d　%s" % (tab_no, _fix_quotes(cap)))
                set_cn_font(cr, table_opts["title_font"],
                            Pt(table_opts["title_size"]),
                            bold=table_opts["title_bold"])
            fp = _resolve_table(sec.get("key", ""))
            if os.path.exists(fp):
                _add_rtf_table(doc, fp, font_size=Pt(table_font_size),
                               opts=table_opts,
                               with_title=not cap)
            else:
                print(f"[cnpaper] WARNING: table file not found: {fp}")
            continue

        # ---- 自动编号（已有编号则跳过）----
        if (auto_number and title and level in (1, 2)
                and raw_title not in ("摘要", "关键词", "Keywords",
                                      "参考文献", "References")):
            if level == 1:
                h1_no += 1
                h2_no = 0
                if not re.match(r"^[一二三四五六七八九十]+、", title):
                    title = "%s、%s" % (_cn_num(h1_no), title)
            else:
                h2_no += 1
                if not re.match(r"^（[一二三四五六七八九十]+）", title):
                    title = "（%s）%s" % (_cn_num(h2_no), title)

        # --- title block (level 0 + has title + not abstract/keywords) ---
        if level == 0 and title and title not in ('关键词', 'Keywords') \
                and title != '摘要':
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_after = Pt(4)
            p.paragraph_format.line_spacing = 1.0
            r = p.add_run(title)
            set_cn_font(r, f_title, size_title, bold=title_bold)
            if body:
                p2 = doc.add_paragraph()
                p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p2.paragraph_format.space_after = Pt(12)
                p2.paragraph_format.line_spacing = 1.0
                r2 = p2.add_run(body)
                set_cn_font(r2, f_sub, size_sub)
                body = ""

        # --- equation ---
        elif body.startswith('EQUATION:'):
            eq = body.replace('EQUATION:', '').strip()
            eq_no += 1 if eq_number else 0
            _add_display_equation(doc, eq, number=eq_no if eq_number else None,
                                  tab_pos=content_w)
            continue

        # --- abstract ---
        elif title == '摘要':
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.first_line_indent = abstract_indent
            p.paragraph_format.space_before = Pt(6)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = abstract_ls
            r = p.add_run(abstract_label + '：')
            set_cn_font(r, f_abs_lbl, size_abs_lbl, bold=True)
            if body:
                r2 = p.add_run(body)
                set_cn_font(r2, f_abs, size_abs)
                body = ""

        # --- keywords ---
        elif title in ('关键词', 'Keywords'):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.first_line_indent = abstract_indent
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(4)
            p.paragraph_format.line_spacing = 1.0
            r = p.add_run(keyword_label + '：')
            set_cn_font(r, f_kw_lbl, size_kw_lbl, bold=True)
            if body:
                r2 = p.add_run(body)
                set_cn_font(r2, f_kw, size_kw)
                body = ""

        # --- 参考文献标题（小标题，不是一级标题样式）---
        elif level == 1 and raw_title in ("参考文献", "References"):
            rp = doc.add_paragraph()
            rp.alignment = ref_head["align"]
            rp.paragraph_format.space_before = Pt(12)
            rp.paragraph_format.space_after = Pt(6)
            rp.paragraph_format.line_spacing = 1.0
            rp.paragraph_format.keep_with_next = True
            rr = rp.add_run(raw_title)
            set_cn_font(rr, ref_head["font"], ref_head["size"],
                        bold=ref_head["bold"])

        # --- L1 heading ---
        elif level == 1 and title:
            p = doc.add_paragraph()
            p.alignment = h1[3]
            p.paragraph_format.space_before = Pt(12)
            p.paragraph_format.space_after = Pt(6)
            p.paragraph_format.line_spacing = 1.0
            p.paragraph_format.keep_with_next = True
            r = p.add_run(title)
            set_cn_font(r, h1[0], h1[1], bold=h1[2])

        # --- L2 heading ---
        elif level == 2 and title:
            p = doc.add_paragraph()
            p.alignment = h2[3]
            p.paragraph_format.space_before = Pt(6)
            p.paragraph_format.space_after = Pt(3)
            p.paragraph_format.line_spacing = 1.0
            p.paragraph_format.keep_with_next = True
            r = p.add_run(title)
            set_cn_font(r, h2[0], h2[1], bold=h2[2])

        # --- L3 heading (figure/table caption) ---
        elif level == 3 and title:
            p = doc.add_paragraph()
            p.alignment = h3[3]
            p.paragraph_format.space_before = Pt(10)
            p.paragraph_format.space_after = Pt(4)
            p.paragraph_format.line_spacing = 1.0
            r = p.add_run(title)
            set_cn_font(r, h3[0], h3[1], bold=h3[2])

        # ---- body text dispatch ----
        if not body:
            continue
        body = body.strip()
        if not body:
            continue

        if body.startswith("TABLE:"):
            key = body.replace("TABLE:", "").strip()
            fp = _resolve_table(key)
            if os.path.exists(fp):
                _add_rtf_table(doc, fp, font_size=Pt(table_font_size),
                               opts=table_opts)
            else:
                print(f"[cnpaper] WARNING: table file not found: {fp}")
            continue

        if body.startswith("IMAGE:"):
            # IMAGE:file.png 或 IMAGE:file.png:14（指定宽度 cm）
            # 从末尾分离宽度: 文件名本身可能含盘符冒号(C:\fig\a.png),
            # split(":") 会截断成 'C', 用「末尾冒号+数字」正则只在确有宽度时拆分
            spec = body.replace("IMAGE:", "").strip()
            m = re.match(r"^(.*?):(\d+(?:\.\d+)?)$", spec)
            name = (m.group(1) if m else spec).strip()
            width = Cm(float(m.group(2))) if m else Cm(10)
            fp = _resolve_image(name)
            if os.path.exists(fp):
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.line_spacing = 1.0
                p.paragraph_format.keep_with_next = True
                r = p.add_run()
                r.add_picture(fp, width=width)
            else:
                print(f"[cnpaper] WARNING: image file not found: {fp}")
            continue

        if body.startswith("KEYWORDS:"):
            continue

        if body.startswith("SUBTITLE:"):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_after = Pt(12)
            p.paragraph_format.line_spacing = 1.0
            r = p.add_run(body.replace("SUBTITLE:", "").strip())
            set_cn_font(r, f_sub, size_sub, bold=subtitle_bold)
            continue

        if body.startswith("AUTHOR:"):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_after = Pt(10)
            p.paragraph_format.line_spacing = 1.0
            r = p.add_run(body.replace("AUTHOR:", "").strip())
            set_cn_font(r, f_author, size_author, bold=author_bold)
            continue

        if body.startswith("REF:"):
            ref_text = body.replace("REF:", "").strip()
            p = doc.add_paragraph()
            p.paragraph_format.first_line_indent = Pt(ref_indent)
            p.paragraph_format.line_spacing = ref_line_spacing
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            for kind, seg in _iter_inline(ref_text):
                if kind == "text":
                    r = p.add_run(_fix_quotes(seg))
                    set_cn_font(r, f_ref, Pt(ref_size))
                elif kind == "fn":
                    add_footnote(p, doc, seg, font=fn_font, size=fn_size)
                else:
                    r = p.add_run(_fix_quotes(seg))
                    set_cn_font(r, f_ref, Pt(ref_size), bold=(kind == "b"))
                    if kind == "i":
                        r.italic = True
            continue

        if body.startswith("NOTE:"):
            p = doc.add_paragraph()
            r = p.add_run(body.replace("NOTE:", "").strip())
            set_cn_font(r, f_note, size_note)
            r.italic = note_italic
            continue

        # regular text — split on \n\n
        for para_text in body.split("\n\n"):
            para_text = para_text.strip()
            if not para_text:
                # 空段落也添加
                doc.add_paragraph()
                continue

            if para_text.startswith("EQUATION:"):
                eq = para_text.replace("EQUATION:", "").strip()
                eq_no += 1 if eq_number else 0
                _add_display_equation(doc, eq, number=eq_no if eq_number else None,
                                      tab_pos=content_w)
                continue

            if para_text.startswith("NOINDENT:"):
                text = para_text.replace("NOINDENT:", "").strip()
                p = doc.add_paragraph()
                p.paragraph_format.first_line_indent = Pt(0)
                p.paragraph_format.line_spacing = body_line_spacing
                p.paragraph_format.space_before = Pt(6)
                p.paragraph_format.space_after = Pt(6)
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                r = p.add_run(_fix_quotes(text))
                set_cn_font(r, f_body, Pt(body_size))
                continue

            # H1/H2/H3 hypothesis line → bold (H1N1、H2O 等不误判)
            if re.match(r'^H[1-4][ab]?[：:．.。\s]', para_text):
                p = doc.add_paragraph()
                p.paragraph_format.first_line_indent = Pt(body_indent)
                p.paragraph_format.line_spacing = body_line_spacing
                p.paragraph_format.space_before = body_sp_before
                p.paragraph_format.space_after = body_sp_after
                r = p.add_run(_fix_quotes(para_text))
                set_cn_font(r, f_body, Pt(body_size), bold=True)
            else:
                _add_body_paragraph(doc, para_text,
                                    size=Pt(body_size),
                                    indent=Pt(body_indent),
                                    line_spacing=body_line_spacing,
                                    space_before=body_sp_before,
                                    space_after=body_sp_after,
                                    font=f_body,
                                    fn_font=fn_font, fn_size=fn_size)

    doc.save(output)
    if not output.startswith("__"):  # suppress print during tests
        print(f"论文已保存: {output}")


# ==============================================================================
# 来自 paper.py
# ==============================================================================
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
            print(f"[cnpaper] WARNING: 未知参数被忽略: {sorted(unknown)}")
        render_kwargs = {k: v for k, v in kwargs.items()
                         if k in _EXPECTED_RENDER_KWARGS}
        # Merge output into kwargs — render_from_dicts takes output as positional arg2
        render_from_dicts(self.sections, output, **render_kwargs)


# ==============================================================================
# 来自 preview.py
# ==============================================================================
import os
import subprocess


_SOFFICE_CANDIDATES = [
    r"C:\Program Files\LibreOffice\program\soffice.com",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.com",
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    "/usr/bin/soffice",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
]


def find_soffice():
    for p in _SOFFICE_CANDIDATES:
        if os.path.exists(p):
            return p
    raise FileNotFoundError(
        "找不到 LibreOffice（soffice）。装一个，或改用 office-render skill。"
    )


def _run(cmd, timeout):
    r = subprocess.run(cmd, capture_output=True, timeout=timeout)
    out = (r.stdout or b"").decode("utf-8", "replace").strip()
    err = (r.stderr or b"").decode("utf-8", "replace").strip()
    return r.returncode, out, err


def docx_to_pdf(docx_path, outdir=None, timeout=180, workdir=None):
    """docx → pdf。workdir 用于放 LibreOffice 用户配置，必须可写。"""
    docx_path = os.path.abspath(docx_path)
    base = workdir or os.path.dirname(docx_path)
    outdir = outdir or os.path.join(base, ".pdf")
    os.makedirs(outdir, exist_ok=True)
    profile = os.path.join(base, ".lo_profile")
    os.makedirs(profile, exist_ok=True)

    cmd = [
        find_soffice(), "--headless", "--norestore", "--nologo",
        "-env:UserInstallation=file:///" + profile.replace("\\", "/"),
        "--convert-to", "pdf", "--outdir", outdir, docx_path,
    ]
    # 上一次转换的产物若还被占用（Win 上偶发），LibreOffice 会以 0x11b 写失败
    pdf = os.path.join(outdir, os.path.splitext(os.path.basename(docx_path))[0] + ".pdf")
    try:
        if os.path.exists(pdf):
            os.remove(pdf)
    except OSError:
        pass
    rc, out, err = _run(cmd, timeout)
    if rc != 0 or not os.path.exists(pdf):
        raise RuntimeError(
            "LibreOffice 转换失败 (rc=%s, pdf_exists=%s)\nstdout: %s\nstderr: %s\n"
            "若提示「无法完成 LibreOffice 用户安装」，说明配置目录不可写：%s"
            % (rc, os.path.exists(pdf), out[:500], err[:500], profile)
        )
    return pdf


def pdf_to_png(pdf_path, outdir, dpi=150, pages=None):
    import pymupdf
    os.makedirs(outdir, exist_ok=True)
    doc = pymupdf.open(pdf_path)
    idx = range(doc.page_count) if pages is None else pages
    out = []
    for i in idx:
        pix = doc[i].get_pixmap(dpi=dpi)
        path = os.path.join(outdir, "p%d.png" % (i + 1))
        pix.save(path)
        out.append(path)
    return out


def preview(docx_path, outdir=None, dpi=150, pages=None, workdir=None):
    """渲染预览图，返回 PNG 路径列表。"""
    docx_path = os.path.abspath(docx_path)
    base = workdir or os.path.dirname(docx_path)
    outdir = outdir or os.path.join(base, "preview")
    pdf = docx_to_pdf(docx_path, os.path.join(base, ".pdf"), workdir=base)
    return pdf_to_png(pdf, outdir, dpi=dpi, pages=pages)
