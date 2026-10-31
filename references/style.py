# -*- coding: utf-8 -*-
"""Style：把排版参数从代码里搬到可编辑的规格文件。

设计要点
--------
1. 一份通用版式（base.yaml）+ 若干差异规格（erj.yaml），用 ``extends`` 继承；
2. 用户自定义规格写在 ``~/.dsh/econ-cn-docx/specs/``，skill 升级不会冲掉；
3. 每个值带来源（默认 / 预设 / 规格 / 调用参数），``describe()`` 打印生效值，
   避免「改了没生效」这种静默失败。

用法::

    from references.style import Style, load_spec, resolve

    st = Style.load(spec="erj")
    st.describe()
    render_from_dicts(sections, "论文.docx", spec="erj")
"""

import os

import yaml

__all__ = [
    "DEFAULTS", "Style", "load_spec", "resolve",
    "USER_SPEC_DIR", "SKILL_SPEC_DIR", "list_specs", "save",
]

USER_SPEC_DIR = os.path.expanduser("~/.dsh/econ-cn-docx/specs")
SKILL_SPEC_DIR = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "specs"
)

# ═══════════════════════════════════════════════════════════════
# 参数全集（单一事实来源：pipeline 从这里取默认值）
# ═══════════════════════════════════════════════════════════════

DEFAULTS = dict(
    # 目录与表格映射
    tables_dir="tables", figures_dir="figures", table_map=None,
    # 页面
    margin_top=2.54, margin_bottom=2.54, margin_left=3.18, margin_right=2.54,
    # 标题块
    title_size=18, title_font="黑体", title_bold=True,
    subtitle_size=14, subtitle_font="黑体", subtitle_bold=False,
    author_size=12, author_font="宋体", author_bold=False,
    # 摘要 / 关键词
    abstract_label="摘要", abstract_label_font="黑体", abstract_label_size=10.5,
    abstract_font="楷体", abstract_size=10.5, abstract_indent=0,
    keyword_label="关键词", keyword_label_font="黑体", keyword_label_size=10.5,
    keyword_font="楷体", keyword_size=10.5,
    # 正文
    body_size=12, body_font="宋体", body_line_spacing=1.5, body_indent=24,
    body_space_before=6, body_space_after=6,
    # 各级标题
    heading1_size=16, heading1_font="黑体", heading1_bold=True, heading1_align="center",
    heading2_size=14, heading2_font="黑体", heading2_bold=True, heading2_align="left",
    heading3_size=10.5, heading3_font="黑体", heading3_bold=True, heading3_align="center",
    # 表格
    table_font_size=9, table_font="宋体", table_header_font="黑体", table_header_bold=True,
    table_rule_top=12, table_rule_mid=8, table_rule_bottom=12,
    table_title_size=10.5, table_title_font="黑体", table_title_bold=True,
    table_note_size=9, table_note_font="宋体",
    # 表注 / 脚注样式（NOTE: 行）
    note_size=9, note_font="宋体", note_italic=True,
    # 参考文献
    ref_size=10, ref_font="宋体", ref_line_spacing=1.25, ref_indent=0,
    # 页码
    page_number=False, page_number_size=9, page_number_font="宋体",
    # 自动编号（一、/（一）/图1/表1/公式(1)）
    auto_number=True,
    equation_number=False,
    # 参考文献标题
    ref_heading_font="黑体", ref_heading_size=10.5, ref_heading_bold=True,
    ref_heading_align="left",
    # 脚注
    footnote_font="宋体", footnote_size=7.5, footnote_per_page=True,
    footnote_num_format="decimal",   # decimal / decimalEnclosedCircle / chicago
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
    raise FileNotFoundError(
        "找不到规格 '%s'，可用: %s" % (name, ", ".join(list_specs()))
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

    不指定 spec 就用 base.yaml。
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
    return path
