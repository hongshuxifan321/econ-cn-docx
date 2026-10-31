# -*- coding: utf-8 -*-
"""预览：docx → pdf（LibreOffice headless）→ 逐页 PNG（PyMuPDF）。

用法::

    from references.preview import preview
    pngs = preview("论文.docx")          # → preview/p1.png, p2.png ...

注意（Windows / 沙箱）
--------------------
1. 用 ``soffice.com``（控制台入口），``soffice.exe`` 是 GUI 子系统，拿不到退出码；
2. LibreOffice 的用户配置必须放在**可写目录**，放 %TEMP%（沙箱下是
   ``...\\Temp\\dsh-xxxx``）会直接弹「无法完成 LibreOffice 用户安装」并返回 77；
   这里默认用 docx 同级的 ``.lo_profile``；
3. 子进程输出按 UTF-8 解码，中文环境下 text=True 会用 GBK 解码而抛错。
"""

import os
import subprocess

__all__ = ["preview", "docx_to_pdf", "pdf_to_png", "find_soffice"]

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
