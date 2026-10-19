#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
文档读写工具 —— 不依赖任何 MCP，纯 Python 库实现

覆盖 Word(.docx) / Excel(.xlsx) / PowerPoint(.pptx) / PDF 四类文档的读与写，
并支持把老格式 .doc / .xls 转成新格式。

命令行：
    # 读
    python doc_io.py read  <文件> [--json]          # 自动识别格式，输出文本/JSON
    python doc_io.py info  <文件>                    # 只输出结构摘要

    # 写
    python doc_io.py write-docx <输出.docx> --md <内容.md> [--table <表格.json>] [--img a.png,b.png]
    python doc_io.py write-xlsx <输出.xlsx> --csv <数据.csv>     # 或 --json <行数组.json>
    python doc_io.py write-pptx <输出.pptx> --md <幻灯片.md>     # 用 --- 分隔页
    python doc_io.py write-pdf  <输出.pdf>  --md <内容.md>

    # 转换老格式
    python doc_io.py convert <源文件.doc|.xls> <目标.docx|.xlsx>

Python API：
    from doc_io import read_any, write_docx, write_xlsx, write_pptx, write_pdf
版本：v1.0 (2026-10-13)
"""

import io
import os
import sys
import csv
import json

try:
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    if hasattr(sys.stderr, 'reconfigure'):
        sys.stderr.reconfigure(encoding='utf-8')
except Exception:
    pass

BODY_FONT = "仿宋_GB2312"
HEAD_FONT = "黑体"
TABLE_FONT = "宋体"


# ==================== 读 ====================

def read_docx(path):
    """读 Word：正文段落 + 表格（按原始先后顺序）"""
    from docx import Document
    from docx.oxml.ns import qn
    doc = Document(path)
    out = []
    for child in doc.element.body.iterchildren():
        if child.tag == qn('w:p'):
            t = "".join(n.text or "" for n in child.iter(qn('w:t'))).strip()
            if t:
                out.append(t)
        elif child.tag == qn('w:tbl'):
            rows = []
            for tr in child.findall(qn('w:tr')):
                cells = ["".join(n.text or "" for n in tc.iter(qn('w:t'))).strip()
                         for tc in tr.findall(qn('w:tc'))]
                rows.append(cells)
            out.append({"__table__": rows})
    return out


def read_xlsx(path):
    """读 Excel：{工作表: [[单元格...], ...]}"""
    from openpyxl import load_workbook
    wb = load_workbook(path, data_only=True)
    data = {}
    for ws in wb.worksheets:
        data[ws.title] = [[c for c in row] for row in ws.iter_rows(values_only=True)]
    return data


def read_pptx(path):
    """读 PowerPoint：每页的文本行"""
    from pptx import Presentation
    prs = Presentation(path)
    slides = []
    for i, s in enumerate(prs.slides, 1):
        lines = []
        for shape in s.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    t = "".join(r.text for r in para.runs).strip()
                    if t:
                        lines.append(t)
        slides.append({"slide": i, "lines": lines})
    return slides


def read_pdf(path):
    """读 PDF：每页文本"""
    try:
        import pdfplumber
        with pdfplumber.open(path) as pdf:
            return [{"page": i, "text": (p.extract_text() or "")} for i, p in enumerate(pdf.pages, 1)]
    except ImportError:
        from pypdf import PdfReader
        return [{"page": i, "text": (p.extract_text() or "")}
                for i, p in enumerate(PdfReader(path).pages, 1)]


def read_any(path):
    """按扩展名自动选择读取方式。.doc/.xls 先转新格式再读。"""
    ext = os.path.splitext(path)[1].lower()
    if ext in ('.doc', '.xls'):
        path = convert_legacy(path)
        ext = os.path.splitext(path)[1].lower()
    if ext == '.docx':
        return read_docx(path)
    if ext in ('.xlsx', '.xlsm'):
        return read_xlsx(path)
    if ext == '.pptx':
        return read_pptx(path)
    if ext == '.pdf':
        return read_pdf(path)
    if ext in ('.txt', '.md', '.csv', '.json'):
        with open(path, 'r', encoding='utf-8-sig', errors='replace') as f:
            return f.read()
    raise ValueError(f"不支持的格式：{ext}")


# ==================== 写 ====================

def _load_table(v):
    """表格数据：既接受 JSON 字符串，也接受 JSON 文件路径。

    Windows/PowerShell 传命令行 JSON 时引号常被吃掉，所以推荐改用文件：
        --table table.json
    """
    if not v:
        return None
    if isinstance(v, str) and os.path.exists(v):
        with open(v, 'r', encoding='utf-8-sig') as f:
            return json.load(f)
    if isinstance(v, str):
        try:
            return json.loads(v)
        except json.JSONDecodeError as e:
            raise SystemExit(
                f"[ERROR] --table 既不是有效 JSON 也不是存在的文件：{v[:40]}...\n"
                f"        命令行引号可能被 shell 吃掉，建议把表格写进 .json 文件后传文件路径。\n"
                f"        ({e})")
    return v


def _set_font(run, font=BODY_FONT, size=12, bold=False):
    from docx.shared import Pt
    from docx.oxml.ns import qn
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font)
    run._element.rPr.rFonts.set(qn('w:ascii'), "Times New Roman")
    run._element.rPr.rFonts.set(qn('w:hAnsi'), "Times New Roman")


def write_docx(path, content="", table_data=None, image_paths=None):
    """写 Word。#/##/### 映射标题，其余按正文。table_data 为二维数组，图片可锚点插入。"""
    from docx import Document
    from docx.shared import Cm, Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn

    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = sec.bottom_margin = Cm(2.54)
    sec.left_margin = sec.right_margin = Cm(3.17)
    st = doc.styles['Normal']
    st.font.name = BODY_FONT
    st.font.size = Pt(12)
    st.element.rPr.rFonts.set(qn('w:eastAsia'), BODY_FONT)

    import re as _re
    anchor_re = _re.compile(r'^!\[[^\]]*\]\((.+?)\)\s*$')
    used_imgs = set()

    for line in (content or "").splitlines():
        s = line.rstrip()
        if not s.strip():
            continue
        m = anchor_re.match(s.strip())
        if m:                                   # 图片锚点
            p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.add_run().add_picture(m.group(1).strip(), width=Cm(14))
            used_imgs.add(m.group(1).strip())
            continue
        if s.startswith('### '):
            p = doc.add_paragraph(); _set_font(p.add_run(s[4:]), HEAD_FONT, 12, True)
        elif s.startswith('## '):
            p = doc.add_paragraph(); _set_font(p.add_run(s[3:]), HEAD_FONT, 14, True)
        elif s.startswith('# '):
            p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            _set_font(p.add_run(s[2:]), HEAD_FONT, 16, True)
        else:
            p = doc.add_paragraph()
            p.paragraph_format.first_line_indent = Pt(24)
            p.paragraph_format.line_spacing = Pt(24)
            _set_font(p.add_run(s))

    if table_data:
        rows = _load_table(table_data)
        t = doc.add_table(rows=len(rows), cols=len(rows[0]))
        t.style = 'Table Grid'
        for ri, row in enumerate(rows):
            for ci, v in enumerate(row[:len(rows[0])]):
                c = t.cell(ri, ci); c.text = ""
                _set_font(c.paragraphs[0].add_run(str(v)), TABLE_FONT, 10.5, ri == 0)

    for img in (image_paths or ([] if isinstance(image_paths, str) else [])):
        img = img.strip()
        if img and img not in used_imgs and os.path.exists(img):
            doc.add_picture(img, width=Cm(14))
    if isinstance(image_paths, str):
        for img in [x.strip() for x in image_paths.split(',') if x.strip()]:
            if img not in used_imgs and os.path.exists(img):
                doc.add_picture(img, width=Cm(14))

    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    doc.save(path)
    return path


def write_xlsx(path, rows=None, csv_path=None):
    """写 Excel：rows 为二维数组，或从 CSV 读"""
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, Border, Side
    if csv_path:
        with open(csv_path, 'r', encoding='utf-8-sig', newline='') as f:
            rows = [r for r in csv.reader(f)]
    rows = rows or []
    wb = Workbook(); ws = wb.active
    thin = Side(style='thin'); bd = Border(left=thin, right=thin, top=thin, bottom=thin)
    for ri, row in enumerate(rows, 1):
        for ci, v in enumerate(row, 1):
            c = ws.cell(row=ri, column=ci, value=v)
            c.border = bd
            c.font = Font(name='宋体', size=11, bold=(ri == 1))
            c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    wb.save(path)
    return path


def write_pptx(path, content=""):
    """写 PPT：内容用单独一行的 --- 分页，每页第一行为标题，其余为正文"""
    from pptx import Presentation
    from pptx.util import Inches, Pt
    prs = Presentation()
    blank = prs.slide_layouts[6]
    pages = [p for p in (content or "").split('\n---\n') if p.strip()] or [""]
    for pg in pages:
        lines = [l.strip() for l in pg.strip().splitlines() if l.strip()]
        slide = prs.slides.add_slide(blank)
        if not lines:
            continue
        tb = slide.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(8.4), Inches(1))
        tf = tb.text_frame; tf.text = lines[0]
        tf.paragraphs[0].runs[0].font.size = Pt(28)
        tf.paragraphs[0].runs[0].font.bold = True
        if len(lines) > 1:
            b = slide.shapes.add_textbox(Inches(0.9), Inches(1.7), Inches(8.2), Inches(5))
            bf = b.text_frame
            for i, l in enumerate(lines[1:]):
                para = bf.paragraphs[0] if i == 0 else bf.add_paragraph()
                para.text = l
                for r in para.runs:
                    r.font.size = Pt(16)
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    prs.save(path)
    return path


def write_pdf(path, content="", table_data=None):
    """写 PDF（支持中文，使用 reportlab 内置 CJK 字体，无需外部字体文件）"""
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import cm
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib import colors

    pdfmetrics.registerFont(UnicodeCIDFont('STSong-Light'))
    body = ParagraphStyle('body', fontName='STSong-Light', fontSize=11, leading=20)
    head = ParagraphStyle('head', fontName='STSong-Light', fontSize=15, leading=24, spaceAfter=8)

    flow = []
    for line in (content or "").splitlines():
        s = line.strip()
        if not s:
            continue
        if s.startswith('#'):
            flow.append(Paragraph(s.lstrip('# ').replace('&', '&amp;').replace('<', '&lt;'), head))
        else:
            flow.append(Paragraph(s.replace('&', '&amp;').replace('<', '&lt;'), body))
        flow.append(Spacer(1, 4))
    if table_data:
        rows = _load_table(table_data)
        t = Table([[Paragraph(str(c), body) for c in r] for r in rows], repeatRows=1)
        t.setStyle(TableStyle([
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('BACKGROUND', (0, 0), (-1, 0), colors.whitesmoke),
        ]))
        flow.append(t)

    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    SimpleDocTemplate(path, pagesize=A4,
                      topMargin=2.54 * cm, bottomMargin=2.54 * cm,
                      leftMargin=3.17 * cm, rightMargin=3.17 * cm).build(flow)
    return path


# ==================== 老格式转换 ====================

def convert_legacy(src, dst=None):
    """把 .doc / .xls 转成 .docx / .xlsx。

    优先用 Windows 的 Word/Excel COM；失败则尝试 LibreOffice；都不行则报错提示。
    """
    ext = os.path.splitext(src)[1].lower()
    if dst is None:
        new_ext = '.docx' if ext == '.doc' else '.xlsx'
        dst = os.path.splitext(src)[0] + '_converted' + new_ext
    if os.path.abspath(src) == os.path.abspath(dst):
        return dst

    # 方案一：Word / Excel COM（Windows + 装了 Office）
    try:
        import win32com.client  # noqa
        if ext == '.doc':
            app = win32com.client.Dispatch('Word.Application')
            app.Visible = False; app.DisplayAlerts = 0
            d = app.Documents.Open(os.path.abspath(src), ReadOnly=True)
            d.SaveAs2(os.path.abspath(dst), 16)
            d.Close(False); app.Quit()
        else:
            app = win32com.client.Dispatch('Excel.Application')
            app.Visible = False; app.DisplayAlerts = 0
            wb = app.Workbooks.Open(os.path.abspath(src), ReadOnly=True)
            wb.SaveAs(os.path.abspath(dst), 51)
            wb.Close(False); app.Quit()
        if os.path.exists(dst):
            return dst
    except Exception as e:
        print(f"[INFO] COM 转换不可用（{e}），尝试 LibreOffice...", file=sys.stderr)

    # 方案二：LibreOffice
    import shutil
    import subprocess
    soffice = shutil.which('soffice') or shutil.which('libreoffice')
    for cand in (r"C:\Program Files\LibreOffice\program\soffice.exe",
                 r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"):
        if not soffice and os.path.exists(cand):
            soffice = cand
    if soffice:
        outdir = os.path.dirname(os.path.abspath(dst))
        target = 'docx' if ext == '.doc' else 'xlsx'
        subprocess.run([soffice, '--headless', '--convert-to', target,
                        '--outdir', outdir, os.path.abspath(src)],
                       check=False, timeout=180)
        auto = os.path.join(outdir, os.path.splitext(os.path.basename(src))[0] + '.' + target)
        if os.path.exists(auto):
            if os.path.abspath(auto) != os.path.abspath(dst):
                shutil.move(auto, dst)
            return dst

    raise RuntimeError(
        f"无法转换 {src}：本机既没有 Microsoft Office（COM 不可用），也没有 LibreOffice。\n"
        f"请安装其中之一，或先用其他工具把 {ext} 另存为 "
        f"{'.docx' if ext == '.doc' else '.xlsx'} 后再处理。")


# ==================== 命令行 ====================

def main():
    import argparse
    ap = argparse.ArgumentParser(description='文档读写工具（docx/xlsx/pptx/pdf，不依赖 MCP）')
    sub = ap.add_subparsers(dest='cmd', required=True)

    p = sub.add_parser('read'); p.add_argument('path'); p.add_argument('--json', action='store_true')
    p = sub.add_parser('info'); p.add_argument('path')
    p = sub.add_parser('write-docx'); p.add_argument('out')
    p.add_argument('--md', default=None); p.add_argument('--table', default=None)
    p.add_argument('--img', default=None)
    p = sub.add_parser('write-xlsx'); p.add_argument('out')
    p.add_argument('--csv', default=None); p.add_argument('--json', dest='rows_json', default=None)
    p = sub.add_parser('write-pptx'); p.add_argument('out'); p.add_argument('--md', default=None)
    p = sub.add_parser('write-pdf'); p.add_argument('out')
    p.add_argument('--md', default=None); p.add_argument('--table', default=None)
    p = sub.add_parser('convert'); p.add_argument('src'); p.add_argument('dst')

    a = ap.parse_args()

    if a.cmd == 'read':
        data = read_any(a.path)
        if a.json or not isinstance(data, str):
            print(json.dumps(data, ensure_ascii=False, indent=2, default=str))
        else:
            print(data)
    elif a.cmd == 'info':
        data = read_any(a.path)
        if isinstance(data, str):
            print(f"文本，{len(data)} 字符")
        elif isinstance(data, list):
            n_tbl = sum(1 for x in data if isinstance(x, dict) and '__table__' in x)
            print(f"列表：{len(data)} 项，其中表格 {n_tbl} 个")
        else:
            print(f"字典：键={list(data.keys())[:10]}")
    elif a.cmd == 'write-docx':
        md = open(a.md, 'r', encoding='utf-8-sig').read() if a.md else ''
        print("已写入：", write_docx(a.out, md, a.table, a.img))
    elif a.cmd == 'write-xlsx':
        rows = json.load(open(a.rows_json, encoding='utf-8-sig')) if a.rows_json else None
        print("已写入：", write_xlsx(a.out, rows, a.csv))
    elif a.cmd == 'write-pptx':
        md = open(a.md, 'r', encoding='utf-8-sig').read() if a.md else ''
        print("已写入：", write_pptx(a.out, md))
    elif a.cmd == 'write-pdf':
        md = open(a.md, 'r', encoding='utf-8-sig').read() if a.md else ''
        print("已写入：", write_pdf(a.out, md, a.table))
    elif a.cmd == 'convert':
        print("已转换：", convert_legacy(a.src, a.dst))


if __name__ == '__main__':
    main()
