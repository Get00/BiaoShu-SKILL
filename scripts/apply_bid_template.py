#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
底板填空工具 —— 按招标文件"响应文件格式"生成响应文件

核心思想：招标文件附有响应文件格式时，那是"填空题"。
本工具直接把招标文件中的格式章节**整段复制**为新文档底板，
字体、字号、加粗、行距、缩进、下划线、说明文字、操作提示行全部原样保留，
然后只在【下划线空白 run】【指定文本】【表格空白单元格】三处填入本企业信息。

禁止：脱离底板从零重建文档；改动题干文字；套用通用标书格式。

用法：
    # 1) 先看看底板里有哪些空要填
    python apply_bid_template.py "采购文件.docx" --section "响应文件格式" --list-blanks

    # 2) 按 fields.json 填空
    python apply_bid_template.py "采购文件.docx" --section "响应文件格式" \
        --fields fields.json --output "响应文件.docx"

fields.json 格式：
{
  "blanks": ["XX省电力有限公司（采购人）", "XX技术有限公司", "..."],
  "replacements": [
      {"find": "致：采购人", "replace": "致：XX省电力有限公司（采购人）"}
  ],
  "cells": [
      {"table": 0, "row": 1, "col": 2, "value": "158.00"}
  ]
}

版本：v1.0 (2026-10-13)
"""

import io
import sys
import os
import re
import json
import copy
import argparse

from docx import Document
from docx.oxml.ns import qn

try:
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    if hasattr(sys.stderr, 'reconfigure'):
        sys.stderr.reconfigure(encoding='utf-8')
except Exception:
    pass

BLANK_RE = re.compile(r'^[\s_＿\-—﹍﹏　]*$')          # 全是下划线/空格/破折号 → 空白填空位
USCORE_RE = re.compile(r'[_＿﹍﹏]{2,}')               # 连续两条以上下划线 → 内嵌填空位

# 章号写法：第X章 / 第X卷 / 第X篇 / 第X节 / 第X部分（不同招标文件叫法不同）
CHAPTER_RE = r'第[一二三四五六七八九十百0-9]+(?:部分|[章卷篇节])'
NUM_RE = r'第[一二三四五六七八九十百0-9]+'


def _is_blank(text: str) -> bool:
    return bool(BLANK_RE.match(text or ""))


def _iter_blocks(doc):
    """按文档顺序产出 ('p'|'tbl', element)"""
    body = doc.element.body
    for child in body.iterchildren():
        if child.tag == qn('w:p'):
            yield 'p', child
        elif child.tag == qn('w:tbl'):
            yield 'tbl', child


def _para_text(p_el) -> str:
    return "".join(t.text or "" for t in p_el.iter(qn('w:t')))


def find_start_index(texts, section_re):
    """定位格式章节起点。

    章标题在文档里通常出现两次——目录页一次、正文一次，且正文中还可能有
    “见第五章响应文件格式”这类引用。因此：
      1) 先用“整章标题”严格匹配，命中多个时取**最后一个**（即正文真章标题）；
      2) 再退回“含章号的引用”匹配，同样取最后一个；
      3) 最后才用宽松子串匹配，取第一个。
    """
    patterns = [
        (rf'^\s*{CHAPTER_RE}\s*{section_re}\s*$', True),
        (rf'{CHAPTER_RE}\s*{section_re}', True),
        (section_re, False),
    ]
    for pat, take_last in patterns:
        hits = [i for i, t in enumerate(texts) if t and re.search(pat, t)]
        if hits:
            return hits[-1] if take_last else hits[0]
    return None


def find_section_bounds(src_doc, section_re: str, end_re: str = None):
    """定位格式章节的起止块元素。

    结束位置默认停在**同级**的下一个标题处：起点是“第六章”就只认下一个“第X章”，
    不会把“第一部分 商务文件”这种下级标题误当边界。
    """
    blocks = list(_iter_blocks(src_doc))
    texts = [(_para_text(el) if kind == 'p' else '') for kind, el in blocks]

    start_idx = find_start_index(texts, section_re)
    if start_idx is None:
        return None, None, blocks

    if end_re:
        pat_end = end_re
    else:
        m = re.search(rf'{NUM_RE}(部分|[章卷篇节])', texts[start_idx])
        unit = m.group(1) if m else '章'
        pat_end = rf'^\s*{NUM_RE}{re.escape(unit)}\s*\S'

    end_idx = None
    for j in range(start_idx + 1, len(texts)):
        if texts[j] and re.search(pat_end, texts[j]):
            end_idx = j
            break
    return start_idx, end_idx, blocks


def build_from_section(src_doc, blocks, start_idx, end_idx) -> Document:
    """把源文档中 [start_idx, end_idx) 的全部块深拷贝为新文档底板"""
    new_doc = Document()
    body = new_doc.element.body
    for child in list(body):
        body.remove(child)

    end = len(blocks) if end_idx is None else end_idx
    for kind, el in blocks[start_idx:end]:
        body.append(copy.deepcopy(el))

    src_sectpr = src_doc.element.body.find(qn('w:sectPr'))
    if src_sectpr is not None:
        body.append(copy.deepcopy(src_sectpr))
    return new_doc


def list_blanks(doc):
    """列出文档中的填空位，供人工核对"""
    out = []
    for i, p in enumerate(doc.paragraphs):
        ctx = p.text.strip()
        for j, r in enumerate(p.runs):
            if (r.font.underline and _is_blank(r.text)) or USCORE_RE.search(r.text or ""):
                out.append({
                    "para": i,
                    "run": j,
                    "kind": "下划线空白" if r.font.underline else "下划线字符",
                    "context": ctx[:60],
                    "raw": repr(r.text)[:40],
                })
    return out


def fill_blanks(doc, values):
    """按出现顺序，把 values 依次填入带下划线的空白 run（保留下划线格式）"""
    it = iter(values)
    n = 0
    for p in doc.paragraphs:
        for r in p.runs:
            if n >= len(values):
                return n
            if r.font.underline and _is_blank(r.text):
                r.text = str(next(it))
                n += 1
            elif USCORE_RE.search(r.text or ""):
                # 内嵌下划线（如 "致：___（采购人）"），只替换下划线那一段
                def _sub(m):
                    nonlocal n
                    if n >= len(values):
                        return m.group(0)
                    v = str(values[n]); n += 1
                    return v
                r.text = USCORE_RE.sub(_sub, r.text)
    return n


def apply_replacements(doc, reps):
    """文本替换：优先在单个 run 内替换（保留格式）；跨 run 则整段重写到首个 run"""
    hits = 0
    for rep in reps:
        find, repl = rep.get("find", ""), rep.get("replace", "")
        if not find:
            continue
        done = False
        for p in doc.paragraphs:
            for r in p.runs:
                if find in (r.text or ""):
                    r.text = r.text.replace(find, repl)
                    hits += 1
                    done = True
                    break
            if done:
                break
        if not done:  # 跨 run：整段合并写入首个 run
            for p in doc.paragraphs:
                if find in p.text:
                    runs = p.runs
                    if not runs:
                        continue
                    merged = p.text.replace(find, repl)
                    for r in runs[1:]:
                        r.text = ""
                    runs[0].text = merged
                    hits += 1
                    break
        if not done:
            print(f"  ⚠️  未找到待替换文本：{find}")
    return hits


def fill_cells(doc, cells):
    n = 0
    for c in cells:
        try:
            tb = doc.tables[int(c["table"])]
            cell = tb.rows[int(c["row"])].cells[int(c["col"])]
            if cell.paragraphs and cell.paragraphs[0].runs:
                cell.paragraphs[0].runs[0].text = str(c["value"])
                for r in cell.paragraphs[0].runs[1:]:
                    r.text = ""
            else:
                cell.text = str(c["value"])
            n += 1
        except Exception as e:
            print(f"  ⚠️  单元格填充失败 {c}: {e}")
    return n


def main():
    ap = argparse.ArgumentParser(description='底板填空工具（保留招标文件响应文件格式）')
    ap.add_argument('source', help='招标文件路径（含响应文件格式）')
    ap.add_argument('--section', default=r'响应文件格式',
                    help='格式章节起始锚点（正则，默认：响应文件格式）')
    ap.add_argument('--end', default=None,
                    help='格式章节结束锚点（正则，可选；默认自动停在同级下一个标题）')
    ap.add_argument('--fields', help='fields.json 路径')
    ap.add_argument('--output', '-o', help='输出 docx 路径')
    ap.add_argument('--list-blanks', action='store_true', help='仅列出填空位，不生成文件')
    args = ap.parse_args()

    if not os.path.exists(args.source):
        print(f"[ERROR] 文件不存在：{args.source}")
        sys.exit(1)

    src = Document(args.source)
    start_idx, end_idx, blocks = find_section_bounds(src, args.section, args.end)
    if start_idx is None:
        print(f"[ERROR] 未找到格式章节锚点：{args.section}")
        print("        请用 --section 指定正确的章节标题（如 “响应文件格式”“投标文件格式”）")
        sys.exit(2)

    doc = build_from_section(src, blocks, start_idx, end_idx)
    print(f"[INFO] 已复制底板：块 {start_idx} → {end_idx if end_idx is not None else len(blocks)}")
    print(f"[INFO] 底板段落数：{len(doc.paragraphs)}，表格数：{len(doc.tables)}")

    if args.list_blanks:
        blanks = list_blanks(doc)
        print(f"\n共发现 {len(blanks)} 个填空位：\n")
        for b in blanks:
            print(f"  段{b['para']:>4} run{b['run']:<3} [{b['kind']}] 原文={b['raw']:<12} 上下文：{b['context']}")
        print("\n把要填的内容按上述顺序写入 fields.json 的 \"blanks\" 数组即可。")
        return

    if not args.fields or not args.output:
        print("[ERROR] 生成文件需同时指定 --fields 和 --output")
        sys.exit(1)

    # utf-8-sig：兼容记事本/PowerShell 写出的带 BOM 的 JSON
    with open(args.fields, 'r', encoding='utf-8-sig') as f:
        fields = json.load(f)

    nb = fill_blanks(doc, fields.get("blanks", []))
    nr = apply_replacements(doc, fields.get("replacements", []))
    nc = fill_cells(doc, fields.get("cells", []))
    print(f"[INFO] 已填空：{nb} 个下划线位，{nr} 处文本替换，{nc} 个单元格")

    doc.save(args.output)
    print(f"✅ 响应文件已生成（格式保留招标文件原样）：{args.output}")
    print("   建议用 check_format_drift.py 复核格式是否与招标文件一致。")


if __name__ == '__main__':
    main()
