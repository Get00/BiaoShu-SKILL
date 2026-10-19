#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
抓格式 —— 从招标文件中抽取"投标/响应文件格式"章节，原样复制为独立 Word 文档

整段深拷贝，字体、字号、加粗、行距、缩进、下划线、页码域全部保留，一字不改。

用法：
    python extract_format.py "招标文件.docx"                      # 自动找格式章节
    python extract_format.py "招标文件.docx" --section "响应文件格式"
    python extract_format.py "招标文件.docx" --section "第四卷 响应文件格式" --output "格式.docx"
    python extract_format.py "招标文件.docx" --list               # 只列出候选锚点
"""
import sys, os, re, copy, argparse
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

from docx import Document
from docx.oxml.ns import qn

# 章号写法：第X章 / 第X卷 / 第X篇 / 第X节 / 第X部分
CHAPTER_RE = r'第[一二三四五六七八九十百0-9]+(?:部分|[章卷篇节])\s*'
ANCHORS = ["响应文件格式", "投标文件格式", "应答文件格式", "响应文件组成", "投标文件组成"]


def iter_blocks(doc):
    for child in doc.element.body.iterchildren():
        if child.tag == qn('w:p'):
            yield 'p', child
        elif child.tag == qn('w:tbl'):
            yield 'tbl', child


def para_text(el):
    return "".join(t.text or "" for t in el.iter(qn('w:t')))


def find_start(texts, section_re):
    """定位章节起点：整章标题取最后一次匹配（避开目录页），宽松匹配取第一次"""
    pats = [
        (rf'^\s*{CHAPTER_RE}{section_re}\s*$', True),
        (rf'{CHAPTER_RE}{section_re}', True),
        (re.escape(section_re), False),
    ]
    for pat, take_last in pats:
        hits = [i for i, t in enumerate(texts) if t and re.search(pat, t)]
        if hits:
            return hits[-1] if take_last else hits[0], pat
    return None, None


def extract(src_path, section_re=None, output=None, end_re=None, list_only=False):
    doc = Document(src_path)
    blocks = list(iter_blocks(doc))
    texts = [(para_text(el) if k == 'p' else '') for k, el in blocks]

    # 候选锚点：指定优先，否则依次尝试常用名称
    cands = [section_re] if section_re else ANCHORS
    chosen = None
    for c in cands:
        idx, pat = find_start(texts, c)
        if idx is not None:
            chosen = (c, idx, pat)
            break

    if list_only or chosen is None:
        print(f"\n文件：{os.path.basename(src_path)}")
        print(f"  总块数 {len(blocks)}（段落 {len(doc.paragraphs)}，表格 {len(doc.tables)}）")
        print("  --- 候选锚点出现情况 ---")
        for c in ANCHORS:
            hits = [i for i, t in enumerate(texts) if t and c in t]
            if hits:
                print(f"    “{c}”：块 {hits}  首条：{texts[hits[0]].strip()[:50]}")
        if chosen:
            print(f"  ✅ 将采用：“{chosen[0]}”，起点块 {chosen[1]}")
        else:
            print("  ❌ 未找到格式章节锚点")
        return chosen

    name, start, pat = chosen
    # 结束位置：显式 --end 优先；否则自动停在**同级**的下一个标题处。
    # 起点是“第六章”就只认下一个“第X章”，不会把“第一部分 商务文件”这种下级标题误当边界。
    m = re.search(r'第[一二三四五六七八九十百0-9]+(部分|[章卷篇节])', texts[start])
    unit = m.group(1) if m else '章'
    pat_end = end_re or rf'^\s*第[一二三四五六七八九十百0-9]+{re.escape(unit)}\s*\S'
    end = len(blocks)
    for j in range(start + 1, len(blocks)):
        if texts[j] and re.search(pat_end, texts[j]):
            end = j
            break

    # 建底板：深拷贝 [start, end) 的全部块
    out = Document()
    body = out.element.body
    for child in list(body):
        body.remove(child)
    for k, el in blocks[start:end]:
        body.append(copy.deepcopy(el))
    src_sect = doc.element.body.find(qn('w:sectPr'))
    if src_sect is not None:
        body.append(copy.deepcopy(src_sect))

    if output is None:
        base = os.path.splitext(os.path.basename(src_path))[0]
        output = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                              f"{base}__投标文件格式.docx")
    out.save(output)

    print(f"\n✅ 已抓取：{os.path.basename(src_path)}")
    print(f"   锚点：“{name}”（匹配式 {pat}），起始块 {start}，结束块 {end}")
    print(f"   产出：段落 {len(out.paragraphs)}，表格 {len(out.tables)}")
    print(f"   文件：{output}")
    return chosen


def main():
    ap = argparse.ArgumentParser(description='从招标文件中抽取投标/响应文件格式')
    ap.add_argument('source', help='招标文件路径（.docx）')
    ap.add_argument('--section', '-s', default=None, help='格式章节标题（默认自动识别）')
    ap.add_argument('--end', '-e', default=None, help='结束锚点（正则，可选）')
    ap.add_argument('--output', '-o', default=None, help='输出 docx 路径')
    ap.add_argument('--list', '-l', action='store_true', help='只列出候选锚点，不产出文件')
    args = ap.parse_args()

    if not os.path.exists(args.source):
        print(f"[ERROR] 文件不存在：{args.source}")
        sys.exit(1)

    res = extract(args.source, args.section, args.output, args.end, args.list)
    sys.exit(0 if res else 2)


if __name__ == '__main__':
    main()
