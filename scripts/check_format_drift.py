#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
格式漂移检查工具 —— 复核响应文件是否忠实保留了招标文件的响应文件格式

检查两种"对格式文件的实质性修改"（采购文件普遍列为否决情形）：
  A. 题干缺失或改动：招标文件格式里的说明文字、操作提示行被删改
  B. 格式漂移：字体、字号、加粗、下划线、行距与招标文件原格式不一致

用法：
    # 整体核对（拿整个格式章节 vs 整份响应文件）
    python check_format_drift.py "响应文件.docx" --source "采购文件.docx" --section "响应文件格式"

    # 分文件核对（把 --section 收窄到该部分的标题，减少无关条目干扰）
    python check_format_drift.py "01_商务文件.docx" --source "采购文件.docx" --section "二、商务文件"

说明：
    --section 优先按“第X章/卷/篇 响应文件格式”整章标题定位（目录页和正文各出现一次时取正文那次），
    默认值“响应文件格式”即可自动锁定格式章节。
    用整章锚点核对时，响应文件若只包含其中一部分（如仅商务文件），未被包含的部分会集中报
    “段落缺失”，属正常现象——按章节分别核对时请收窄 --section。

退出码：0=未发现漂移；1=发现漂移
版本：v1.0 (2026-10-13)
"""

import io
import sys
import os
import re
import difflib
import argparse

from docx import Document
from docx.oxml.ns import qn

try:
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
except Exception:
    pass

BLANK_OR_PUNCT = re.compile(r'[\s_＿﹍﹏—\-　]+')
CHAPTER_RE = r'第[一二三四五六七八九十百0-9]+(?:部分|[章卷篇节])\s*'
NUM_RE = r'第[一二三四五六七八九十百0-9]+'


def norm(t: str) -> str:
    return BLANK_OR_PUNCT.sub('', t or '')


def run_fmt(r):
    """返回 run 的格式指纹：(中文字体, 字号pt, 加粗, 下划线)"""
    rPr = r._element.rPr
    ea = None
    sz = None
    if rPr is not None:
        rf = rPr.find(qn('w:rFonts'))
        if rf is not None:
            ea = rf.get(qn('w:eastAsia'))
        s = rPr.find(qn('w:sz'))
        if s is not None:
            sz = int(s.get(qn('w:val'))) / 2.0
    return (ea, sz, bool(r.font.bold), bool(r.font.underline))


def para_line_spacing(p):
    ls = p.paragraph_format.line_spacing
    try:
        return round(ls.pt, 1) if hasattr(ls, 'pt') else ls
    except Exception:
        return ls


def _find_start(texts, section_re):
    """定位格式章节起点。

    章标题通常出现两次（目录页 + 正文），正文中还可能有“见第五章响应文件格式”这类引用，
    因此整章标题取**最后一次**匹配，宽松子串才取第一次。
    章号支持 第X章 / 第X卷 / 第X篇 / 第X节 / 第X部分。
    """
    patterns = [
        (rf'^\s*{CHAPTER_RE}{section_re}\s*$', True),
        (rf'{CHAPTER_RE}{section_re}', True),
        (section_re, False),
    ]
    for pat, take_last in patterns:
        hits = [i for i, t in enumerate(texts) if t and re.search(pat, t)]
        if hits:
            return hits[-1] if take_last else hits[0]
    return None


def section_paragraphs(doc, section_re, end_re=None):
    """取 doc 中格式章节的段落列表。

    结束边界与 extract_format / apply_bid_template 保持一致：默认自动停在**同级**的
    下一个标题处（起点是“第四卷”就只认下一个“第X卷”），避免把紧随其后的其他章节
    也算进来而误报“段落缺失”。显式传 end_re 时以 end_re 为准。
    """
    paras = doc.paragraphs
    texts = [p.text for p in paras]
    start = _find_start(texts, section_re)
    if start is None:
        return []

    if end_re:
        pat_end = end_re
    else:
        m = re.search(rf'{NUM_RE}(部分|[章卷篇节])', texts[start] or '')
        unit = m.group(1) if m else '章'
        pat_end = rf'^\s*{NUM_RE}{re.escape(unit)}\s*\S'

    out = []
    for p in paras[start + 1:]:
        if re.search(pat_end, p.text):
            break
        out.append(p)
    return out


def fragments(text):
    """按空白/下划线切分，取出长度>=6的非空片段（用于题干包含性检查）"""
    parts = re.split(r'[\s_＿﹍﹏　\t]+', text or '')
    return [p for p in parts if len(p) >= 6]


def check(target_path, source_path, section_re, end_re=None):
    if not os.path.exists(target_path):
        print(f"[ERROR] 找不到响应文件：{target_path}")
        return 2
    if not os.path.exists(source_path):
        print(f"[ERROR] 找不到招标文件：{source_path}")
        return 2

    src = Document(source_path)
    tgt = Document(target_path)

    src_paras = section_paragraphs(src, section_re, end_re)
    tgt_paras = section_paragraphs(tgt, section_re, end_re)
    if not src_paras:
        print(f"[ERROR] 招标文件中未按锚点找到格式章节：{section_re}")
        return 2

    # ---- 配对：段落数一致时按序号一一配对（底板填空的典型情况）；
    #      否则退回相似度匹配。按序号配对可避免重复文本（如“项目名称：，采购编号：…”）
    #      错位比对导致误报。
    by_index = len(src_paras) == len(tgt_paras)
    if by_index:
        pairs = list(zip(src_paras, tgt_paras))
    else:
        tgt_norms = [(p, norm(p.text)) for p in tgt_paras]
        pairs = []
        for sp in src_paras:
            sn = norm(sp.text)
            best, best_ratio = None, 0.0
            for tp, tn in tgt_norms:
                r = difflib.SequenceMatcher(None, sn, tn).ratio()
                if r > best_ratio:
                    best, best_ratio = tp, r
            pairs.append((sp, best if best_ratio >= 0.75 else None))

    missing, fmt_drift = [], []

    for sp, tp in pairs:
        s_text = sp.text
        # ---- A. 题干完整性：原文的非空片段必须全部出现在对应段落里
        if len(norm(s_text)) >= 8:
            if tp is None:
                missing.append((s_text.strip()[:70], "段落缺失"))
            else:
                t_norm = norm(tp.text)
                for frag in fragments(s_text):
                    if norm(frag) not in t_norm:
                        missing.append((frag[:60], "片段缺失"))
        if tp is None:
            continue

        # ---- B. 格式一致性：比较首个非空 run 的字体/字号/加粗/下划线
        s_fmts = [run_fmt(r) for r in sp.runs if (r.text or '').strip()]
        t_fmts = [run_fmt(r) for r in tp.runs if (r.text or '').strip()]
        if s_fmts and t_fmts and s_fmts[:1] != t_fmts[:1]:
            fmt_drift.append((s_text.strip()[:50], s_fmts[0], t_fmts[0]))
        sls, tls = para_line_spacing(sp), para_line_spacing(tp)
        if sls and tls and sls != tls:
            fmt_drift.append((s_text.strip()[:50], f"行距{sls}", f"行距{tls}"))

    print("=" * 68)
    print("格式漂移检查报告")
    print("=" * 68)
    print(f"招标文件格式段落数：{len(src_paras)}    响应文件对应段落数：{len(tgt_paras)}")
    print(f"配对方式：{'按序号一一配对' if by_index else '按文本相似度配对'}")

    print(f"\n【A. 题干完整性】疑似被删改：{len(missing)} 处")
    if missing:
        for t, why in missing[:40]:
            print(f"  ❌ [{why}] {t}")
        if len(missing) > 40:
            print(f"  ...（另有 {len(missing) - 40} 处）")
        print("     ↑ 这些文字在响应文件里找不到，说明题干被删掉或改写了。")
    else:
        print("  ✅ 未发现题干缺失或被改。")

    print(f"\n【B. 格式一致性】疑似漂移：{len(fmt_drift)} 处")
    if fmt_drift:
        for t, s, tt in fmt_drift[:40]:
            print(f"  ⚠️  {t}")
            print(f"       招标文件={s}  →  响应文件={tt}")
        if len(fmt_drift) > 40:
            print(f"  ...（另有 {len(fmt_drift) - 40} 处）")
    else:
        print("  ✅ 未发现字体/字号/加粗/下划线/行距漂移。")

    ok = not missing and not fmt_drift
    print("\n" + ("✅ 通过：响应文件忠实保留了招标文件格式。" if ok
                  else "❌ 未通过：存在题干删改或格式漂移，可能导致“对格式文件实质性修改”被否决。"))
    return 0 if ok else 1


def main():
    ap = argparse.ArgumentParser(description='复核响应文件是否忠实保留招标文件的响应文件格式')
    ap.add_argument('target', help='响应文件 docx 路径')
    ap.add_argument('--source', required=True, help='招标文件 docx 路径')
    ap.add_argument('--section', default=r'响应文件格式', help='格式章节锚点（正则）')
    ap.add_argument('--end', default=None, help='格式章节结束锚点（正则，可选）')
    args = ap.parse_args()
    sys.exit(check(args.target, args.source, args.section, args.end))


if __name__ == '__main__':
    main()
