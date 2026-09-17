#!/usr/bin/env python3
"""Repair native bibliography numbering and clickable cross-references.

Reference-list entries are native Word numbered-list items using ``[%1]`` and a
TAB suffix. Body citations such as ``[1]`` are converted to ``REF ... \\n \\h``
fields that read the referenced paragraph number, so reordering bibliography
entries updates citations after Word refreshes fields.
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
import tempfile
import zipfile
from copy import deepcopy
from pathlib import Path

from lxml import etree
from profile_config import load_profile, style, get

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
NS = {'w': W, 'm': M}


def qn(name: str) -> str:
    prefix, local = name.split(':', 1)
    return f'{{{NS[prefix]}}}{local}'


def text_of(p):
    return ''.join(p.xpath('.//w:t/text()', namespaces=NS))


def instructions(p):
    return [' '.join(x.split()) for x in p.xpath('.//w:instrText/text()', namespaces=NS)]


def style_id(p):
    vals = p.xpath('./w:pPr/w:pStyle/@w:val', namespaces=NS)
    return vals[0] if vals else ''


def style_maps(styles_root):
    by_id, by_name = {}, {}
    if styles_root is None:
        return by_id, by_name
    for st in styles_root.xpath('./w:style', namespaces=NS):
        sid = st.get(qn('w:styleId'), '')
        names = st.xpath('./w:name/@w:val', namespaces=NS)
        name = names[0] if names else sid
        by_id[sid] = name
        by_name[name.casefold()] = sid
    return by_id, by_name


def next_numeric_attr(nodes, attr_name: str) -> int:
    vals = []
    for n in nodes:
        v = n.get(qn(attr_name))
        if v and str(v).isdigit():
            vals.append(int(v))
    return max(vals, default=0) + 1


def ensure_reference_numbering(styles_root, numbering_root, profile) -> tuple[str, str]:
    """Ensure reference style uses a native list ``[1]`` + TAB."""
    if styles_root is None or numbering_root is None:
        raise RuntimeError('styles.xml and numbering.xml are required for native reference numbering')

    _, by_name = style_maps(styles_root)
    ref_style_name = style(profile, 'references', 'References')
    sid = by_name.get(ref_style_name.casefold())
    if not sid:
        raise RuntimeError(f'reference style not found: {ref_style_name}')
    st = next((x for x in styles_root.xpath('./w:style', namespaces=NS) if x.get(qn('w:styleId')) == sid), None)
    if st is None:
        raise RuntimeError(f'reference style not found: {ref_style_name}')

    fmt = str(get(profile, 'references', 'numbering', 'format', default='[%1]'))
    suffix = str(get(profile, 'references', 'numbering', 'suffix', default='tab'))
    tab_pos = int(get(profile, 'references', 'numbering', 'tab_position_twips', default=720))
    left = int(get(profile, 'references', 'numbering', 'left_indent_twips', default=720))
    hanging = int(get(profile, 'references', 'numbering', 'hanging_twips', default=720))

    current_num = st.xpath('./w:pPr/w:numPr/w:numId/@w:val', namespaces=NS)
    if current_num:
        num_id = current_num[0]
        abs_id = numbering_root.xpath(f'./w:num[@w:numId="{num_id}"]/w:abstractNumId/@w:val', namespaces=NS)
        if abs_id:
            lvl = numbering_root.xpath(f'./w:abstractNum[@w:abstractNumId="{abs_id[0]}"]/w:lvl[@w:ilvl="0"]', namespaces=NS)
            if lvl:
                text = lvl[0].xpath('./w:lvlText/@w:val', namespaces=NS)
                suff = lvl[0].xpath('./w:suff/@w:val', namespaces=NS)
                if text == [fmt] and (not suffix or suff == [suffix]):
                    return sid, num_id

    abstract_id = str(next_numeric_attr(numbering_root.xpath('./w:abstractNum', namespaces=NS), 'w:abstractNumId'))
    num_id = str(next_numeric_attr(numbering_root.xpath('./w:num', namespaces=NS), 'w:numId'))

    abstract = etree.SubElement(numbering_root, qn('w:abstractNum'))
    abstract.set(qn('w:abstractNumId'), abstract_id)
    multi = etree.SubElement(abstract, qn('w:multiLevelType'))
    multi.set(qn('w:val'), 'singleLevel')
    lvl = etree.SubElement(abstract, qn('w:lvl'))
    lvl.set(qn('w:ilvl'), '0')
    start = etree.SubElement(lvl, qn('w:start')); start.set(qn('w:val'), '1')
    num_fmt = etree.SubElement(lvl, qn('w:numFmt')); num_fmt.set(qn('w:val'), 'decimal')
    lvl_text = etree.SubElement(lvl, qn('w:lvlText')); lvl_text.set(qn('w:val'), fmt)
    suff = etree.SubElement(lvl, qn('w:suff')); suff.set(qn('w:val'), suffix)
    pstyle = etree.SubElement(lvl, qn('w:pStyle')); pstyle.set(qn('w:val'), sid)
    ppr = etree.SubElement(lvl, qn('w:pPr'))
    tabs = etree.SubElement(ppr, qn('w:tabs'))
    tab = etree.SubElement(tabs, qn('w:tab')); tab.set(qn('w:val'), 'num'); tab.set(qn('w:pos'), str(tab_pos))
    ind = etree.SubElement(ppr, qn('w:ind')); ind.set(qn('w:left'), str(left)); ind.set(qn('w:hanging'), str(hanging))

    num = etree.SubElement(numbering_root, qn('w:num'))
    num.set(qn('w:numId'), num_id)
    aid = etree.SubElement(num, qn('w:abstractNumId')); aid.set(qn('w:val'), abstract_id)

    ppr_style = st.find(qn('w:pPr'))
    if ppr_style is None:
        ppr_style = etree.SubElement(st, qn('w:pPr'))
    old = ppr_style.find(qn('w:numPr'))
    if old is not None:
        ppr_style.remove(old)
    numpr = etree.SubElement(ppr_style, qn('w:numPr'))
    ilvl = etree.SubElement(numpr, qn('w:ilvl')); ilvl.set(qn('w:val'), '0')
    nid = etree.SubElement(numpr, qn('w:numId')); nid.set(qn('w:val'), num_id)
    return sid, num_id


class BookmarkFactory:
    def __init__(self, root):
        ids = [int(v) for v in root.xpath('//w:bookmarkStart/@w:id', namespaces=NS) if str(v).isdigit()]
        self.next_id = max(ids, default=0) + 1
        self.names = set(root.xpath('//w:bookmarkStart/@w:name', namespaces=NS))

    def pair(self, preferred: str):
        name, suffix = preferred, 2
        while name in self.names:
            name = f'{preferred}_{suffix}'; suffix += 1
        self.names.add(name)
        value = str(self.next_id); self.next_id += 1
        start = etree.Element(qn('w:bookmarkStart')); start.set(qn('w:id'), value); start.set(qn('w:name'), name)
        end = etree.Element(qn('w:bookmarkEnd')); end.set(qn('w:id'), value)
        return start, end, name


def clone_rpr(src):
    return deepcopy(src) if src is not None else None


def text_run(text: str, rpr=None):
    r = etree.Element(qn('w:r'))
    if rpr is not None:
        r.append(clone_rpr(rpr))
    t = etree.SubElement(r, qn('w:t'))
    if text[:1].isspace() or text[-1:].isspace():
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    t.text = text
    return r


def field_runs(instruction: str, cached: str, rpr=None):
    out = []
    for kind, value in (('begin', None), (None, instruction), ('separate', None), (None, cached), ('end', None)):
        r = etree.Element(qn('w:r'))
        if rpr is not None:
            r.append(clone_rpr(rpr))
        if kind:
            node = etree.SubElement(r, qn('w:fldChar')); node.set(qn('w:fldCharType'), kind)
            if kind == 'begin':
                node.set(qn('w:dirty'), 'true')
        elif value == instruction:
            node = etree.SubElement(r, qn('w:instrText')); node.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve'); node.text = value
        else:
            node = etree.SubElement(r, qn('w:t')); node.text = value
        out.append(r)
    return out


REFERENCE_PREFIX = re.compile(r'^(?:[\[［【(（]\s*(?P<bracket>\d+)\s*[\]］】)）]|(?P<plain>\d+)\s*[.、)）])\s*')


def strip_reference_prefix(paragraph) -> str | None:
    visible = text_of(paragraph)
    match = REFERENCE_PREFIX.match(visible)
    if not match:
        return None
    nodes = paragraph.xpath('.//w:t', namespaces=NS)
    remaining = len(match.group(0))
    for t in nodes:
        value = t.text or ''
        if remaining <= 0:
            break
        take = min(remaining, len(value))
        t.text = value[take:]
        remaining -= take
    return match.group('bracket') or match.group('plain')


def add_reference_targets(paragraphs, ref_style_id, factory, report):
    mapping = {}
    seq = 0
    for p in paragraphs:
        if style_id(p) != ref_style_id or not text_of(p).strip():
            continue
        seq += 1
        typed = strip_reference_prefix(p)
        if typed and typed != str(seq):
            report['warnings'].append(f'reference typed number {typed} normalized to native list position {seq}')
        existing = p.xpath('./w:bookmarkStart[starts-with(@w:name, "_FmtRef")]/@w:name', namespaces=NS)
        if existing:
            bookmark = existing[0]
        else:
            start, end, bookmark = factory.pair(f'_FmtRef{seq:04d}')
            ppr = p.find(qn('w:pPr'))
            p.insert(1 if ppr is not None else 0, start)
            p.append(end)
        mapping[str(seq)] = {'bookmark': bookmark, 'mode': 'list'}
        report['targets']['references'] += 1
    return mapping


def field_range(p):
    children = list(p); begins = []; ends = []
    for i, child in enumerate(children):
        if child.tag != qn('w:r'):
            continue
        vals = child.xpath('./w:fldChar/@w:fldCharType', namespaces=NS)
        if vals == ['begin']:
            begins.append(i)
        elif vals == ['end']:
            ends.append(i)
    return (min(begins), max(ends)) if begins and ends else None


def add_caption_targets(paragraphs, factory, report, profile):
    figure_label = str(get(profile, 'captions', 'figure', 'label', default='图'))
    table_label = str(get(profile, 'captions', 'table', 'label', default='表'))
    mapping = {'figure': {}, 'table': {}, 'equation': {}}
    for p in paragraphs:
        seq = next((f for f in instructions(p) if f.upper().startswith('SEQ ')), '')
        if not seq:
            continue
        parts = seq.split(); label = parts[1] if len(parts) > 1 else ''
        kind = 'figure' if label == figure_label else 'table' if label == table_label else 'equation' if label in {'式', 'Equation'} else None
        if not kind:
            continue
        visible = text_of(p).strip()
        match = re.search(r'(\d+(?:[-－.]\d+)?)', visible)
        span = field_range(p)
        if not match or span is None:
            report['unresolved'].append(f'{kind} caption target not recognized: {visible[:80]}')
            continue
        number = match.group(1).replace('－', '-').replace('.', '-')
        prefix = {'figure': 'Fig', 'table': 'Tbl', 'equation': 'Eq'}[kind]
        start, end, bookmark = factory.pair(f'_Fmt{prefix}{len(mapping[kind]) + 1:04d}')
        p.insert(span[0], start); p.insert(span[1] + 2, end)
        mapping[kind][number] = bookmark
        report['targets'][kind + 's'] += 1
    return mapping


def add_typed_equation_targets(paragraphs, names, factory, mapping, report, profile):
    equation_style = style(profile, 'equation', 'Equation').casefold()
    label = str(get(profile, 'formulas', 'label', default='式'))
    for paragraph in paragraphs:
        if any(f.upper().startswith(f'SEQ {label}'.upper()) for f in instructions(paragraph)):
            continue
        st = names.get(style_id(paragraph), '').casefold()
        has_math = bool(paragraph.xpath('.//m:oMath|.//m:oMathPara', namespaces=NS))
        if not has_math and st != equation_style and 'equation' not in st and '公式' not in st:
            continue
        visible = text_of(paragraph)
        match = re.search(r'(?:式\s*)?[（(]?\s*(\d+)(?:[-－.](\d+))?\s*[）)]?\s*$', visible)
        if not match:
            continue
        chapter, sequence = match.group(1), match.group(2) or match.group(1)
        number = chapter if match.group(2) is None else f'{chapter}-{sequence}'
        text_nodes = paragraph.xpath('.//w:t', namespaces=NS)
        if not text_nodes:
            report['unresolved'].append(f'equation {number} number is not editable text')
            continue
        last = text_nodes[-1]
        last.text = re.sub(r'(?:式\s*)?[（(]?\s*\d+(?:[-－.]\d+)?\s*[）)]?\s*$', '', last.text or '')
        rpr = last.getparent().find(qn('w:rPr')) if last.getparent() is not None else None
        start, end, bookmark = factory.pair(f'_FmtEq{len(mapping["equation"]) + 1:04d}')
        paragraph.append(text_run('（', rpr)); paragraph.append(start)
        if match.group(2) is not None:
            chapter_level = int(get(profile, 'formulas', 'chapter_heading_level', default=1))
            for node in field_runs(f' STYLEREF {chapter_level} \\n ', chapter, rpr):
                paragraph.append(node)
            separator = str(get(profile, 'formulas', 'separator', default='-'))
            paragraph.append(text_run(separator, rpr))
            instruction = f' SEQ {label} \\* ARABIC \\s {chapter_level} '
        else:
            instruction = f' SEQ {label} \\* ARABIC '
        for node in field_runs(instruction, sequence, rpr):
            paragraph.append(node)
        paragraph.append(end); paragraph.append(text_run('）', rpr))
        mapping['equation'][number] = bookmark
        report['targets']['equations'] += 1


def paragraph_rpr(p):
    rprs = p.xpath('.//w:r[w:t]/w:rPr', namespaces=NS)
    signatures = {etree.tostring(x) for x in rprs}
    return rprs[0] if len(signatures) <= 1 and rprs else None


def paragraph_safe_to_rewrite(p):
    if instructions(p):
        return False
    forbidden = [
        './/w:hyperlink', './/w:drawing', './/w:pict', './/m:oMath', './/m:oMathPara',
        './/w:bookmarkStart', './/w:bookmarkEnd', './/w:commentRangeStart', './/w:commentRangeEnd',
        './/w:commentReference', './/w:footnoteReference', './/w:endnoteReference', './/w:sdt', './/w:tab', './/w:br'
    ]
    return not any(p.xpath(expr, namespaces=NS) for expr in forbidden)


def append_ref(p, bookmark, cached, rpr, paragraph_number=False):
    switch = ' \\n' if paragraph_number else ''
    for node in field_runs(f' REF {bookmark}{switch} \\h ', cached, rpr):
        p.append(node)


def build_token(profile):
    labels = [
        str(get(profile, 'captions', 'figure', 'label', default='图')),
        str(get(profile, 'captions', 'table', 'label', default='表')),
    ]
    formula_labels = get(profile, 'formulas', 'reference_labels', default=['式', '公式']) or ['式', '公式']
    if isinstance(formula_labels, str):
        formula_labels = [formula_labels]
    labels += [str(x) for x in formula_labels]
    labels += ['Figure', 'Table']
    alt = '|'.join(sorted({re.escape(x) for x in labels}, key=len, reverse=True))
    return re.compile(
        r'\[(?P<cites>\d+(?:\s*[,，;；、\-–—]\s*\d+)*)\]'
        r'|(?P<label>' + alt + r')\s*(?P<number>\d+(?:[-－.]\d+)?)',
        re.I,
    )


def repair_body_references(paragraphs, names, ref_map, target_map, report, profile):
    excluded = ('toc', 'caption', 'reference', '目录', '题注', '参考')
    token = build_token(profile)
    fig_label = str(get(profile, 'captions', 'figure', 'label', default='图')).casefold()
    tbl_label = str(get(profile, 'captions', 'table', 'label', default='表')).casefold()
    for p in paragraphs:
        st = names.get(style_id(p), '').casefold()
        if any(x in st for x in excluded):
            continue
        visible = text_of(p)
        existing_instructions = instructions(p)
        if any(x.upper().startswith('REF ') for x in existing_instructions):
            continue
        matches = list(token.finditer(visible))
        if not matches:
            continue
        if not paragraph_safe_to_rewrite(p):
            report['ambiguous'].append(f'complex paragraph skipped: {visible[:100]}')
            continue
        rpr = paragraph_rpr(p)
        if rpr is None and len(p.xpath('.//w:r[w:t]', namespaces=NS)) > 1:
            report['ambiguous'].append(f'mixed formatting skipped: {visible[:100]}')
            continue
        ppr = p.find(qn('w:pPr'))
        for child in list(p):
            if child is not ppr:
                p.remove(child)
        cursor = 0
        for match in matches:
            if match.start() > cursor:
                p.append(text_run(visible[cursor:match.start()], rpr))
            original = match.group(0)
            if match.group('cites'):
                nums = re.findall(r'\d+', match.group('cites'))
                if any(n not in ref_map for n in nums):
                    p.append(text_run(original, rpr)); report['unresolved'].append(f'citation target missing: {original}')
                else:
                    p.append(text_run('[', rpr)); inner = match.group('cites'); inner_cur = 0
                    for nm in re.finditer(r'\d+', inner):
                        if nm.start() > inner_cur:
                            p.append(text_run(inner[inner_cur:nm.start()], rpr))
                        n = nm.group(0)
                        append_ref(p, ref_map[n]['bookmark'], n, rpr, paragraph_number=True)
                        report['converted']['citations'] += 1
                        inner_cur = nm.end()
                    if inner_cur < len(inner):
                        p.append(text_run(inner[inner_cur:], rpr))
                    p.append(text_run(']', rpr))
            else:
                label = match.group('label')
                number = match.group('number').replace('－', '-').replace('.', '-')
                kind = 'figure' if label.casefold() in {fig_label, 'figure'} else 'table' if label.casefold() in {tbl_label, 'table'} else 'equation'
                bookmark = target_map[kind].get(number)
                if not bookmark:
                    p.append(text_run(original, rpr)); report['unresolved'].append(f'{kind} target missing: {original}')
                else:
                    prefix = original[:original.find(match.group('number'))]
                    p.append(text_run(prefix, rpr)); append_ref(p, bookmark, match.group('number'), rpr)
                    report['converted'][kind + 's'] += 1
            cursor = match.end()
        if cursor < len(visible):
            p.append(text_run(visible[cursor:], rpr))


def repair(
    input_path: Path,
    output_path: Path,
    strict: bool = False,
    profile_path: Path | None = None,
    template_path: Path | None = None,
) -> dict:
    profile = load_profile(profile_path, template_path)
    with zipfile.ZipFile(input_path) as zin:
        document = etree.fromstring(zin.read('word/document.xml'))
        styles = etree.fromstring(zin.read('word/styles.xml'))
        numbering = etree.fromstring(zin.read('word/numbering.xml'))
        files = {info.filename: zin.read(info.filename) for info in zin.infolist()}
    names, by_name = style_maps(styles)
    paragraphs = document.xpath('//w:body//w:p', namespaces=NS)
    report = {
        'targets': {'references': 0, 'figures': 0, 'tables': 0, 'equations': 0},
        'converted': {'citations': 0, 'figures': 0, 'tables': 0, 'equations': 0},
        'unresolved': [], 'ambiguous': [], 'warnings': [],
    }
    factory = BookmarkFactory(document)
    native_reference_numbering = bool(get(profile, 'references', 'numbering', 'native_list_required', default=True))
    ref_map = {}
    if native_reference_numbering:
        ref_style_id, _ = ensure_reference_numbering(styles, numbering, profile)
        ref_map = add_reference_targets(paragraphs, ref_style_id, factory, report)
    else:
        report['warnings'].append('native bibliography numbering disabled by active template/profile')
    target_map = add_caption_targets(paragraphs, factory, report, profile)
    add_typed_equation_targets(paragraphs, names, factory, target_map, report, profile)
    repair_body_references(paragraphs, names, ref_map, target_map, report, profile)
    files['word/document.xml'] = etree.tostring(document, xml_declaration=True, encoding='UTF-8', standalone='yes')
    files['word/styles.xml'] = etree.tostring(styles, xml_declaration=True, encoding='UTF-8', standalone='yes')
    files['word/numbering.xml'] = etree.tostring(numbering, xml_declaration=True, encoding='UTF-8', standalone='yes')
    target = output_path
    if input_path.resolve() == output_path.resolve():
        fd, temp = tempfile.mkstemp(prefix='crossrefs-', suffix='.docx', dir=output_path.parent)
        os.close(fd); target = Path(temp)
    try:
        with zipfile.ZipFile(target, 'w', zipfile.ZIP_DEFLATED) as zout:
            for name, data in files.items():
                zout.writestr(name, data)
        if target != output_path:
            target.replace(output_path)
    finally:
        if target != output_path and target.exists():
            target.unlink()
    if strict and (report['unresolved'] or report['ambiguous']):
        raise RuntimeError(json.dumps(report, ensure_ascii=False))
    return report


def main():
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument('input', type=Path); ap.add_argument('output', type=Path)
    ap.add_argument('--profile', type=Path); ap.add_argument('--template', type=Path); ap.add_argument('--strict', action='store_true')
    args = ap.parse_args()
    print(json.dumps(repair(args.input, args.output, args.strict, args.profile, args.template), ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
