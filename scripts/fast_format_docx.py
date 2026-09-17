#!/usr/bin/env python3
"""Template/profile-driven conservative DOCX formatter."""
from __future__ import annotations

import argparse
import json
import re
import sys
import tempfile
import zipfile
from copy import deepcopy
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from lxml import etree

from profile_config import load_profile, style, get, normalized_titles
from repair_cross_references import repair
from template_runtime import apply_table_rules, apply_template_sections, ensure_page_fields

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
DEFAULT_TEMPLATE = Path(__file__).resolve().parents[1] / 'templates/generic-formal/template.docx'
DEFAULT_PROFILE = Path(__file__).resolve().parents[1] / 'templates/generic-formal/profile.yaml'


def wt(name):
    return f'{{{W}}}{name}'


def merge_settings(source: bytes, template: bytes) -> bytes:
    root = etree.fromstring(source)
    donor = etree.fromstring(template)
    for name in ('captions', 'updateFields'):
        old = root.find(wt(name))
        if old is not None:
            root.remove(old)
        new = donor.find(wt(name))
        if new is not None:
            root.append(deepcopy(new))
    update = root.find(wt('updateFields'))
    if update is None:
        update = etree.SubElement(root, wt('updateFields'))
    update.set(wt('val'), 'true')
    return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone='yes')


def apply_template_parts(template: Path, source: Path, output: Path) -> None:
    parts = {
        'word/styles.xml': 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml',
        'word/numbering.xml': 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml',
        'word/theme/theme1.xml': 'application/vnd.openxmlformats-officedocument.theme+xml',
        'word/fontTable.xml': 'application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml',
    }
    with zipfile.ZipFile(template) as zt, zipfile.ZipFile(source) as zs:
        repl = {name: zt.read(name) for name in parts if name in zt.namelist()}
        if 'word/settings.xml' in zt.namelist() and 'word/settings.xml' in zs.namelist():
            repl['word/settings.xml'] = merge_settings(zs.read('word/settings.xml'), zt.read('word/settings.xml'))
        types = etree.fromstring(zs.read('[Content_Types].xml'))
        for name, ctype in parts.items():
            if name not in repl:
                continue
            part = '/' + name
            node = next((n for n in types.findall(f'{{{CT}}}Override') if n.get('PartName') == part), None)
            if node is None:
                node = etree.SubElement(types, f'{{{CT}}}Override')
                node.set('PartName', part)
            node.set('ContentType', ctype)
        repl['[Content_Types].xml'] = etree.tostring(types, xml_declaration=True, encoding='UTF-8', standalone='yes')
        with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as zout:
            seen = set()
            for info in zs.infolist():
                seen.add(info.filename)
                zout.writestr(info, repl.get(info.filename, zs.read(info.filename)))
            for name, data in repl.items():
                if name not in seen:
                    zout.writestr(name, data)


def style_name(p):
    try:
        return p.style.name or ''
    except (KeyError, ValueError):
        return ''


def paragraph_outline_level(p) -> int | None:
    """Return effective Word outline level, following basedOn style ancestry."""
    direct = p._p.pPr.find(qn('w:outlineLvl')) if p._p.pPr is not None else None
    if direct is not None:
        value = direct.get(qn('w:val'))
        return int(value) if value and value.isdigit() else None
    try:
        current = p.style
    except (KeyError, ValueError):
        return None
    seen = set()
    while current is not None and current.style_id not in seen:
        seen.add(current.style_id)
        ppr = current._element.pPr
        node = ppr.find(qn('w:outlineLvl')) if ppr is not None else None
        if node is not None:
            value = node.get(qn('w:val'))
            return int(value) if value and value.isdigit() else None
        current = current.base_style
    return None


class StyleResolver:
    def __init__(self, doc: Document, template: Path):
        self.by_id = {s.style_id: s for s in doc.styles}
        self.by_doc_name = {(s.name or '').casefold(): s for s in doc.styles}
        self.raw_name_to_id = {}
        with zipfile.ZipFile(template) as zf:
            root = etree.fromstring(zf.read('word/styles.xml'))
        for st in root.findall(wt('style')):
            sid = st.get(wt('styleId'))
            name = st.find(wt('name'))
            raw = name.get(wt('val')) if name is not None else None
            if sid and raw:
                self.raw_name_to_id[raw.casefold()] = sid

    def resolve(self, requested: str):
        key = requested.casefold()
        sid = self.raw_name_to_id.get(key)
        if sid and sid in self.by_id:
            return self.by_id[sid]
        if requested in self.by_id:
            return self.by_id[requested]
        return self.by_doc_name.get(key)


def norm(text):
    return ''.join(text.split()).casefold()


def classify(p, index, state, profile):
    text = p.text.strip()
    compact = norm(text)
    name = style_name(p).casefold()
    styles = profile.get('styles', {})
    refs = normalized_titles(profile, 'references', ['参考文献', 'References'])
    acks = normalized_titles(profile, 'acknowledgements', ['致谢', '致  谢', 'Acknowledgements'])
    appendices = normalized_titles(profile, 'appendices', ['附录', 'Appendix'])
    toc_titles = normalized_titles(profile, 'toc', ['目录', 'Contents'])
    cn_abs = normalized_titles(profile, 'chinese_abstract', ['摘要', '中文摘要', '摘  要'])
    en_abs = normalized_titles(profile, 'english_abstract', ['Abstract'])

    if not text:
        role = 'appendix_body' if state == 'appendix' else 'body'
        return style(profile, role, 'Appendix Body' if state == 'appendix' else 'Body Text Generic'), None, None, state
    if compact in refs:
        return style(profile, 'special_heading_1', 'Special Heading 1'), None, None, 'references'
    if compact in acks:
        return style(profile, 'special_heading_1', 'Special Heading 1'), None, None, 'acknowledgements'
    is_appendix_heading = (
        compact in appendices
        or bool(re.match(r'^附录(?:[A-Z一二三四五六七八九十0-9]+)(?:[^正文].*)?$', compact, re.I))
        or bool(re.match(r'^appendix(?:[A-Z0-9]+)(?:\s.*)?$', compact, re.I))
    )
    if is_appendix_heading:
        return style(profile, 'appendix_heading', 'Appendix Heading'), None, None, 'appendix'
    if compact in toc_titles:
        return style(profile, 'toc_title', 'TOC Title'), None, None, 'toc'
    if compact in cn_abs:
        return style(profile, 'chinese_abstract_title', 'Abstract Title CN'), None, None, 'abstract_cn'
    if compact in en_abs:
        return style(profile, 'english_abstract_title', 'Abstract Title EN'), None, None, 'abstract_en'

    if state == 'references':
        return style(profile, 'references', 'References'), None, None, state
    if state == 'appendix' and paragraph_outline_level(p) is None:
        return style(profile, 'appendix_body', 'Appendix Body'), None, None, state
    if state == 'abstract_cn':
        return style(profile, 'chinese_abstract_body', 'Abstract Body CN'), None, None, state
    if state == 'abstract_en':
        return style(profile, 'english_abstract_body', 'Abstract Body EN'), None, None, state

    if 'equation' in name or '公式' in name or p._p.xpath('.//m:oMath | .//m:oMathPara'):
        return style(profile, 'equation', 'Equation'), None, None, state

    labels = {
        str(get(profile, 'captions', 'figure', 'label', default='图')): 'figure_caption',
        str(get(profile, 'captions', 'table', 'label', default='表')): 'table_caption',
    }
    for label, role in labels.items():
        match = re.match(rf'^{re.escape(label)}\s*(\d+)(?:[-－.]([0-9]+))?\s*(.+)$', text, re.I)
        if match:
            chapter, sequence, title = match.groups()
            fallback = 'Figure Caption' if role == 'figure_caption' else 'Table Caption'
            return style(profile, role, fallback), None, {
                'label': label,
                'chapter': chapter,
                'sequence': sequence or '1',
                'title': title.strip(),
            }, state

    toc = re.search(r'toc\s*([1-9])', name)
    if toc:
        return style(profile, f'toc_{toc.group(1)}', f'TOC {toc.group(1)}'), None, None, state

    max_level = int(get(profile, 'headings', 'max_level', default=4))
    outline = paragraph_outline_level(p)
    if outline is not None and 0 <= outline < max_level:
        return style(profile, f'heading_{outline + 1}', f'Heading {outline + 1}'), None, None, 'body'

    heading = re.search(r'(?:heading|标题)\s*([1-9])', name)
    if heading:
        level = min(int(heading.group(1)), max_level)
        return style(profile, f'heading_{level}', f'Heading {level}'), None, None, 'body'
    if 'title' in name and 'subtitle' not in name:
        return styles.get('title', 'Title'), None, None, state
    if 'subtitle' in name:
        return styles.get('subtitle', 'Subtitle'), None, None, state
    if 'list bullet' in name:
        return styles.get('list_bullet', 'List Bullet'), None, None, state
    if 'list number' in name:
        return styles.get('list_number', 'List Number'), None, None, state

    typed = re.match(r'^(\d+(?:[.．]\d+){0,8})[ \t、.．]+(.+)$', text)
    if typed and len(text) <= 120 and not text.endswith(('。', '；', ';')):
        level = min(typed.group(1).replace('．', '.').count('.') + 1, max_level)
        return style(profile, f'heading_{level}', f'Heading {level}'), typed.group(2).strip(), None, 'body'
    if re.match(r'^[•·●○▪-]\s*', text):
        return styles.get('list_bullet', 'List Bullet'), re.sub(r'^[•·●○▪-]\s*', '', text), None, state
    numbered = re.match(r'^\d+[.)、]\s*(.+)$', text)
    if numbered:
        return styles.get('list_number', 'List Number'), numbered.group(1).strip(), None, state
    if index < 12:
        sizes = [r.font.size.pt for r in p.runs if r.font.size]
        if sizes and max(sizes) >= 18 and len(text) <= 80:
            return styles.get('title', 'Title'), None, None, state
    return style(profile, 'body', 'Body Text Generic'), None, None, state


def clear_direct_format(p):
    ppr = p._p.get_or_add_pPr()
    for name in ('jc', 'spacing', 'ind', 'tabs', 'pBdr', 'shd', 'numPr', 'outlineLvl'):
        node = ppr.find(qn(f'w:{name}'))
        if node is not None:
            ppr.remove(node)
    for run in p._p.xpath('.//w:r'):
        rpr = run.find(qn('w:rPr'))
        if rpr is None:
            continue
        for name in ('rFonts', 'sz', 'szCs', 'color', 'highlight', 'spacing', 'kern', 'position'):
            node = rpr.find(qn(f'w:{name}'))
            if node is not None:
                rpr.remove(node)


def add_field(p, instruction, cached):
    for kind, value in (('begin', None), (None, instruction), ('separate', None), (None, cached), ('end', None)):
        run = OxmlElement('w:r')
        if kind:
            node = OxmlElement('w:fldChar')
            node.set(qn('w:fldCharType'), kind)
        elif value == instruction:
            node = OxmlElement('w:instrText')
            node.set(qn('xml:space'), 'preserve')
            node.text = value
        else:
            node = OxmlElement('w:t')
            node.text = value
        run.append(node)
        p._p.append(run)


def replace_caption(p, data, profile):
    for child in list(p._p):
        if child.tag != qn('w:pPr'):
            p._p.remove(child)
    p.add_run(data['label'])
    figure_label = get(profile, 'captions', 'figure', 'label', default='图')
    kind = 'figure' if data['label'] == figure_label else 'table'
    chapter_level = int(get(profile, 'captions', kind, 'chapter_heading_level', default=1))
    separator = str(get(profile, 'captions', kind, 'separator', default='-'))
    add_field(p, f' STYLEREF {chapter_level} \\n ', data['chapter'])
    p.add_run(separator)
    add_field(p, f" SEQ {data['label']} \\* ARABIC \\s {chapter_level} ", data['sequence'])
    p.add_run(' ' + data['title'])
    # Alignment comes from the active caption style/template, never hard-coded here.


def has_field(root, field):
    return any(field in (n.text or '').upper().split() for n in root.xpath('.//w:instrText'))


def ensure_toc(doc, profile):
    if has_field(doc.element, 'TOC'):
        return False
    toc_style = style(profile, 'toc_title', 'TOC Title').casefold()
    toc_texts = normalized_titles(profile, 'toc', ['目录', 'Contents'])
    title = next((p for p in doc.paragraphs if style_name(p).casefold() == toc_style or norm(p.text) in toc_texts), None)
    if title is None:
        return False
    new = OxmlElement('w:p')
    title._p.addnext(new)
    p = Paragraph(new, title._parent)
    p.style = style(profile, 'toc_1', 'TOC 1')
    depth = int(get(profile, 'contents', 'depth', default=3))
    add_field(p, f' TOC \\o "1-{depth}" \\h \\z \\u ', '打开文档后更新目录')
    return True


def format_docx(source, template, output, add_toc, profile_path=None):
    profile = load_profile(profile_path, template)
    if source.resolve() == output.resolve():
        raise ValueError('output must differ from source')
    src = Document(source)
    plans = []
    state = 'body'
    for i, p in enumerate(src.paragraphs):
        role, repl, caption, state = classify(p, i, state, profile)
        plans.append((role, repl, caption, state))

    output.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix='docx-format-') as temp:
        merged = Path(temp) / 'merged.docx'
        apply_template_parts(template, source, merged)
        doc = Document(merged)
        if len(doc.paragraphs) != len(plans):
            raise RuntimeError('paragraph count changed while applying template')
        resolver = StyleResolver(doc, template)
        stats = {'paragraphs': len(plans), 'headings': 0, 'captions': 0, 'lists': 0, 'fallback_styles': 0}
        heading_roles = {style(profile, f'heading_{i}', f'Heading {i}') for i in range(1, int(get(profile, 'headings', 'max_level', default=4)) + 1)}
        list_roles = {profile.get('styles', {}).get('list_bullet', 'List Bullet'), profile.get('styles', {}).get('list_number', 'List Number')}

        for p, (role, repl, caption, _state) in zip(doc.paragraphs, plans):
            target = resolver.resolve(role)
            if target is None:
                target = resolver.resolve('Normal')
                stats['fallback_styles'] += 1
            p.style = target
            clear_direct_format(p)
            if repl is not None:
                p.text = repl
            if caption:
                replace_caption(p, caption, profile)
                stats['captions'] += 1
            if role in heading_roles:
                stats['headings'] += 1
            if role in list_roles:
                stats['lists'] += 1

        stats['table_rules'] = apply_table_rules(doc, profile, resolver)
        stats.update(apply_template_sections(doc, template, profile))
        stats['page_fields_added'] = ensure_page_fields(doc, profile)
        stats['toc_added'] = ensure_toc(doc, profile) if add_toc else False
        doc.save(output)

    with zipfile.ZipFile(output) as zf:
        if zf.testzip() is not None:
            raise RuntimeError('generated DOCX failed ZIP integrity check')
    return stats


def main():
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument('--source', type=Path, required=True)
    ap.add_argument('--output', type=Path, required=True)
    ap.add_argument('--template', type=Path, default=DEFAULT_TEMPLATE)
    ap.add_argument('--profile', type=Path)
    ap.add_argument('--ensure-toc', action='store_true')
    ap.add_argument('--skip-cross-references', action='store_true')
    ap.add_argument('--strict-cross-references', action='store_true')
    args = ap.parse_args()
    for path in (args.source, args.template):
        if not path.is_file():
            ap.error(f'file not found: {path}')
    if args.profile is not None and not args.profile.is_file():
        ap.error(f'file not found: {args.profile}')
    result = format_docx(args.source, args.template, args.output, args.ensure_toc, args.profile)
    if not args.skip_cross_references:
        result['cross_references'] = repair(
            args.output,
            args.output,
            args.strict_cross_references,
            args.profile,
            args.template,
        )
    print(json.dumps(result, ensure_ascii=False))


if __name__ == '__main__':
    main()
