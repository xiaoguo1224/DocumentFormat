#!/usr/bin/env python3
from __future__ import annotations

from copy import deepcopy
from pathlib import Path
from typing import Any
import re
import zipfile

from lxml import etree
import yaml

DEFAULT_PROFILE = Path(__file__).resolve().parents[1] / 'templates/generic-formal/profile.yaml'
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS = {'w': W}


def get(profile: dict[str, Any], *keys: str, default=None):
    cur: Any = profile
    for key in keys:
        if not isinstance(cur, dict) or key not in cur:
            return default
        cur = cur[key]
    return cur


def style(profile: dict[str, Any], role: str, fallback: str) -> str:
    return str(get(profile, 'styles', role, default=fallback) or fallback)


def normalized_titles(profile: dict[str, Any], section: str, defaults: list[str]) -> set[str]:
    values = get(profile, 'sections', 'titles', section, default=None)
    if values is None:
        values = defaults
    if isinstance(values, str):
        values = [values]
    return {''.join(str(v).split()).casefold() for v in values}


def _template_style_records(template: Path) -> list[dict[str, Any]]:
    with zipfile.ZipFile(template) as zf:
        root = etree.fromstring(zf.read('word/styles.xml'))
        document = etree.fromstring(zf.read('word/document.xml')) if 'word/document.xml' in zf.namelist() else None
    usage = {}
    if document is not None:
        for sid in document.xpath('//w:pPr/w:pStyle/@w:val', namespaces=NS):
            usage[sid] = usage.get(sid, 0) + 1
    records = []
    for st in root.xpath('./w:style[@w:type="paragraph"]', namespaces=NS):
        sid = st.get(f'{{{W}}}styleId', '')
        names = st.xpath('./w:name/@w:val', namespaces=NS)
        name = names[0] if names else sid
        outline = st.xpath('./w:pPr/w:outlineLvl/@w:val', namespaces=NS)
        numpr = bool(st.xpath('./w:pPr/w:numPr', namespaces=NS))
        page_break = bool(st.xpath('./w:pPr/w:pageBreakBefore', namespaces=NS))
        records.append({'style_id': sid, 'name': name, 'outline': int(outline[0]) if outline and str(outline[0]).isdigit() else None, 'usage': usage.get(sid, 0), 'numbered': numpr, 'page_break_before': page_break})
    return records


def _pick_by_alias(records: list[dict[str, Any]], aliases: list[str]) -> str | None:
    aliases_cf = [a.casefold() for a in aliases]
    exact = []
    for priority, alias in enumerate(aliases_cf):
        for record in records:
            if record['name'].casefold() == alias or record['style_id'].casefold() == alias:
                exact.append((record.get('usage', 0), -priority, record['name']))
    if exact:
        exact.sort(reverse=True)
        return exact[0][2]
    fuzzy = []
    for priority, alias in enumerate(aliases_cf):
        for record in records:
            hay = f"{record['name']} {record['style_id']}".casefold()
            if alias in hay:
                fuzzy.append((record.get('usage', 0), -priority, record['name']))
    if fuzzy:
        fuzzy.sort(reverse=True)
        return fuzzy[0][2]
    return None




def _field_in_root(root, field: str) -> bool:
    if root is None:
        return False
    needle = field.upper()
    return any(needle in (text or '').upper().split() for text in root.xpath('.//w:instrText/text()', namespaces=NS))


def _template_native_hints(template: Path, reference_style_name: str | None = None) -> dict[str, Any]:
    hints: dict[str, Any] = {}
    with zipfile.ZipFile(template) as zf:
        document = etree.fromstring(zf.read('word/document.xml')) if 'word/document.xml' in zf.namelist() else None
        settings = etree.fromstring(zf.read('word/settings.xml')) if 'word/settings.xml' in zf.namelist() else None
        styles = etree.fromstring(zf.read('word/styles.xml')) if 'word/styles.xml' in zf.namelist() else None
        numbering = etree.fromstring(zf.read('word/numbering.xml')) if 'word/numbering.xml' in zf.namelist() else None

        toc = _field_in_root(document, 'TOC')
        page_location = None
        for name in zf.namelist():
            if not ((name.startswith('word/header') or name.startswith('word/footer')) and name.endswith('.xml')):
                continue
            root = etree.fromstring(zf.read(name))
            if _field_in_root(root, 'PAGE'):
                top = name.startswith('word/header')
                para = next((p for p in root.xpath('.//w:p', namespaces=NS) if _field_in_root(p, 'PAGE')), None)
                jc = para.xpath('./w:pPr/w:jc/@w:val', namespaces=NS)[0] if para is not None and para.xpath('./w:pPr/w:jc/@w:val', namespaces=NS) else 'center'
                align = {'left': 'left', 'right': 'right', 'center': 'center'}.get(jc, 'center')
                page_location = f"{'top' if top else 'bottom'}_{align}"
                break
        hints['toc_field'] = toc
        hints['page_field'] = page_location is not None
        hints['page_position'] = page_location

        section_formats = []
        if document is not None:
            for sect in document.xpath('//w:sectPr', namespaces=NS):
                vals = sect.xpath('./w:pgNumType/@w:fmt', namespaces=NS)
                if vals:
                    section_formats.append(vals[0])
        hints['section_formats'] = section_formats

        labels = []
        if settings is not None:
            labels = settings.xpath('./w:captions/w:caption/@w:name', namespaces=NS)
        hints['caption_labels'] = labels

        if reference_style_name and styles is not None:
            style_nodes = styles.xpath('./w:style[w:name/@w:val=$n]', namespaces=NS, n=reference_style_name)
            if style_nodes:
                ref_sid = style_nodes[0].get(f'{{{W}}}styleId')
                nums = style_nodes[0].xpath('./w:pPr/w:numPr/w:numId/@w:val', namespaces=NS)
                if nums and numbering is not None:
                    num_id = nums[0]
                    abs_ids = numbering.xpath('./w:num[@w:numId=$n]/w:abstractNumId/@w:val', namespaces=NS, n=num_id)
                    if abs_ids:
                        levels = numbering.xpath('./w:abstractNum[@w:abstractNumId=$n]/w:lvl[@w:ilvl="0"]', namespaces=NS, n=abs_ids[0])
                        if levels:
                            text = levels[0].xpath('./w:lvlText/@w:val', namespaces=NS)
                            suffix = levels[0].xpath('./w:suff/@w:val', namespaces=NS)
                            hints['reference_numbering'] = {
                                'required': True,
                                'format': text[0] if text else None,
                                'suffix': suffix[0] if suffix else None,
                            }
                if 'reference_numbering' not in hints and document is not None and ref_sid:
                    samples = []
                    for paragraph in document.xpath('//w:p[w:pPr/w:pStyle/@w:val=$sid]', namespaces=NS, sid=ref_sid):
                        samples.append(''.join(paragraph.xpath('.//w:t/text()', namespaces=NS)).strip())
                    if any(re.match(r'^\s*[\[［【(（]\s*1\s*[\]］】)）]', text) for text in samples):
                        hints['reference_numbering'] = {'required': True, 'format': '[%1]', 'suffix': 'tab'}
    return hints

def infer_profile_from_template(profile: dict[str, Any], template: Path) -> dict[str, Any]:
    """Conservatively infer semantic style mappings from a custom Word template.

    The Word file remains the visual source of truth. This only fills semantic style
    names when the user supplies a template without a companion profile.
    """
    result = deepcopy(profile)
    styles = result.setdefault('styles', {})
    records = _template_style_records(template)
    max_level = int(get(result, 'headings', 'max_level', default=4))

    for level in range(1, max_level + 1):
        candidates = [r for r in records if r['outline'] == level - 1]
        if candidates:
            # Prefer a conventional heading-like name, otherwise the first outline style.
            candidates.sort(key=lambda r: (-r.get('usage', 0), -int(r.get('numbered', False)), 0 if re.search(r'(标题|章|节)', r['name'], re.I) else 1 if re.search(r'heading', r['name'], re.I) else 2))
            styles[f'heading_{level}'] = candidates[0]['name']

    aliases = {
        'body': ['Body Text Generic', 'Body Text', '正文样式', '正文', '正文文本', 'Normal'],
        'references': ['References', 'Reference', '参考文献', '文献条目'],
        'figure_caption': ['Figure Caption', '图题', '图注', 'Caption'],
        'table_caption': ['Table Caption', '表题', '表注', 'Caption'],
        'toc_title': ['TOC Title', '目录标题'],
        'appendix_heading': ['Appendix Heading', '附录标题'],
        'appendix_body': ['Appendix Body', '附录正文'],
        'equation': ['Equation', '公式'],
        'special_heading_1': ['Special Heading 1', '特殊标题', '参考文献标题'],
    }
    for role, names in aliases.items():
        picked = _pick_by_alias(records, names)
        if picked:
            styles[role] = picked
    for i in range(1, 4):
        picked = _pick_by_alias(records, [f'TOC {i}', f'目录 {i}', f'目录{i}'])
        if picked:
            styles[f'toc_{i}'] = picked

    # Infer behavior from the custom template itself so generic defaults do not
    # override a template that intentionally omits or places features differently.
    h1_name = styles.get('heading_1')
    h1_record = next((r for r in records if r['name'] == h1_name), None)
    if h1_record is not None:
        headings = result.setdefault('headings', {})
        headings['use_page_break_before'] = bool(h1_record.get('page_break_before'))
        headings['chapter_starts_new_page'] = bool(h1_record.get('page_break_before'))

    hints = _template_native_hints(template, styles.get('references'))
    pagination = result.setdefault('pagination', {})
    pagination['page_field_required'] = bool(hints.get('page_field'))
    if hints.get('page_position'):
        pagination['position'] = hints['page_position']
    formats = hints.get('section_formats') or []
    if formats:
        if len(set(formats)) > 1:
            pagination.setdefault('front_matter', {})['format'] = formats[0]
            pagination.setdefault('body', {})['format'] = formats[-1]
        else:
            pagination.setdefault('body', {})['format'] = formats[-1]

    contents = result.setdefault('contents', {})
    contents['enabled'] = bool(hints.get('toc_field'))
    contents['native_toc_field_required'] = bool(hints.get('toc_field'))

    labels = hints.get('caption_labels') or []
    captions = result.setdefault('captions', {})
    if labels:
        # Preserve existing role labels when registered; otherwise infer common English labels.
        if '图' not in labels and 'Figure' in labels:
            captions.setdefault('figure', {})['label'] = 'Figure'
        if '表' not in labels and 'Table' in labels:
            captions.setdefault('table', {})['label'] = 'Table'

    ref_hint = hints.get('reference_numbering')
    if ref_hint:
        refs = result.setdefault('references', {})
        numbering_cfg = refs.setdefault('numbering', {})
        numbering_cfg['native_list_required'] = True
        if ref_hint.get('format'):
            numbering_cfg['format'] = ref_hint['format']
        if ref_hint.get('suffix'):
            numbering_cfg['suffix'] = ref_hint['suffix']
    elif styles.get('references'):
        result.setdefault('references', {}).setdefault('numbering', {})['native_list_required'] = False

    result['_profile_inferred_from_template'] = str(template)
    return result


def load_profile(path: Path | None = None, template: Path | None = None) -> dict[str, Any]:
    target = path or DEFAULT_PROFILE
    with target.open('r', encoding='utf-8') as f:
        data = yaml.safe_load(f) or {}
    data.setdefault('_profile_path', str(target))
    if path is None and template is not None:
        try:
            if template.resolve() != (DEFAULT_PROFILE.parent / 'template.docx').resolve():
                data = infer_profile_from_template(data, template)
        except (OSError, KeyError, zipfile.BadZipFile, etree.XMLSyntaxError):
            # A malformed template should be rejected by the formatter/validator later;
            # semantic inference is best-effort and must not mask the underlying error.
            pass
    return data
