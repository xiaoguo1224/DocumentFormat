#!/usr/bin/env python3
"""Profile-driven structural validator for Word-native DOCX features."""
from __future__ import annotations
import argparse,json,re,sys,zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
from profile_config import load_profile, style, get

W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

def read_xml(zf,name):
    try:return ET.fromstring(zf.read(name))
    except KeyError:return None

def val(el,key='val'):return el.get(W+key) if el is not None else None

def onoff(el):
    if el is None:return False
    v=val(el);return v is None or str(v).lower() not in {'0','false','off','no'}

def paragraph_text(p):return ''.join(t.text or '' for t in p.iter(W+'t'))
def field_text(root):return ' '.join((x.text or '').strip() for x in root.iter(W+'instrText') if (x.text or '').strip())

def validate(path:Path,profile_path:Path|None=None,require_toc=False,require_captions=False,require_multilevel=False,require_crossrefs=False,require_reference_numbering=False):
    profile=load_profile(profile_path);checks=[]
    def add(name,status,detail):checks.append({'check':name,'status':status,'detail':detail})
    with zipfile.ZipFile(path) as zf:
        doc=read_xml(zf,'word/document.xml');styles=read_xml(zf,'word/styles.xml');numbering=read_xml(zf,'word/numbering.xml');settings=read_xml(zf,'word/settings.xml')
        add('styles.xml','PASS' if styles is not None else 'FAIL','styles present' if styles is not None else 'missing word/styles.xml')
        add('numbering.xml','PASS' if numbering is not None else ('FAIL' if require_multilevel or require_reference_numbering else 'WARN'),'numbering present' if numbering is not None else 'missing word/numbering.xml')
        info={};name_to_id={}
        if styles is not None:
            for st in styles.findall(W+'style'):
                sid=val(st,'styleId');nm=val(st.find(W+'name')) or sid;name_to_id[nm.casefold()]=sid
                ppr=st.find(W+'pPr');outline=ppr.find(W+'outlineLvl') if ppr is not None else None;num_id=ilvl=None
                if ppr is not None:
                    numpr=ppr.find(W+'numPr')
                    if numpr is not None:num_id=val(numpr.find(W+'numId'));ilvl=val(numpr.find(W+'ilvl'))
                info[sid]={'name':nm,'outline':val(outline),'page_break_before':onoff(ppr.find(W+'pageBreakBefore')) if ppr is not None else False,'num_id':num_id,'ilvl':ilvl}
        max_level=int(get(profile,'headings','max_level',default=4));heading_ids=[];bindings=[]
        for i in range(1,max_level+1):
            sid=name_to_id.get(style(profile,f'heading_{i}',f'Heading {i}').casefold());heading_ids.append(sid);bindings.append(info.get(sid))
        heading_ok=all(bindings)
        if heading_ok:
            ids=[x['num_id'] for x in bindings];levels=[x['ilvl'] for x in bindings]
            heading_ok=all(ids) and len(set(ids))==1 and levels==[str(i) for i in range(max_level)]
        add('heading multilevel binding','PASS' if heading_ok else ('FAIL' if require_multilevel else 'WARN'),f'{max_level} heading levels share native numbering' if heading_ok else 'profile heading styles are not fully bound to one native multilevel list')
        num_levels_ok=False
        if numbering is not None and heading_ok:
            num_id=bindings[0]['num_id'];abs_id=None
            for n in numbering.findall(W+'num'):
                if val(n,'numId')==num_id:abs_id=val(n.find(W+'abstractNumId'));break
            if abs_id is not None:
                for a in numbering.findall(W+'abstractNum'):
                    if val(a,'abstractNumId')==abs_id:num_levels_ok=len(a.findall(W+'lvl'))>=max_level;break
        add('multilevel numbering definition','PASS' if num_levels_ok else ('FAIL' if require_multilevel else 'WARN'),'required levels found' if num_levels_ok else 'required heading numbering definition not confirmed')

        fields=[];static_heading=[];static_cross=[];caption_candidates=[];ref_paras=[]
        ref_style_id=name_to_id.get(style(profile,'references','References').casefold())
        if doc is not None:
            fields.append(field_text(doc))
            for p in doc.iter(W+'p'):
                txt=paragraph_text(p).strip();ppr=p.find(W+'pPr');sid=None
                if ppr is not None:
                    ps=ppr.find(W+'pStyle');sid=val(ps)
                if sid in heading_ids and re.match(r'^\d+(?:\.\d+){0,8}\s+\S',txt):static_heading.append(txt[:80])
                if sid==ref_style_id and txt:ref_paras.append((p,txt))
                if re.match(r'^(图|表|Figure|Table)\s*\d',txt,re.I):caption_candidates.append((txt[:100],field_text(p)))
                style_name=info.get(sid,{}).get('name','').casefold();excluded=any(x in style_name for x in ('toc','caption','reference','目录','题注','参考'))
                para_fields=field_text(p).upper()
                if not excluded and ' REF ' not in f' {para_fields} ':
                    fig_label=str(get(profile,'captions','figure','label',default='图'))
                    tbl_label=str(get(profile,'captions','table','label',default='表'))
                    formula_labels=get(profile,'formulas','reference_labels',default=['式','公式']) or ['式','公式']
                    if isinstance(formula_labels,str): formula_labels=[formula_labels]
                    alt='|'.join(sorted({re.escape(x) for x in [fig_label,tbl_label,*map(str,formula_labels),'Figure','Table']},key=len,reverse=True))
                    pattern=r'\[(?:\d+[\s,，;；、\-–—]*)+\]|(?:'+alt+r')\s*\d+(?:[-－.]\d+)?'
                    for m in re.finditer(pattern,txt,re.I):static_cross.append(m.group(0))
        for name in zf.namelist():
            if (name.startswith('word/header') or name.startswith('word/footer')) and name.endswith('.xml'):
                r=read_xml(zf,name)
                if r is not None:fields.append(field_text(r))
        blob='\n'.join(fields).upper();toc_ok=bool(re.search(r'(^|\s)TOC(\s|$)',blob));seq_ok=bool(re.search(r'(^|\s)SEQ\s',blob));page_ok=bool(re.search(r'(^|\s)PAGE(\s|$)',blob))
        ref_targets=re.findall(r'(?:^|\s)REF\s+([A-Z_][A-Z0-9_.]*)',blob,re.I);ref_ok=bool(ref_targets)
        add('TOC field','PASS' if toc_ok else ('FAIL' if require_toc else 'WARN'),'TOC field found' if toc_ok else 'no TOC field detected')
        add('SEQ caption fields','PASS' if seq_ok else ('FAIL' if require_captions else 'WARN'),'SEQ field(s) found' if seq_ok else 'no SEQ fields detected')
        expected_labels={str(get(profile,'captions','figure','label',default='图')),str(get(profile,'captions','table','label',default='表'))}
        labels=set()
        if settings is not None:
            caps=settings.find(W+'captions')
            if caps is not None:
                for cap in caps.findall(W+'caption'):
                    name=val(cap,'name')
                    if name:labels.add(name)
        labels_ok=expected_labels.issubset(labels)
        add('registered caption labels','PASS' if labels_ok else ('FAIL' if require_captions else 'WARN'),f'expected={sorted(expected_labels)} found={sorted(labels)}')
        add('cross-reference fields','PASS' if ref_ok else ('FAIL' if require_crossrefs and static_cross else 'WARN'),f'{len(ref_targets)} REF field(s) found' if ref_ok else 'no REF fields detected')
        bookmarks=set()
        if doc is not None:bookmarks={val(n,'name').upper() for n in doc.iter(W+'bookmarkStart') if val(n,'name')}
        missing=sorted({t for t in ref_targets if t.upper() not in bookmarks})
        add('REF target integrity','FAIL' if missing else 'PASS',f'missing targets: {missing}' if missing else 'all REF targets resolve')
        add('remaining typed cross-references','FAIL' if require_crossrefs and static_cross else ('WARN' if static_cross else 'PASS'),f'{len(static_cross)} remain: {static_cross[:8]}' if static_cross else 'no recognizable typed cross-references remain')
        add('page-number fields','PASS' if page_ok else 'WARN','PAGE field found' if page_ok else 'no PAGE field detected')
        add('possible typed heading numbers','WARN' if static_heading else 'PASS',f'{len(static_heading)} possible typed headings' if static_heading else 'none')
        if caption_candidates:
            bad=[x for x in caption_candidates if 'SEQ ' not in (x[1] or '').upper()];add('possible static captions','WARN' if bad else 'PASS',f'{len(bad)} caption-looking paragraph(s) lack SEQ' if bad else 'caption-looking paragraphs contain SEQ')
        h1_sid=heading_ids[0] if heading_ids else None;need_pb=bool(get(profile,'headings','use_page_break_before',default=True));pb=info.get(h1_sid,{}).get('page_break_before',False)
        add('chapter pageBreakBefore','PASS' if (pb or not need_pb) else 'FAIL','configured correctly' if (pb or not need_pb) else 'required heading style lacks pageBreakBefore')

        # Bibliography native numbering contract: [%1] + TAB, and body REF fields use paragraph-number switch.
        ref_native=False;ref_detail='reference style not found'
        if ref_style_id and numbering is not None:
            num_id=info.get(ref_style_id,{}).get('num_id');ilvl=info.get(ref_style_id,{}).get('ilvl')
            if num_id:
                abs_id=None
                for n in numbering.findall(W+'num'):
                    if val(n,'numId')==num_id:abs_id=val(n.find(W+'abstractNumId'));break
                if abs_id:
                    for a in numbering.findall(W+'abstractNum'):
                        if val(a,'abstractNumId')==abs_id:
                            lvl=next((x for x in a.findall(W+'lvl') if val(x,'ilvl')==(ilvl or '0')),None)
                            if lvl is not None:
                                text=val(lvl.find(W+'lvlText'));suff=val(lvl.find(W+'suff'));expected=str(get(profile,'references','numbering','format',default='[%1]'));exp_suff=str(get(profile,'references','numbering','suffix',default='tab'))
                                ref_native=(text==expected and suff==exp_suff);ref_detail=f'lvlText={text!r}, suffix={suff!r}, numId={num_id}'
        add('reference native numbering','PASS' if ref_native else ('FAIL' if require_reference_numbering else 'WARN'),ref_detail)
        typed_ref_prefix=[txt for _,txt in ref_paras if re.match(r'^\s*[\[［【(（]?\d+[\]］】)）.、]?\s+',txt)]
        add('reference typed prefixes','FAIL' if require_reference_numbering and typed_ref_prefix else ('WARN' if typed_ref_prefix else 'PASS'),f'{len(typed_ref_prefix)} typed prefixes remain' if typed_ref_prefix else 'bibliography numbering is not embedded in entry text')
        ref_para_fields=[]
        if doc is not None:
            for p in doc.iter(W+'p'):
                ft=field_text(p)
                if 'REF _FmtRef' in ft:ref_para_fields.append(ft)
        bad_ref_fields=[x for x in ref_para_fields if '\\N' not in x.upper()]
        add('bibliography body REF switch','FAIL' if require_reference_numbering and bad_ref_fields else ('WARN' if bad_ref_fields else 'PASS'),f'{len(bad_ref_fields)} bibliography REF field(s) lack paragraph-number switch' if bad_ref_fields else 'bibliography REF fields use paragraph-number switch')

    overall='FAIL' if any(c['status']=='FAIL' for c in checks) else 'WARN' if any(c['status']=='WARN' for c in checks) else 'PASS'
    return {'file':str(path),'profile':str(profile.get('_profile_path','')),'overall':overall,'checks':checks}

def main():
    if hasattr(sys.stdout,'reconfigure'):sys.stdout.reconfigure(encoding='utf-8')
    ap=argparse.ArgumentParser();ap.add_argument('docx',type=Path);ap.add_argument('--profile',type=Path);ap.add_argument('--require-toc',action='store_true');ap.add_argument('--require-captions',action='store_true');ap.add_argument('--require-multilevel',action='store_true');ap.add_argument('--require-crossrefs',action='store_true');ap.add_argument('--require-reference-numbering',action='store_true');args=ap.parse_args()
    r=validate(args.docx,args.profile,args.require_toc,args.require_captions,args.require_multilevel,args.require_crossrefs,args.require_reference_numbering);print(json.dumps(r,ensure_ascii=False,indent=2));raise SystemExit(1 if r['overall']=='FAIL' else 0)
if __name__=='__main__':main()
