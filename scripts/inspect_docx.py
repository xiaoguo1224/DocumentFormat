#!/usr/bin/env python3
"""Inspect native Word structure plus reusable visual style/section metadata."""
from __future__ import annotations
import argparse,json,sys,zipfile
from collections import Counter
from pathlib import Path
from xml.etree import ElementTree as ET
W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

def x(z,n):
    try:return ET.fromstring(z.read(n))
    except KeyError:return None

def a(el,k='val'):return el.get(W+k) if el is not None else None

def on(el):
    if el is None:return False
    v=a(el);return v is None or str(v).lower() not in {'0','false','off','no'}

def style_payload(st):
    ppr=st.find(W+'pPr');rpr=st.find(W+'rPr');num_id=ilvl=None
    if ppr is not None:
        np=ppr.find(W+'numPr')
        if np is not None:num_id=a(np.find(W+'numId'));ilvl=a(np.find(W+'ilvl'))
    fonts=rpr.find(W+'rFonts') if rpr is not None else None
    spacing=ppr.find(W+'spacing') if ppr is not None else None
    ind=ppr.find(W+'ind') if ppr is not None else None
    return {
      'style_id':a(st,'styleId'),'name':a(st.find(W+'name')),'type':a(st,'type'),'based_on':a(st.find(W+'basedOn')),
      'outline_level':a(ppr.find(W+'outlineLvl')) if ppr is not None else None,'page_break_before':on(ppr.find(W+'pageBreakBefore')) if ppr is not None else False,
      'keep_with_next':on(ppr.find(W+'keepNext')) if ppr is not None else False,'keep_together':on(ppr.find(W+'keepLines')) if ppr is not None else False,
      'num_id':num_id,'num_level':ilvl,
      'paragraph':{'alignment':a(ppr.find(W+'jc')) if ppr is not None else None,'before':a(spacing,'before'),'after':a(spacing,'after'),'line':a(spacing,'line'),'line_rule':a(spacing,'lineRule'),'left':a(ind,'left'),'right':a(ind,'right'),'first_line':a(ind,'firstLine'),'hanging':a(ind,'hanging')},
      'run':{'ascii_font':a(fonts,'ascii'),'east_asia_font':a(fonts,'eastAsia'),'hansi_font':a(fonts,'hAnsi'),'size_half_points':a(rpr.find(W+'sz')) if rpr is not None else None,'bold':on(rpr.find(W+'b')) if rpr is not None else False,'italic':on(rpr.find(W+'i')) if rpr is not None else False}
    }

def inspect(path:Path):
    with zipfile.ZipFile(path) as z:
        styles=x(z,'word/styles.xml');numbering=x(z,'word/numbering.xml');doc=x(z,'word/document.xml');settings=x(z,'word/settings.xml')
        result={'file':str(path),'styles':[],'numbering':{'multilevel':[]},'caption_labels':[],'document':{},'sections':[],'headers':sorted(n for n in z.namelist() if n.startswith('word/header') and n.endswith('.xml')),'footers':sorted(n for n in z.namelist() if n.startswith('word/footer') and n.endswith('.xml'))}
        if styles is not None:result['styles']=[style_payload(st) for st in styles.findall(W+'style')]
        if numbering is not None:
            for ab in numbering.findall(W+'abstractNum'):
                levels=[]
                for lvl in ab.findall(W+'lvl'):
                    levels.append({'level':a(lvl,'ilvl'),'start':a(lvl.find(W+'start')),'format':a(lvl.find(W+'numFmt')),'text':a(lvl.find(W+'lvlText')),'suffix':a(lvl.find(W+'suff')),'pstyle':a(lvl.find(W+'pStyle'))})
                result['numbering']['multilevel'].append({'abstract_num_id':a(ab,'abstractNumId'),'levels':levels})
            result['numbering']['num_count']=len(numbering.findall(W+'num'))
        if settings is not None:
            caps=settings.find(W+'captions')
            if caps is not None:
                for c in caps.findall(W+'caption'):result['caption_labels'].append({'name':a(c,'name'),'position':a(c,'pos'),'chapter_number':a(c,'chapNum'),'heading_level':a(c,'heading'),'number_format':a(c,'numFmt')})
        if doc is not None:
            usage=Counter();fields=Counter();paras=0
            for p in doc.iter(W+'p'):
                paras+=1;ppr=p.find(W+'pPr')
                if ppr is not None:
                    ps=ppr.find(W+'pStyle')
                    if ps is not None:usage[a(ps)]+=1
                for instr in p.iter(W+'instrText'):
                    t=(instr.text or '').strip()
                    if t:fields[t.split()[0].upper()]+=1
            for sect in doc.iter(W+'sectPr'):
                pg=sect.find(W+'pgSz');mar=sect.find(W+'pgMar');pn=sect.find(W+'pgNumType')
                result['sections'].append({'page_width_twips':a(pg,'w'),'page_height_twips':a(pg,'h'),'orientation':a(pg,'orient'),'margins_twips':{'top':a(mar,'top'),'right':a(mar,'right'),'bottom':a(mar,'bottom'),'left':a(mar,'left'),'header':a(mar,'header'),'footer':a(mar,'footer')},'page_number':{'format':a(pn,'fmt'),'start':a(pn,'start')}})
            result['document']={'paragraph_count':paras,'style_usage':dict(usage),'field_types':dict(fields),'section_count':len(result['sections'])}
        return result

def main():
    if hasattr(sys.stdout,'reconfigure'):sys.stdout.reconfigure(encoding='utf-8')
    ap=argparse.ArgumentParser();ap.add_argument('docx',type=Path);args=ap.parse_args();print(json.dumps(inspect(args.docx),ensure_ascii=False,indent=2))
if __name__=='__main__':main()
