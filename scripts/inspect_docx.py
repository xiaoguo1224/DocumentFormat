#!/usr/bin/env python3
"""Inspect Word-native structures in a .docx/.dotx template.

This reports structure rather than trying to infer every visual rule.
"""

from __future__ import annotations

import argparse
import json
import sys
import zipfile
from collections import Counter
from pathlib import Path
from xml.etree import ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def xml_from_zip(zf: zipfile.ZipFile, name: str):
    try:
        return ET.fromstring(zf.read(name))
    except KeyError:
        return None


def attr(el, name):
    return el.get(W + name) if el is not None else None


def onoff(el):
    if el is None:
        return False
    v = attr(el, "val")
    return v is None or str(v).lower() not in {"0", "false", "off", "no"}


def inspect(path: Path) -> dict:
    with zipfile.ZipFile(path) as zf:
        styles = xml_from_zip(zf, "word/styles.xml")
        numbering = xml_from_zip(zf, "word/numbering.xml")
        document = xml_from_zip(zf, "word/document.xml")

        settings = xml_from_zip(zf, "word/settings.xml")
        result = {
            "file": str(path),
            "has_styles_xml": styles is not None,
            "has_numbering_xml": numbering is not None,
            "styles": [],
            "numbering": {},
            "caption_labels": [],
            "document": {},
            "headers": sorted([n for n in zf.namelist() if n.startswith("word/header") and n.endswith(".xml")]),
            "footers": sorted([n for n in zf.namelist() if n.startswith("word/footer") and n.endswith(".xml")]),
        }

        if styles is not None:
            for st in styles.findall(W + "style"):
                style_id = attr(st, "styleId")
                typ = attr(st, "type")
                name_el = st.find(W + "name")
                based = st.find(W + "basedOn")
                ppr = st.find(W + "pPr")
                outline = ppr.find(W + "outlineLvl") if ppr is not None else None
                page_break = ppr.find(W + "pageBreakBefore") if ppr is not None else None
                num_id = None
                ilvl = None
                if ppr is not None:
                    numpr = ppr.find(W + "numPr")
                    if numpr is not None:
                        n = numpr.find(W + "numId")
                        l = numpr.find(W + "ilvl")
                        num_id = attr(n, "val")
                        ilvl = attr(l, "val")
                result["styles"].append({
                    "style_id": style_id,
                    "name": attr(name_el, "val"),
                    "type": typ,
                    "based_on": attr(based, "val"),
                    "outline_level": attr(outline, "val"),
                    "page_break_before": onoff(page_break),
                    "num_id": num_id,
                    "num_level": ilvl,
                })

        if numbering is not None:
            abs_nums = numbering.findall(W + "abstractNum")
            nums = numbering.findall(W + "num")
            result["numbering"] = {
                "abstract_num_count": len(abs_nums),
                "num_count": len(nums),
                "multilevel": [],
            }
            for a in abs_nums:
                levels = []
                for lvl in a.findall(W + "lvl"):
                    ilvl = attr(lvl, "ilvl")
                    num_fmt = attr(lvl.find(W + "numFmt"), "val")
                    lvl_text = attr(lvl.find(W + "lvlText"), "val")
                    pstyle = attr(lvl.find(W + "pStyle"), "val")
                    levels.append({"level": ilvl, "format": num_fmt, "text": lvl_text, "pstyle": pstyle})
                result["numbering"]["multilevel"].append({
                    "abstract_num_id": attr(a, "abstractNumId"),
                    "levels": levels,
                })

        if settings is not None:
            caps = settings.find(W + "captions")
            if caps is not None:
                for cap in caps.findall(W + "caption"):
                    result["caption_labels"].append({
                        "name": attr(cap, "name"),
                        "position": attr(cap, "pos"),
                        "chapter_number": attr(cap, "chapNum"),
                        "heading_level": attr(cap, "heading"),
                        "number_format": attr(cap, "numFmt"),
                    })

        if document is not None:
            style_usage = Counter()
            fields = Counter()
            page_break_before_count = 0
            numpr_count = 0
            paragraphs = 0
            for p in document.iter(W + "p"):
                paragraphs += 1
                ppr = p.find(W + "pPr")
                if ppr is not None:
                    ps = ppr.find(W + "pStyle")
                    if ps is not None:
                        style_usage[attr(ps, "val")] += 1
                    if onoff(ppr.find(W + "pageBreakBefore")):
                        page_break_before_count += 1
                    if ppr.find(W + "numPr") is not None:
                        numpr_count += 1
                for instr in p.iter(W + "instrText"):
                    text = (instr.text or "").strip()
                    if text:
                        head = text.split()[0].upper()
                        fields[head] += 1
            # Include fields from headers/footers too.
            for name in zf.namelist():
                if (name.startswith("word/header") or name.startswith("word/footer")) and name.endswith(".xml"):
                    root = xml_from_zip(zf, name)
                    if root is None:
                        continue
                    for instr in root.iter(W + "instrText"):
                        text = (instr.text or "").strip()
                        if text:
                            fields[text.split()[0].upper()] += 1
            result["document"] = {
                "paragraph_count": paragraphs,
                "style_usage": dict(style_usage),
                "field_types": dict(fields),
                "page_break_before_paragraphs": page_break_before_count,
                "numbered_paragraphs_direct": numpr_count,
                "section_count": len(list(document.iter(W + "sectPr"))),
            }

        return result


def main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser()
    ap.add_argument("docx", type=Path)
    ap.add_argument("--json", action="store_true")
    args = ap.parse_args()
    data = inspect(args.docx)
    print(json.dumps(data, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
