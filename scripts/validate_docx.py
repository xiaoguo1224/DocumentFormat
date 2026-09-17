#!/usr/bin/env python3
"""Structural validator for Word-native DOCX features.

It detects common visual-only fallbacks and understands style-bound heading numbering.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def read_xml(zf, name):
    try:
        return ET.fromstring(zf.read(name))
    except KeyError:
        return None


def val(el, key="val"):
    return el.get(W + key) if el is not None else None


def onoff(el):
    if el is None:
        return False
    v = val(el)
    return v is None or str(v).lower() not in {"0", "false", "off", "no"}


def paragraph_text(p):
    return "".join(t.text or "" for t in p.iter(W + "t"))


def field_text(root):
    return " ".join((x.text or "").strip() for x in root.iter(W + "instrText") if (x.text or "").strip())


def validate(path: Path, require_toc=False, require_captions=False, require_multilevel=False, require_crossrefs=False):
    checks = []
    with zipfile.ZipFile(path) as zf:
        doc = read_xml(zf, "word/document.xml")
        styles = read_xml(zf, "word/styles.xml")
        numbering = read_xml(zf, "word/numbering.xml")
        settings = read_xml(zf, "word/settings.xml")

        def add(name, status, detail):
            checks.append({"check": name, "status": status, "detail": detail})

        add("styles.xml", "PASS" if styles is not None else "FAIL", "styles present" if styles is not None else "missing word/styles.xml")
        add("numbering.xml", "PASS" if numbering is not None else ("FAIL" if require_multilevel else "WARN"), "numbering present" if numbering is not None else "missing word/numbering.xml")

        style_info = {}
        if styles is not None:
            for st in styles.findall(W + "style"):
                sid = val(st, "styleId")
                ppr = st.find(W + "pPr")
                outline = ppr.find(W + "outlineLvl") if ppr is not None else None
                num_id = ilvl = None
                if ppr is not None:
                    numpr = ppr.find(W + "numPr")
                    if numpr is not None:
                        num_id = val(numpr.find(W + "numId"))
                        ilvl = val(numpr.find(W + "ilvl"))
                style_info[sid] = {
                    "name": val(st.find(W + "name")) or sid,
                    "outline": val(outline),
                    "page_break_before": onoff(ppr.find(W + "pageBreakBefore")) if ppr is not None else False,
                    "num_id": num_id,
                    "ilvl": ilvl,
                }

        # Strong multilevel check: Heading1-4 should share a numId and map to levels 0-3.
        heading_bindings = [style_info.get(f"Heading{i}") for i in range(1, 5)]
        heading_native_ok = all(heading_bindings)
        if heading_native_ok:
            ids = [x["num_id"] for x in heading_bindings]
            levels = [x["ilvl"] for x in heading_bindings]
            heading_native_ok = all(ids) and len(set(ids)) == 1 and levels == ["0", "1", "2", "3"]
        add(
            "Heading 1-4 multilevel binding",
            "PASS" if heading_native_ok else ("FAIL" if require_multilevel else "WARN"),
            "Heading1-4 share one native multilevel list with levels 0-3" if heading_native_ok else "Heading1-4 are not fully bound to one native multilevel numbering definition",
        )

        # Verify the referenced numbering definition really has at least four levels.
        numbering_levels_ok = False
        if numbering is not None and heading_native_ok:
            num_id = heading_bindings[0]["num_id"]
            abstract_id = None
            for num in numbering.findall(W + "num"):
                if val(num, "numId") == num_id:
                    abstract_id = val(num.find(W + "abstractNumId"))
                    break
            if abstract_id is not None:
                for abstract in numbering.findall(W + "abstractNum"):
                    if val(abstract, "abstractNumId") == abstract_id:
                        levels = abstract.findall(W + "lvl")
                        numbering_levels_ok = len(levels) >= 4
                        break
        add(
            "multilevel numbering definition",
            "PASS" if numbering_levels_ok else ("FAIL" if require_multilevel else "WARN"),
            "4+ numbering levels found" if numbering_levels_ok else "required 4-level heading numbering definition not confirmed",
        )

        fields = []
        heading_paras = []
        static_heading_number_candidates = []
        caption_candidates = []
        static_crossref_candidates = []

        if doc is not None:
            fields.append(field_text(doc))
            for p in doc.iter(W + "p"):
                txt = paragraph_text(p).strip()
                ppr = p.find(W + "pPr")
                sid = None
                if ppr is not None:
                    ps = ppr.find(W + "pStyle")
                    sid = val(ps)
                    if sid in style_info and style_info[sid]["outline"] is not None:
                        heading_paras.append((sid, txt, ppr))
                is_outline = sid in style_info and style_info[sid]["outline"] is not None
                if is_outline and sid not in {"TOC1", "TOC2", "TOC3", "TOC4"} and re.match(r"^\d+(?:\.\d+){0,5}\s+\S", txt):
                    static_heading_number_candidates.append(txt[:80])
                if re.match(r"^(图|表|Figure|Table)\s*\d", txt, re.I):
                    caption_candidates.append((txt[:100], field_text(p)))
                style_name = style_info.get(sid, {}).get("name", "").lower()
                excluded = any(x in style_name for x in ("toc", "caption", "reference", "目录", "题注", "参考"))
                para_fields = field_text(p).upper()
                if not excluded and " REF " not in f" {para_fields} ":
                    for match in re.finditer(r"\[(?:\d+[\s,，;；、\-–—]*)+\]|(?:图|表|式|公式|Figure|Table)\s*\d+(?:[-－.]\d+)?", txt, re.I):
                        static_crossref_candidates.append(match.group(0))

        # Include fields in headers/footers.
        for name in zf.namelist():
            if (name.startswith("word/header") or name.startswith("word/footer")) and name.endswith(".xml"):
                root = read_xml(zf, name)
                if root is not None:
                    fields.append(field_text(root))

        field_blob = "\n".join(fields).upper()
        toc_ok = re.search(r"(^|\s)TOC(\s|$)", field_blob) is not None
        seq_ok = re.search(r"(^|\s)SEQ\s", field_blob) is not None
        ref_targets = re.findall(r"(?:^|\s)REF\s+([A-Z_][A-Z0-9_.]*)", field_blob, re.I)
        ref_ok = bool(ref_targets)
        page_ok = re.search(r"(^|\s)PAGE(\s|$)", field_blob) is not None

        add("TOC field", "PASS" if toc_ok else ("FAIL" if require_toc else "WARN"), "TOC field found" if toc_ok else "no TOC field detected")
        add("SEQ caption fields", "PASS" if seq_ok else ("FAIL" if require_captions else "WARN"), "SEQ field(s) found" if seq_ok else "no SEQ fields detected")

        caption_labels = set()
        if settings is not None:
            caps = settings.find(W + "captions")
            if caps is not None:
                for cap in caps.findall(W + "caption"):
                    name = val(cap, "name")
                    if name:
                        caption_labels.add(name)
        labels_ok = {"图", "表"}.issubset(caption_labels)
        add(
            "registered caption labels",
            "PASS" if labels_ok else ("FAIL" if require_captions else "WARN"),
            "registered native labels include 图 and 表" if labels_ok else f"registered labels found: {sorted(caption_labels)}",
        )

        add(
            "cross-reference fields",
            "PASS" if ref_ok else ("FAIL" if require_crossrefs and static_crossref_candidates else "WARN"),
            f"{len(ref_targets)} REF field(s) found" if ref_ok else "no REF fields detected",
        )
        bookmark_names = set()
        if doc is not None:
            bookmark_names = {val(node, "name").upper() for node in doc.iter(W + "bookmarkStart") if val(node, "name")}
        missing_targets = sorted({target for target in ref_targets if target.upper() not in bookmark_names})
        add(
            "REF target integrity",
            "FAIL" if missing_targets else "PASS",
            f"missing bookmark target(s): {missing_targets}" if missing_targets else "all REF targets resolve to bookmarks",
        )
        add(
            "remaining typed cross-references",
            "FAIL" if require_crossrefs and static_crossref_candidates else ("WARN" if static_crossref_candidates else "PASS"),
            f"{len(static_crossref_candidates)} typed reference(s) remain: {static_crossref_candidates[:8]}" if static_crossref_candidates else "no recognizable typed cross-references remain in body text",
        )
        add("page-number fields", "PASS" if page_ok else "WARN", "PAGE field found" if page_ok else "no PAGE field detected in main document/header/footer XML")

        if heading_paras:
            add("outline headings", "PASS", f"{len(heading_paras)} heading/outline paragraphs detected")
        else:
            add("outline headings", "WARN", "no outline-level heading paragraphs detected from styles")

        if static_heading_number_candidates:
            add("possible typed heading numbers", "WARN", f"{len(static_heading_number_candidates)} paragraph(s) look manually prefixed; inspect if numbering is truly native")
        else:
            add("possible typed heading numbers", "PASS", "no obvious typed numeric heading prefixes detected")

        if caption_candidates:
            without_seq = [x for x in caption_candidates if "SEQ " not in (x[1] or "").upper()]
            add("possible static captions", "WARN" if without_seq else "PASS", f"{len(without_seq)} caption-looking paragraph(s) have no SEQ field" if without_seq else "caption-looking paragraphs contain SEQ fields")

        h1_pagebreak = style_info.get("Heading1", {}).get("page_break_before", False)
        add("Heading 1 pageBreakBefore", "PASS" if h1_pagebreak else "FAIL", "Heading 1 style uses pageBreakBefore" if h1_pagebreak else "Heading 1 style does not use pageBreakBefore")

    overall = "PASS"
    if any(c["status"] == "FAIL" for c in checks):
        overall = "FAIL"
    elif any(c["status"] == "WARN" for c in checks):
        overall = "WARN"
    return {"file": str(path), "overall": overall, "checks": checks}


def main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser()
    ap.add_argument("docx", type=Path)
    ap.add_argument("--require-toc", action="store_true")
    ap.add_argument("--require-captions", action="store_true")
    ap.add_argument("--require-multilevel", action="store_true")
    ap.add_argument("--require-crossrefs", action="store_true")
    args = ap.parse_args()
    result = validate(args.docx, args.require_toc, args.require_captions, args.require_multilevel, args.require_crossrefs)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    raise SystemExit(1 if result["overall"] == "FAIL" else 0)


if __name__ == "__main__":
    main()
