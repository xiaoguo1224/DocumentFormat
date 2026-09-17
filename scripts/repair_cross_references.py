#!/usr/bin/env python3
"""Repair native, clickable bibliography/figure/table/equation cross-references."""

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

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
NS = {"w": W, "m": M}


def qn(name: str) -> str:
    prefix, local = name.split(":", 1)
    return f"{{{NS[prefix]}}}{local}"


def text_of(paragraph: etree._Element) -> str:
    return "".join(paragraph.xpath(".//w:t/text()", namespaces=NS))


def instructions(paragraph: etree._Element) -> list[str]:
    return [" ".join(x.split()) for x in paragraph.xpath(".//w:instrText/text()", namespaces=NS)]


def style_id(paragraph: etree._Element) -> str:
    values = paragraph.xpath("./w:pPr/w:pStyle/@w:val", namespaces=NS)
    return values[0] if values else ""


def style_names(styles_root: etree._Element | None) -> dict[str, str]:
    if styles_root is None:
        return {}
    result = {}
    for style in styles_root.xpath("./w:style", namespaces=NS):
        sid = style.get(qn("w:styleId"), "")
        names = style.xpath("./w:name/@w:val", namespaces=NS)
        result[sid] = names[0] if names else sid
    return result


def numbered_style_ids(styles_root: etree._Element | None) -> set[str]:
    """返回「自身或继承链上带 numPr」的段落样式 id 集合。"""
    if styles_root is None:
        return set()
    own: dict[str, bool] = {}
    based_on: dict[str, str] = {}
    for style in styles_root.xpath("./w:style[@w:type='paragraph']", namespaces=NS):
        sid = style.get(qn("w:styleId"), "")
        own[sid] = bool(style.xpath("./w:pPr/w:numPr", namespaces=NS))
        parent = style.xpath("./w:basedOn/@w:val", namespaces=NS)
        based_on[sid] = parent[0] if parent else ""

    resolved: set[str] = set()
    for sid in own:
        seen: set[str] = set()
        current = sid
        while current and current not in seen:
            seen.add(current)
            if own.get(current):
                resolved.add(sid)
                break
            current = based_on.get(current, "")
    return resolved


def is_numbered_paragraph(paragraph: etree._Element, numbered_styles: set[str]) -> bool:
    if paragraph.xpath("./w:pPr/w:numPr", namespaces=NS):
        return True
    return style_id(paragraph) in numbered_styles


def clone_rpr(source: etree._Element | None) -> etree._Element | None:
    return deepcopy(source) if source is not None else None


def text_run(text: str, rpr: etree._Element | None = None) -> etree._Element:
    run = etree.Element(qn("w:r"))
    if rpr is not None:
        run.append(clone_rpr(rpr))
    node = etree.SubElement(run, qn("w:t"))
    if text[:1].isspace() or text[-1:].isspace():
        node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    node.text = text
    return run


def field_runs(instruction: str, cached: str, rpr: etree._Element | None = None) -> list[etree._Element]:
    result = []
    for kind, value in (("begin", None), (None, instruction), ("separate", None), (None, cached), ("end", None)):
        run = etree.Element(qn("w:r"))
        if rpr is not None:
            run.append(clone_rpr(rpr))
        if kind:
            node = etree.SubElement(run, qn("w:fldChar"))
            node.set(qn("w:fldCharType"), kind)
            if kind == "begin":
                node.set(qn("w:dirty"), "true")
        elif value == instruction:
            node = etree.SubElement(run, qn("w:instrText"))
            node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            node.text = value
        else:
            node = etree.SubElement(run, qn("w:t"))
            node.text = value
        result.append(run)
    return result


class BookmarkFactory:
    def __init__(self, root: etree._Element):
        ids = [int(v) for v in root.xpath("//w:bookmarkStart/@w:id", namespaces=NS) if str(v).isdigit()]
        self.next_id = max(ids, default=0) + 1
        self.names = set(root.xpath("//w:bookmarkStart/@w:name", namespaces=NS))

    def pair(self, preferred: str) -> tuple[etree._Element, etree._Element, str]:
        name = preferred
        suffix = 2
        while name in self.names:
            name = f"{preferred}_{suffix}"
            suffix += 1
        self.names.add(name)
        value = str(self.next_id)
        self.next_id += 1
        start = etree.Element(qn("w:bookmarkStart"))
        start.set(qn("w:id"), value)
        start.set(qn("w:name"), name)
        end = etree.Element(qn("w:bookmarkEnd"))
        end.set(qn("w:id"), value)
        return start, end, name


def top_run_child(paragraph: etree._Element, node: etree._Element) -> etree._Element | None:
    current = node
    while current is not None and current.getparent() is not paragraph:
        current = current.getparent()
    return current if current is not None and current.tag == qn("w:r") else None


def tab_run(rpr: etree._Element | None = None) -> etree._Element:
    run = etree.Element(qn("w:r"))
    if rpr is not None:
        run.append(clone_rpr(rpr))
    etree.SubElement(run, qn("w:tab"))
    return run


# 参考文献条目里手工输入的编号前缀：`[1]`、`［1］`、`(1)`、`1.`、`1、` 等
REFERENCE_PREFIX = re.compile(
    r"^(?:"
    r"[\[［【(（]\s*(?P<bracket>\d+)\s*[\]］】)）]"
    r"|(?P<plain>\d+)\s*[.、)）]"
    r")\s*"
)


def add_reference_targets(paragraphs, names, factory, report, numbered_styles) -> dict[str, dict]:
    """把参考文献条目的手工编号换成原生结构，并建立可供正文引用的书签。

    条目所在样式若绑定了原生编号（模板 References 样式已绑定 [%1] + 制表符），
    则只剥掉手打前缀并加书签，编号完全交给 Word 列表；
    否则退回 SEQ 域方案，保证在未升级的模板上仍然可用。
    """
    mapping: dict[str, dict] = {}
    for paragraph in paragraphs:
        style = names.get(style_id(paragraph), "").lower()
        if "reference" not in style and "参考" not in style:
            continue
        visible = text_of(paragraph)
        match = REFERENCE_PREFIX.match(visible)
        if not match:
            continue
        number = match.group("bracket") or match.group("plain")
        first_text = next(iter(paragraph.xpath(".//w:t", namespaces=NS)), None)
        if first_text is None:
            continue
        run = top_run_child(paragraph, first_text)
        if run is None or not (first_text.text or "").startswith(match.group(0)):
            report["unresolved"].append(f"reference target [{number}] has split prefix")
            continue
        first_text.text = (first_text.text or "")[len(match.group(0)) :]
        if first_text.text[:1].isspace():
            first_text.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        rpr = run.find(qn("w:rPr"))

        if is_numbered_paragraph(paragraph, numbered_styles):
            mode = "list"
            start, end, bookmark = factory.pair(f"_FmtRef{int(number):04d}")
            ppr = paragraph.find(qn("w:pPr"))
            paragraph.insert(1 if ppr is not None else 0, start)
            paragraph.append(end)
        else:
            mode = "seq"
            start, end, bookmark = factory.pair(f"_FmtRef{int(number):04d}")
            nodes = [text_run("[", rpr), start, *field_runs(" SEQ Ref \\* ARABIC ", number, rpr), end,
                     text_run("]", rpr), tab_run(rpr)]
            index = paragraph.index(run)
            for offset, node in enumerate(nodes):
                paragraph.insert(index + offset, node)

        mapping[number] = {"bookmark": bookmark, "mode": mode}
        report["targets"]["references"] += 1
        report["reference_numbering"][mode] += 1
    return mapping


def field_range(paragraph: etree._Element) -> tuple[int, int] | None:
    children = list(paragraph)
    begins = []
    ends = []
    for index, child in enumerate(children):
        if child.tag != qn("w:r"):
            continue
        chars = child.xpath("./w:fldChar/@w:fldCharType", namespaces=NS)
        if chars == ["begin"]:
            begins.append(index)
        elif chars == ["end"]:
            ends.append(index)
    return (min(begins), max(ends)) if begins and ends else None


def add_caption_targets(paragraphs, names, factory, report) -> dict[str, dict[str, str]]:
    mapping = {"figure": {}, "table": {}, "equation": {}}
    for paragraph in paragraphs:
        visible = text_of(paragraph).strip()
        fields = instructions(paragraph)
        seq = next((f for f in fields if f.upper().startswith("SEQ ")), "")
        if not seq:
            continue
        label = seq.split()[1]
        kind = {"图": "figure", "表": "table", "式": "equation", "Figure": "figure", "Table": "table"}.get(label)
        if not kind:
            continue
        match = re.search(r"(?:图|表|式|Figure|Table)\s*(\d+(?:[-－.]\d+)?)", visible, re.I)
        if not match:
            report["unresolved"].append(f"{kind} caption has no visible number: {visible[:60]}")
            continue
        number = match.group(1).replace("－", "-").replace(".", "-")
        span = field_range(paragraph)
        if span is None:
            report["unresolved"].append(f"{kind} {number} field range not found")
            continue
        prefix = {"figure": "Fig", "table": "Tbl", "equation": "Eq"}[kind]
        start, end, bookmark = factory.pair(f"_Fmt{prefix}{len(mapping[kind]) + 1:04d}")
        paragraph.insert(span[0], start)
        paragraph.insert(span[1] + 2, end)
        mapping[kind][number] = bookmark
        report["targets"][kind + "s"] += 1
    return mapping


def add_typed_equation_targets(paragraphs, names, factory, mapping, report) -> None:
    for paragraph in paragraphs:
        if any(f.upper().startswith("SEQ 式") for f in instructions(paragraph)):
            continue
        style = names.get(style_id(paragraph), "").lower()
        has_math = bool(paragraph.xpath(".//m:oMath|.//m:oMathPara", namespaces=NS))
        if not has_math and "equation" not in style and "公式" not in style:
            continue
        visible = text_of(paragraph)
        match = re.search(r"(?:式\s*)?[（(]?\s*(\d+)(?:[-－.](\d+))?\s*[）)]?\s*$", visible)
        if not match:
            report["unresolved"].append(f"numbered equation not recognized: {visible[:60]}")
            continue
        chapter, sequence = match.group(1), match.group(2) or match.group(1)
        number = chapter if match.group(2) is None else f"{chapter}-{sequence}"
        text_nodes = paragraph.xpath(".//w:t", namespaces=NS)
        if not text_nodes:
            report["unresolved"].append(f"equation {number} number is not editable text")
            continue
        last = text_nodes[-1]
        last.text = re.sub(r"(?:式\s*)?[（(]?\s*\d+(?:[-－.]\d+)?\s*[）)]?\s*$", "", last.text or "")
        rpr = last.getparent().find(qn("w:rPr")) if last.getparent() is not None else None
        start, end, bookmark = factory.pair(f"_FmtEq{len(mapping['equation']) + 1:04d}")
        paragraph.append(text_run("（", rpr))
        paragraph.append(start)
        if match.group(2) is not None:
            for node in field_runs(" STYLEREF 1 \\n ", chapter, rpr):
                paragraph.append(node)
            paragraph.append(text_run("-", rpr))
            instruction = " SEQ 式 \\* ARABIC \\s 1 "
        else:
            instruction = " SEQ 式 \\* ARABIC "
        for node in field_runs(instruction, sequence, rpr):
            paragraph.append(node)
        paragraph.append(end)
        paragraph.append(text_run("）", rpr))
        mapping["equation"][number] = bookmark
        report["targets"]["equations"] += 1


TOKEN = re.compile(
    r"\[(?P<cites>\d+(?:\s*[,，;；、\-–—]\s*\d+)*)\]"
    r"|(?P<label>图|表|式|公式|Figure|Table)\s*(?P<number>\d+(?:[-－.]\d+)?)",
    re.I,
)


def paragraph_rpr(paragraph: etree._Element) -> etree._Element | None:
    rprs = paragraph.xpath(".//w:r[w:t]/w:rPr", namespaces=NS)
    signatures = {etree.tostring(node) for node in rprs}
    if len(signatures) > 1:
        return None
    return rprs[0] if rprs else None


def append_ref(paragraph, bookmark, cached, rpr):
    for node in field_runs(f" REF {bookmark} \\h ", cached, rpr):
        paragraph.append(node)


def repair_body_references(paragraphs, names, ref_map, target_map, report) -> None:
    excluded = ("toc", "caption", "reference", "目录", "题注", "参考")
    for paragraph in paragraphs:
        style = names.get(style_id(paragraph), "").lower()
        if any(token in style for token in excluded):
            continue
        if instructions(paragraph) or paragraph.xpath(".//w:hyperlink|.//w:drawing|.//w:pict|.//m:oMath|.//m:oMathPara", namespaces=NS):
            continue
        visible = text_of(paragraph)
        matches = list(TOKEN.finditer(visible))
        if not matches:
            continue
        rpr = paragraph_rpr(paragraph)
        if rpr is None and len(paragraph.xpath(".//w:r[w:t]", namespaces=NS)) > 1:
            report["ambiguous"].append(f"mixed formatting: {visible[:100]}")
            continue
        ppr = paragraph.find(qn("w:pPr"))
        for child in list(paragraph):
            if child is not ppr:
                paragraph.remove(child)
        cursor = 0
        for match in matches:
            if match.start() > cursor:
                paragraph.append(text_run(visible[cursor : match.start()], rpr))
            original = match.group(0)
            if match.group("cites"):
                numbers = re.findall(r"\d+", match.group("cites"))
                if any(number not in ref_map for number in numbers):
                    paragraph.append(text_run(original, rpr))
                    report["unresolved"].append(f"citation target missing: {original}")
                else:
                    paragraph.append(text_run("[", rpr))
                    inner = match.group("cites")
                    inner_cursor = 0
                    for number_match in re.finditer(r"\d+", inner):
                        if number_match.start() > inner_cursor:
                            paragraph.append(text_run(inner[inner_cursor : number_match.start()], rpr))
                        number = number_match.group(0)
                        append_ref(paragraph, ref_map[number], number, rpr)
                        report["converted"]["citations"] += 1
                        inner_cursor = number_match.end()
                    if inner_cursor < len(inner):
                        paragraph.append(text_run(inner[inner_cursor:], rpr))
                    paragraph.append(text_run("]", rpr))
            else:
                label = match.group("label")
                number = match.group("number").replace("－", "-").replace(".", "-")
                kind = "figure" if label.lower() in {"图", "figure"} else "table" if label.lower() in {"表", "table"} else "equation"
                bookmark = target_map[kind].get(number)
                if bookmark is None:
                    paragraph.append(text_run(original, rpr))
                    report["unresolved"].append(f"{kind} target missing: {original}")
                else:
                    prefix = original[: original.find(match.group("number"))]
                    paragraph.append(text_run(prefix, rpr))
                    append_ref(paragraph, bookmark, match.group("number"), rpr)
                    report["converted"][kind + "s"] += 1
            cursor = match.end()
        if cursor < len(visible):
            paragraph.append(text_run(visible[cursor:], rpr))


def repair(input_path: Path, output_path: Path, strict: bool = False) -> dict:
    with zipfile.ZipFile(input_path) as zin:
        document = etree.fromstring(zin.read("word/document.xml"))
        styles = etree.fromstring(zin.read("word/styles.xml")) if "word/styles.xml" in zin.namelist() else None
        files = {info.filename: zin.read(info.filename) for info in zin.infolist()}
    paragraphs = document.xpath("//w:body//w:p", namespaces=NS)
    names = style_names(styles)
    report = {
        "targets": {"references": 0, "figures": 0, "tables": 0, "equations": 0},
        "converted": {"citations": 0, "figures": 0, "tables": 0, "equations": 0},
        "unresolved": [],
        "ambiguous": [],
    }
    factory = BookmarkFactory(document)
    ref_map = add_reference_targets(paragraphs, names, factory, report)
    target_map = add_caption_targets(paragraphs, names, factory, report)
    add_typed_equation_targets(paragraphs, names, factory, target_map, report)
    repair_body_references(paragraphs, names, ref_map, target_map, report)
    files["word/document.xml"] = etree.tostring(document, xml_declaration=True, encoding="UTF-8", standalone="yes")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    target = output_path
    if input_path.resolve() == output_path.resolve():
        fd, temp_name = tempfile.mkstemp(prefix="crossrefs-", suffix=".docx", dir=output_path.parent)
        os.close(fd)
        target = Path(temp_name)
    try:
        with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as zout:
            for name, data in files.items():
                zout.writestr(name, data)
        if target != output_path:
            target.replace(output_path)
    finally:
        if target != output_path and target.exists():
            target.unlink()
    if strict and (report["unresolved"] or report["ambiguous"]):
        raise RuntimeError(json.dumps(report, ensure_ascii=False))
    return report


def main() -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--strict", action="store_true")
    args = parser.parse_args()
    if not args.input.is_file():
        parser.error(f"file not found: {args.input}")
    print(json.dumps(repair(args.input, args.output, args.strict), ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
