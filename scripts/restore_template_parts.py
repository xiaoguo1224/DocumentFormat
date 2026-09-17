#!/usr/bin/env python3
"""Restore template-native styles/numbering/settings after Word refreshes fields."""

from __future__ import annotations

import argparse
import shutil
import tempfile
import zipfile
from pathlib import Path

from fast_format_docx import DEFAULT_TEMPLATE, apply_template_parts
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def style_maps(docx: Path, template: Path) -> dict[str, str]:
    def read(path: Path):
        with zipfile.ZipFile(path) as zf:
            root = etree.fromstring(zf.read("word/styles.xml"))
        result = {}
        for style in root.findall(f"{{{W}}}style"):
            sid = style.get(f"{{{W}}}styleId")
            name = style.find(f"{{{W}}}name")
            if sid and name is not None and name.get(f"{{{W}}}val"):
                result[sid] = name.get(f"{{{W}}}val").casefold()
        return result

    source = read(docx)
    target_by_name = {name: sid for sid, name in read(template).items()}
    return {sid: target_by_name[name] for sid, name in source.items() if name in target_by_name}


def remap_style_references(source: Path, template: Path, output: Path) -> None:
    mapping = style_maps(source, template)
    with zipfile.ZipFile(source) as zin, zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename.startswith("word/") and info.filename.endswith(".xml") and info.filename not in {"word/styles.xml", "word/numbering.xml"}:
                try:
                    root = etree.fromstring(data)
                    changed = False
                    for tag in ("pStyle", "rStyle", "tblStyle"):
                        for node in root.iter(f"{{{W}}}{tag}"):
                            old = node.get(f"{{{W}}}val")
                            if old in mapping and mapping[old] != old:
                                node.set(f"{{{W}}}val", mapping[old])
                                changed = True
                    if changed:
                        data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")
                except etree.XMLSyntaxError:
                    pass
            zout.writestr(info, data)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--template", type=Path, default=DEFAULT_TEMPLATE)
    args = parser.parse_args()
    for path in (args.input, args.template):
        if not path.is_file():
            parser.error(f"file not found: {path}")
    args.output.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="restore-docx-template-") as temp:
        remapped = Path(temp) / "remapped.docx"
        staged = Path(temp) / "restored.docx"
        remap_style_references(args.input, args.template, remapped)
        apply_template_parts(args.template, remapped, staged)
        shutil.copy2(staged, args.output)
    print(args.output)


if __name__ == "__main__":
    main()
