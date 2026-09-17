#!/usr/bin/env python3
"""Combine rendered DOCX page PNGs into one labeled QA sheet."""

from __future__ import annotations

import argparse
import re
from pathlib import Path

from PIL import Image, ImageDraw


def page_number(path: Path) -> int:
    match = re.search(r"(\d+)(?=\D*$)", path.stem)
    return int(match.group(1)) if match else 0


def build(input_dir: Path, output: Path, columns: int, thumb_width: int) -> int:
    pages = sorted(input_dir.glob("*.png"), key=page_number)
    if not pages:
        raise ValueError(f"no PNG pages found in {input_dir}")
    with Image.open(pages[0]) as first:
        ratio = first.height / first.width
    thumb_height = round(thumb_width * ratio)
    label_height, gap = 28, 16
    rows = (len(pages) + columns - 1) // columns
    sheet = Image.new("RGB", (columns * (thumb_width + gap) + gap, rows * (thumb_height + label_height + gap) + gap), "white")
    draw = ImageDraw.Draw(sheet)
    for index, path in enumerate(pages):
        with Image.open(path) as page:
            thumb = page.convert("RGB")
            thumb.thumbnail((thumb_width, thumb_height))
        x = gap + (index % columns) * (thumb_width + gap)
        y = gap + (index // columns) * (thumb_height + label_height + gap)
        sheet.paste(thumb, (x, y + label_height))
        draw.text((x, y + 4), f"Page {page_number(path) or index + 1}", fill="black")
    output.parent.mkdir(parents=True, exist_ok=True)
    sheet.save(output, optimize=True)
    return len(pages)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input_dir", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--columns", type=int, default=4)
    parser.add_argument("--thumb-width", type=int, default=320)
    args = parser.parse_args()
    if args.columns < 1 or args.thumb_width < 100:
        parser.error("columns must be >= 1 and thumb-width must be >= 100")
    print(f"pages={build(args.input_dir, args.output, args.columns, args.thumb_width)}")


if __name__ == "__main__":
    main()
