from __future__ import annotations

import argparse
import json
import math
from pathlib import Path, PurePosixPath
from zipfile import ZipFile

from lxml import etree
from PIL import Image, ImageDraw, ImageFont


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "v": "urn:schemas-microsoft-com:vml",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("source", type=Path)
    parser.add_argument("output_dir", type=Path)
    args = parser.parse_args()

    with ZipFile(args.source) as archive:
        rels_root = etree.fromstring(archive.read("word/_rels/document.xml.rels"))
        rels = {
            node.get("Id"): node.get("Target")
            for node in rels_root.xpath(".//pr:Relationship", namespaces=NS)
        }
        doc_root = etree.fromstring(archive.read("word/document.xml"))

    records = []
    body = doc_root.find("w:body", NS)
    for block_index, block in enumerate(body, start=1):
        if etree.QName(block).localname != "p":
            continue
        text = "".join(block.xpath(".//w:t/text()", namespaces=NS)).strip()
        rel_ids = block.xpath(".//a:blip/@r:embed", namespaces=NS)
        rel_ids += block.xpath(".//v:imagedata/@r:id", namespaces=NS)
        for rel_id in rel_ids:
            target = rels.get(rel_id, "")
            name = PurePosixPath(target).name if target else ""
            records.append(
                {
                    "block": block_index,
                    "paragraph_text": text,
                    "relationship_id": rel_id,
                    "target": target,
                    "file": name,
                }
            )

    media_dir = args.output_dir / "media"
    for record in records:
        path = media_dir / record["file"]
        try:
            with Image.open(path) as im:
                record["width"] = im.width
                record["height"] = im.height
                record["format"] = im.format
        except Exception as exc:
            record["image_error"] = str(exc)

    (args.output_dir / "image_map.json").write_text(
        json.dumps(records, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    unique = []
    seen = set()
    for record in records:
        if record["file"] and record["file"] not in seen:
            seen.add(record["file"])
            unique.append(record)

    cell_w, cell_h = 360, 280
    cols, rows = 4, 4
    per_sheet = cols * rows
    font = ImageFont.load_default()
    sheet_dir = args.output_dir / "contact_sheets"
    sheet_dir.mkdir(exist_ok=True)

    for sheet_index in range(math.ceil(len(unique) / per_sheet)):
        batch = unique[sheet_index * per_sheet : (sheet_index + 1) * per_sheet]
        sheet = Image.new("RGB", (cols * cell_w, rows * cell_h), "white")
        draw = ImageDraw.Draw(sheet)
        for index, record in enumerate(batch):
            x = (index % cols) * cell_w
            y = (index // cols) * cell_h
            label = f"{record['file']}  P{record['block']:05d}  {record.get('width','?')}x{record.get('height','?')}"
            draw.text((x + 6, y + 5), label, fill="black", font=font)
            path = media_dir / record["file"]
            try:
                with Image.open(path) as im:
                    preview = im.convert("RGB")
                    preview.thumbnail((cell_w - 12, cell_h - 36))
                    px = x + (cell_w - preview.width) // 2
                    py = y + 28 + (cell_h - 32 - preview.height) // 2
                    sheet.paste(preview, (px, py))
            except Exception as exc:
                draw.text((x + 6, y + 35), str(exc), fill="red", font=font)
            draw.rectangle((x, y, x + cell_w - 1, y + cell_h - 1), outline="#cccccc")
        out = sheet_dir / f"sheet-{sheet_index + 1:02d}.jpg"
        sheet.save(out, quality=88)

    print(json.dumps({"placements": len(records), "unique_images": len(unique), "sheets": math.ceil(len(unique)/per_sheet)}, ensure_ascii=False))


if __name__ == "__main__":
    main()
