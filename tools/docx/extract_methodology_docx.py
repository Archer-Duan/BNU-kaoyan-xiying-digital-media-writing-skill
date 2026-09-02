from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from zipfile import ZipFile

from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph
from lxml import etree


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}


def iter_blocks(document: Document):
    body = document.element.body
    for child in body.iterchildren():
        if child.tag == f"{{{W_NS}}}p":
            yield "paragraph", Paragraph(child, document)
        elif child.tag == f"{{{W_NS}}}tbl":
            yield "table", Table(child, document)


def all_text(element) -> str:
    parts = []
    for node in element.xpath('.//*[local-name()="t"]'):
        if node.text:
            parts.append(node.text)
    return "".join(parts).strip()


def normalize(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def read_ooxml_part(archive: ZipFile, name: str) -> list[str]:
    if name not in archive.namelist():
        return []
    root = etree.fromstring(archive.read(name))
    paragraphs = []
    for p in root.xpath(".//w:p", namespaces=NS):
        text = "".join(p.xpath(".//w:t/text()", namespaces=NS))
        text = normalize(text)
        if text:
            paragraphs.append(text)
    return paragraphs


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("source", type=Path)
    parser.add_argument("output_dir", type=Path)
    args = parser.parse_args()

    args.output_dir.mkdir(parents=True, exist_ok=True)
    media_dir = args.output_dir / "media"
    media_dir.mkdir(exist_ok=True)

    doc = Document(args.source)
    lines: list[str] = []
    style_counts: dict[str, int] = {}
    paragraph_count = 0
    table_count = 0
    drawing_count = 0

    for block_index, (kind, block) in enumerate(iter_blocks(doc), start=1):
        if kind == "paragraph":
            paragraph_count += 1
            style = block.style.name if block.style is not None else ""
            style_counts[style] = style_counts.get(style, 0) + 1
            text = normalize(all_text(block._p))
            drawings = len(block._p.xpath('.//*[local-name()="drawing"]'))
            drawing_count += drawings
            if text or drawings:
                marker = f" drawings={drawings}" if drawings else ""
                lines.append(f"P{block_index:05d} style={json.dumps(style, ensure_ascii=False)}{marker}\n{text}")
        else:
            table_count += 1
            lines.append(
                f"T{block_index:05d} rows={len(block.rows)} cols={len(block.columns)}"
            )
            for row_index, row in enumerate(block.rows, start=1):
                cells = [normalize(all_text(cell._tc)) for cell in row.cells]
                lines.append(f"  R{row_index:04d}: " + " | ".join(cells))

    with ZipFile(args.source) as archive:
        names = archive.namelist()
        media_names = [name for name in names if name.startswith("word/media/")]
        for name in media_names:
            target = media_dir / Path(name).name
            target.write_bytes(archive.read(name))

        comments = read_ooxml_part(archive, "word/comments.xml")
        footnotes = read_ooxml_part(archive, "word/footnotes.xml")
        endnotes = read_ooxml_part(archive, "word/endnotes.xml")

        app_props = {}
        if "docProps/app.xml" in names:
            app_root = etree.fromstring(archive.read("docProps/app.xml"))
            for node in app_root:
                local = etree.QName(node).localname
                if local in {"Pages", "Words", "Characters", "Paragraphs", "Lines"}:
                    app_props[local] = node.text

        document_xml = archive.read("word/document.xml")
        last_rendered_breaks = document_xml.count(b"lastRenderedPageBreak")
        explicit_page_breaks = document_xml.count(b'w:type="page"')

    (args.output_dir / "body.txt").write_text("\n\n".join(lines), encoding="utf-8")
    (args.output_dir / "notes.txt").write_text(
        "COMMENTS\n" + "\n".join(comments)
        + "\n\nFOOTNOTES\n" + "\n".join(footnotes)
        + "\n\nENDNOTES\n" + "\n".join(endnotes),
        encoding="utf-8",
    )

    summary = {
        "source": str(args.source),
        "size_bytes": args.source.stat().st_size,
        "paragraphs": paragraph_count,
        "tables": table_count,
        "sections": len(doc.sections),
        "drawings": drawing_count,
        "media": len(media_names),
        "comments": len(comments),
        "footnotes": len(footnotes),
        "endnotes": len(endnotes),
        "style_counts": style_counts,
        "app_properties": app_props,
        "last_rendered_page_breaks": last_rendered_breaks,
        "explicit_page_breaks": explicit_page_breaks,
    }
    (args.output_dir / "summary.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
