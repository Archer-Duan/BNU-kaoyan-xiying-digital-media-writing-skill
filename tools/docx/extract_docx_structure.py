from pathlib import Path
from zipfile import ZipFile

from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph


SOURCE = Path(r"C:\Users\Duan\Downloads\Untitled.docx")


def iter_block_items(parent):
    body = parent.element.body
    for child in body.iterchildren():
        if child.tag.endswith("}p"):
            yield Paragraph(child, parent)
        elif child.tag.endswith("}tbl"):
            yield Table(child, parent)


doc = Document(SOURCE)
print(f"PARAGRAPHS={len(doc.paragraphs)} TABLES={len(doc.tables)} SECTIONS={len(doc.sections)}")

for index, block in enumerate(iter_block_items(doc), start=1):
    if isinstance(block, Paragraph):
        text = block.text.strip()
        drawings = len(block._p.xpath('.//*[local-name()="drawing"]'))
        if text or drawings:
            style = block.style.name if block.style is not None else ""
            marker = f" [DRAWINGS={drawings}]" if drawings else ""
            print(f"P{index:04d} [{style}]{marker} {text}")
    else:
        print(f"T{index:04d} ROWS={len(block.rows)} COLS={len(block.columns)}")
        for row_index, row in enumerate(block.rows, start=1):
            cells = [" ".join(cell.text.split()) for cell in row.cells]
            print(f"  R{row_index:03d}: " + " | ".join(cells))

with ZipFile(SOURCE) as archive:
    media = [name for name in archive.namelist() if name.startswith("word/media/")]
    print(f"MEDIA={len(media)}")
    for name in media:
        info = archive.getinfo(name)
        print(f"  {name} {info.file_size} bytes")
