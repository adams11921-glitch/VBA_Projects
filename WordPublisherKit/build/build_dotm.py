"""Builds dist/WordPublisherKit.dotm: a macro-enabled Word template that carries
the Bulletin / Publisher ribbon (ribbon/customUI14.xml).

The VBA modules in src/ are added afterwards on Windows, either with
build/Add-Macros.ps1 or by importing them by hand (see README.md), because a
VBA project can only be created by Office itself.

Usage:  python3 build/build_dotm.py
"""
import os
import zipfile

KIT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(KIT, "dist", "WordPublisherKit.dotm")

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.ms-word.template.macroEnabledTemplate.main+xml"/>
</Types>"""

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" Target="customUI/customUI14.xml"/>
</Relationships>"""

LINES = [
    ("Word Publisher Kit", True),
    ("This file is a Word add-in. It adds a Bulletin tab and a Publisher tab to the Word ribbon.", False),
    ("To install it, close Word and run Install.cmd from the same folder.", False),
    ("Do not type in this file. Use the Bulletin tab > New Bulletin to start a bulletin.", False),
]


def paragraph(text, bold):
    run_props = "<w:rPr><w:b/><w:sz w:val=\"32\"/></w:rPr>" if bold else ""
    return f"<w:p><w:r>{run_props}<w:t xml:space=\"preserve\">{text}</w:t></w:r></w:p>"


def document_xml():
    body = "".join(paragraph(t, b) for t, b in LINES)
    return ("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">"
            f"<w:body>{body}<w:sectPr/></w:body></w:document>")


def main():
    with open(os.path.join(KIT, "ribbon", "customUI14.xml"), encoding="utf-8") as f:
        ribbon = f.read()
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("word/document.xml", document_xml())
        z.writestr("customUI/customUI14.xml", ribbon)
    print("Wrote", OUT)


if __name__ == "__main__":
    main()
