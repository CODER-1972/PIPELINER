from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from xml.sax.saxutils import escape

RUN_DIR = Path(__file__).resolve().parent


def make_docx_from_txt(txt_path: Path, docx_path: Path) -> None:
    text = txt_path.read_text(encoding="utf-8")
    paragraphs = text.splitlines()

    body_parts = []
    for line in paragraphs:
        if line == "":
            body_parts.append("<w:p/>")
        else:
            body_parts.append(
                '<w:p><w:r><w:t xml:space="preserve">' + escape(line) + "</w:t></w:r></w:p>"
            )

    body_xml = "".join(body_parts)

    content_types = """<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">
  <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>
  <Default Extension=\"xml\" ContentType=\"application/xml\"/>
  <Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>
</Types>"""

    rels = """<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">
  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>
</Relationships>"""

    doc_rels = """<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"/>"""

    document = f"""<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<w:document xmlns:wpc=\"http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas\"
 xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"
 xmlns:o=\"urn:schemas-microsoft-com:office:office\"
 xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"
 xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"
 xmlns:v=\"urn:schemas-microsoft-com:vml\"
 xmlns:wp14=\"http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing\"
 xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\"
 xmlns:w10=\"urn:schemas-microsoft-com:office:word\"
 xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"
 xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\"
 xmlns:wpg=\"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup\"
 xmlns:wpi=\"http://schemas.microsoft.com/office/word/2010/wordprocessingInk\"
 xmlns:wne=\"http://schemas.microsoft.com/office/word/2006/wordml\"
 xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\"
 mc:Ignorable=\"w14 wp14\">
  <w:body>
    {body_xml}
    <w:sectPr>
      <w:pgSz w:w=\"11906\" w:h=\"16838\"/>
      <w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:header=\"708\" w:footer=\"708\" w:gutter=\"0\"/>
      <w:cols w:space=\"708\"/>
      <w:docGrid w:linePitch=\"360\"/>
    </w:sectPr>
  </w:body>
</w:document>"""

    with ZipFile(docx_path, "w", ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("word/document.xml", document)
        zf.writestr("word/_rels/document.xml.rels", doc_rels)


if __name__ == "__main__":
    for base in ["ANALYSIS_REPORT", "PROMPT_REVISAO"]:
        make_docx_from_txt(RUN_DIR / f"{base}.txt", RUN_DIR / f"{base}.docx")
        print(f"Gerado: {base}.docx")
