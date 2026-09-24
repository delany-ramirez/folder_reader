"""Fixtures que construyen un arbol de prueba con documentos reales."""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest
from pypdf import PdfWriter

APP_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Pages>{paginas}</Pages>
  <Words>120</Words>
</Properties>
"""

CORE_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties
    xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
    xmlns:dcterms="http://purl.org/dc/terms/">
  <dcterms:modified>{guardado}</dcterms:modified>
</cp:coreProperties>
"""

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
</Types>
"""


def crear_pdf(destino: Path, paginas: int) -> Path:
    escritor = PdfWriter()
    for _ in range(paginas):
        escritor.add_blank_page(width=612, height=792)
    with destino.open("wb") as fh:
        escritor.write(fh)
    return destino


def crear_docx(destino: Path, paginas: int | None, guardado: str | None = None) -> Path:
    """Genera un .docx minimo; `paginas=None` omite docProps/app.xml."""
    with zipfile.ZipFile(destino, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", CONTENT_TYPES)
        zf.writestr("word/document.xml", "<document/>")
        if paginas is not None:
            zf.writestr("docProps/app.xml", APP_XML.format(paginas=paginas))
        if guardado is not None:
            zf.writestr("docProps/core.xml", CORE_XML.format(guardado=guardado))
    return destino


@pytest.fixture
def arbol(tmp_path: Path) -> Path:
    """Arbol con subcarpetas, documentos validos, basura del sistema y un PDF roto."""
    raiz = tmp_path / "arbol"
    contratos = raiz / "contratos" / "2025"
    informes = raiz / "informes"
    contratos.mkdir(parents=True)
    informes.mkdir(parents=True)

    crear_pdf(raiz / "portada.pdf", 1)
    crear_pdf(contratos / "contrato.pdf", 3)
    crear_pdf(informes / "anual.pdf", 10)
    crear_docx(informes / "acta.docx", 5, guardado="2025-03-04T18:30:00Z")
    crear_docx(informes / "borrador.docx", None)
    (raiz / "notas.txt").write_text("texto plano", encoding="utf-8")
    (informes / "roto.pdf").write_bytes(b"%PDF-1.4 esto no es un pdf valido")

    # Basura que debe ignorarse.
    (raiz / "desktop.ini").write_text("[.ShellClassInfo]", encoding="utf-8")
    (contratos / "Thumbs.db").write_bytes(b"\x00\x01")
    (informes / "~$acta.docx").write_bytes(b"\x00")

    return raiz
