from __future__ import annotations

from pathlib import Path

import pytest

from folder_reader.pages import (
    ESTADO_METADATO,
    ESTADO_NO_APLICA,
    ESTADO_OK,
    contar_paginas,
)


class WordFalso:
    """Sustituto de WordCounter que no necesita Word instalado."""

    def __init__(self, paginas: int | None, estado: str = ESTADO_OK) -> None:
        self._paginas = paginas
        self._estado = estado
        self.abiertos: list[str] = []

    def contar(self, ruta: Path) -> tuple[int | None, str]:
        self.abiertos.append(ruta.name)
        return self._paginas, self._estado


@pytest.mark.parametrize(
    ("relativa", "esperado"),
    [
        ("portada.pdf", 1),
        ("contratos/2025/contrato.pdf", 3),
        ("informes/anual.pdf", 10),
    ],
)
def test_cuenta_paginas_de_pdf(arbol: Path, relativa: str, esperado: int) -> None:
    ruta = arbol / relativa
    assert contar_paginas(ruta, ".pdf") == (esperado, ESTADO_OK)


def test_docx_se_abre_en_word_aunque_haya_metadato(arbol: Path) -> None:
    """Word manda: el metadato puede estar desactualizado respecto al contenido."""
    ruta = arbol / "informes" / "acta.docx"  # su metadato dice 5 paginas
    word = WordFalso(7)

    assert contar_paginas(ruta, ".docx", word=word) == (7, ESTADO_OK)
    assert word.abiertos == ["acta.docx"]


def test_docx_cae_al_metadato_si_word_falla(arbol: Path) -> None:
    ruta = arbol / "informes" / "acta.docx"
    word = WordFalso(None, "Word no pudo abrir el documento: bloqueado")

    paginas, estado = contar_paginas(ruta, ".docx", word=word)
    assert paginas == 5
    assert ESTADO_METADATO in estado
    assert "bloqueado" in estado


def test_docx_sin_word_usa_el_metadato_y_lo_declara(arbol: Path) -> None:
    ruta = arbol / "informes" / "acta.docx"
    assert contar_paginas(ruta, ".docx", word=None) == (5, ESTADO_METADATO)


def test_docx_sin_metadato_y_sin_word(arbol: Path) -> None:
    ruta = arbol / "informes" / "borrador.docx"
    paginas, estado = contar_paginas(ruta, ".docx", word=None)
    assert paginas is None
    assert "Sin metadatos" in estado


def test_docx_sin_metadato_lo_resuelve_word(arbol: Path) -> None:
    ruta = arbol / "informes" / "borrador.docx"
    assert contar_paginas(ruta, ".docx", word=WordFalso(3)) == (3, ESTADO_OK)


def test_doc_prefiere_el_metadato_ole_y_no_abre_word(tmp_path: Path) -> None:
    """El .doc trae el contador en SummaryInformation: no hay que pagar Word."""
    falso = tmp_path / "falso.doc"
    falso.write_text("no soy un documento binario de Word", encoding="utf-8")
    word = WordFalso(9)

    paginas, estado = contar_paginas(falso, ".doc", word=word)
    # Este archivo no es OLE2, asi que si cae en Word: comprueba el encadenado.
    assert (paginas, estado) == (9, ESTADO_OK)
    assert word.abiertos == ["falso.doc"]


def test_doc_que_no_es_ole_sin_word(tmp_path: Path) -> None:
    falso = tmp_path / "falso.doc"
    falso.write_text("no soy un documento binario de Word", encoding="utf-8")
    paginas, estado = contar_paginas(falso, ".doc", word=None)
    assert paginas is None
    assert "OLE2" in estado


def test_pdf_corrupto_no_interrumpe(arbol: Path) -> None:
    ruta = arbol / "informes" / "roto.pdf"
    paginas, estado = contar_paginas(ruta, ".pdf")
    assert paginas is None
    assert estado not in {ESTADO_OK, ESTADO_NO_APLICA}


def test_extension_sin_paginas(arbol: Path) -> None:
    assert contar_paginas(arbol / "notas.txt", ".txt") == (None, ESTADO_NO_APLICA)


def test_archivo_inexistente(tmp_path: Path) -> None:
    paginas, estado = contar_paginas(tmp_path / "fantasma.pdf", ".pdf")
    assert paginas is None
    assert estado.startswith("Error al leer")


class _DocumentsQueFallan:
    def Open(self, *args, **kwargs):  # noqa: N802 - nombre impuesto por la API de Word
        raise RuntimeError("formato bloqueado")


class _WordQueFalla:
    Documents = _DocumentsQueFallan()

    def Quit(self) -> None:  # noqa: N802 - nombre impuesto por la API de Word
        pass


def test_word_se_rinde_tras_varios_fallos_seguidos(tmp_path: Path) -> None:
    """Si Word falla siempre, no tiene sentido pagar el intento en cada archivo."""
    from folder_reader.pages import _FALLOS_SEGUIDOS_PARA_RENDIRSE, WordCounter

    contador = WordCounter()
    contador._word = _WordQueFalla()
    ruta = tmp_path / "cualquiera.docx"

    for _ in range(_FALLOS_SEGUIDOS_PARA_RENDIRSE):
        paginas, estado = contador.contar(ruta)
        assert paginas is None
        assert "formato bloqueado" in estado

    paginas, estado = contador.contar(ruta)
    assert paginas is None
    assert estado.startswith("Word desactivado tras varios fallos")


def test_fecha_guardado_convierte_utc_a_hora_local(arbol: Path) -> None:
    """dcterms:modified viene en UTC; Windows muestra 'Guardado el' en hora local."""
    from datetime import datetime, timezone

    from folder_reader.pages import fecha_guardado

    obtenida = fecha_guardado(arbol / "informes" / "acta.docx", ".docx")
    esperada = (
        datetime(2025, 3, 4, 18, 30, tzinfo=timezone.utc).astimezone().replace(tzinfo=None)
    )
    assert obtenida == esperada
    assert obtenida.tzinfo is None, "Excel no admite fechas con zona horaria"


def test_fecha_guardado_vacia_si_no_hay_metadato(arbol: Path) -> None:
    from folder_reader.pages import fecha_guardado

    assert fecha_guardado(arbol / "informes" / "borrador.docx", ".docx") is None
    assert fecha_guardado(arbol / "notas.txt", ".txt") is None


def test_fecha_guardado_no_falla_con_archivo_corrupto(arbol: Path) -> None:
    from folder_reader.pages import fecha_guardado

    assert fecha_guardado(arbol / "informes" / "roto.pdf", ".pdf") is None


@pytest.mark.parametrize("extension", [".docx", ".xlsx", ".pptx", ".xlsm"])
def test_fecha_guardado_cubre_excel_y_powerpoint(tmp_path: Path, extension: str) -> None:
    """El Explorador muestra 'Guardado el' para todo Office, no solo para Word."""
    from datetime import datetime, timezone

    from folder_reader.pages import fecha_guardado

    from conftest import crear_docx  # mismo contenedor OOXML para los tres

    destino = crear_docx(tmp_path / f"hoja{extension}", None, guardado="2024-06-01T10:00:00Z")
    esperada = (
        datetime(2024, 6, 1, 10, 0, tzinfo=timezone.utc).astimezone().replace(tzinfo=None)
    )
    assert fecha_guardado(destino, extension) == esperada
