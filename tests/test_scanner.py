from __future__ import annotations

import os
import subprocess
from pathlib import Path

import pytest

from folder_reader.scanner import (
    ES_WINDOWS,
    ResultadoRecorrido,
    es_archivo_sistema,
    recorrer,
)


def nombres(raiz: Path, **kwargs) -> set[str]:
    return {a.nombre for a in recorrer(raiz, **kwargs)}


def test_ignora_archivos_del_sistema(arbol: Path) -> None:
    encontrados = nombres(arbol)
    assert "desktop.ini" not in encontrados
    assert "Thumbs.db" not in encontrados
    assert "~$acta.docx" not in encontrados
    assert {"portada.pdf", "contrato.pdf", "anual.pdf", "acta.docx", "notas.txt"} <= encontrados


def test_cuenta_los_omitidos(arbol: Path) -> None:
    resultado = ResultadoRecorrido()
    list(recorrer(arbol, resultado=resultado))
    assert resultado.omitidos_sistema == 3


def test_recorre_subcarpetas_y_arma_ruta_relativa(arbol: Path) -> None:
    relativas = {a.ruta_relativa.replace("\\", "/") for a in recorrer(arbol)}
    assert "contratos/2025/contrato.pdf" in relativas
    assert "portada.pdf" in relativas


def test_carpeta_contenedora_es_la_ruta_relativa(arbol: Path) -> None:
    """Debe identificar la ubicacion aunque varias carpetas se llamen igual."""
    por_nombre = {a.nombre: a for a in recorrer(arbol)}
    assert por_nombre["contrato.pdf"].carpeta == os.path.join("arbol", "contratos", "2025")
    assert por_nombre["anual.pdf"].carpeta == os.path.join("arbol", "informes")
    assert por_nombre["portada.pdf"].carpeta == "arbol"


def test_limite_de_profundidad(arbol: Path) -> None:
    assert nombres(arbol, profundidad=1) == {"portada.pdf", "notas.txt"}
    assert "contrato.pdf" not in nombres(arbol, profundidad=2)
    assert "anual.pdf" in nombres(arbol, profundidad=2)


def test_filtro_de_extensiones(arbol: Path) -> None:
    assert nombres(arbol, extensiones={".docx"}) == {"acta.docx", "borrador.docx"}


def test_metadatos_basicos(arbol: Path) -> None:
    portada = next(a for a in recorrer(arbol) if a.nombre == "portada.pdf")
    assert portada.extension == ".pdf"
    assert portada.tamano_bytes > 0
    assert portada.creado is not None
    assert portada.modificado is not None
    assert Path(portada.ruta_visible).name == "portada.pdf"


def test_deteccion_de_nombres_del_sistema() -> None:
    assert es_archivo_sistema("DESKTOP.INI")
    assert es_archivo_sistema("NTUSER.DAT.LOG1")
    assert not es_archivo_sistema("informe.pdf")


@pytest.mark.skipif(not ES_WINDOWS, reason="los atributos de Windows no existen en POSIX")
def test_no_descarta_carpetas_marcadas_como_sistema(arbol: Path) -> None:
    """OneDrive y las carpetas personalizadas llevan el atributo SYSTEM sin ser ocultas."""
    subprocess.run(["attrib", "+S", str(arbol / "informes")], check=True, shell=True)

    encontrados = nombres(arbol)
    assert "anual.pdf" in encontrados, "una carpeta +S no puede desaparecer del inventario"


@pytest.mark.skipif(not ES_WINDOWS, reason="los atributos de Windows no existen en POSIX")
def test_si_descarta_carpetas_ocultas(arbol: Path) -> None:
    subprocess.run(["attrib", "+H", str(arbol / "informes")], check=True, shell=True)

    assert "anual.pdf" not in nombres(arbol)
    assert "anual.pdf" in nombres(arbol, incluir_ocultos=True)


def test_carpeta_nivel1(arbol: Path) -> None:
    """Agrupa por la rama de primer nivel, sin importar cuanto se anide despues."""
    por_nombre = {a.nombre: a for a in recorrer(arbol)}
    assert por_nombre["contrato.pdf"].carpeta_nivel1 == "contratos"
    assert por_nombre["anual.pdf"].carpeta_nivel1 == "informes"
    assert por_nombre["portada.pdf"].carpeta_nivel1 == ""
