from __future__ import annotations

import io
import os

import pytest

from folder_reader import progreso
from folder_reader.progreso import Progreso


class _Consola(io.StringIO):
    def isatty(self) -> bool:
        return True


def test_en_consola_reescribe_la_linea_con_porcentaje_y_archivo(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(progreso, "_INTERVALO", 0)
    consola = _Consola()
    barra = Progreso(4, salida=consola)

    barra.avanzar(1, "OS507\informe.pdf")
    barra.terminar()

    texto = consola.getvalue()
    assert "\r  [ 25%] 1/4  OS507\informe.pdf" in texto
    assert "[100%] 4/4 en" in texto
    assert texto.endswith("\n")


def test_redirigido_escribe_solo_una_linea_por_decena() -> None:
    archivo = io.StringIO()
    barra = Progreso(100, salida=archivo)

    for hechos in range(100):
        barra.avanzar(hechos, "x.pdf")
    barra.terminar()

    lineas = archivo.getvalue().splitlines()
    assert lineas[0] == "  0% (0/100)"
    assert lineas[-1] == "  100% (100/100)"
    assert len(lineas) == 11
    assert "\r" not in archivo.getvalue()


def test_recorta_las_rutas_largas_al_ancho_de_la_consola(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(progreso, "_INTERVALO", 0)
    monkeypatch.setattr(progreso.shutil, "get_terminal_size", lambda *_: os.terminal_size((60, 20)))
    consola = _Consola()

    Progreso(10, salida=consola).avanzar(3, "carpeta\\" * 20 + "final.pdf")

    linea = consola.getvalue().lstrip("\r")
    assert len(linea) == 59
    assert linea.startswith("  [ 30%] 3/10")
    assert linea.endswith("final.pdf")
