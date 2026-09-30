"""Descompresion de los .zip y .rar que estan en la raiz de la carpeta escaneada.

Solo se descomprimen los archivos del primer nivel: un comprimido dentro de una
subcarpeta (o dentro de otro comprimido) se inventaria como un archivo mas. Cada
comprimido se extrae en una carpeta hermana con su mismo nombre, de modo que el
recorrido posterior lo inventaria como cualquier otra subcarpeta de la raiz.
"""

from __future__ import annotations

import os
import shutil
import subprocess
import time
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

from folder_reader.scanner import (
    ES_WINDOWS,
    _es_oculto,
    con_prefijo_largo,
    es_archivo_sistema,
)

EXTENSIONES_COMPRIMIDAS = frozenset({".zip", ".rar"})

# La extraccion se hace en una carpeta temporal con este sufijo y se renombra al
# terminar. Asi una corrida interrumpida nunca deja una carpeta a medias con el
# nombre definitivo, que la siguiente corrida tomaria por ya descomprimida.
_SUFIJO_TEMPORAL = ".descomprimiendo~"

# Rutas habituales de las herramientas externas cuando no estan en el PATH.
_RUTAS_WINRAR = (
    Path(os.environ.get("ProgramFiles", r"C:\Program Files")) / "WinRAR",
    Path(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")) / "WinRAR",
)
_RUTAS_7ZIP = (
    Path(os.environ.get("ProgramFiles", r"C:\Program Files")) / "7-Zip",
    Path(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")) / "7-Zip",
)


class ErrorDescompresion(Exception):
    """El comprimido no se pudo extraer; el mensaje es legible para el usuario."""


@dataclass(slots=True)
class ResultadoDescompresion:
    """Que paso con cada comprimido de la raiz."""

    descomprimidos: list[str] = field(default_factory=list)
    ya_existian: list[str] = field(default_factory=list)
    fallidos: list[tuple[str, str]] = field(default_factory=list)


def comprimidos_en_raiz(raiz: Path, *, incluir_ocultos: bool = False) -> list[Path]:
    """Lista los .zip y .rar del primer nivel de `raiz`, ordenados por nombre."""
    encontrados = []
    try:
        entradas = list(os.scandir(raiz))
    except OSError:
        return []
    for entrada in entradas:
        if Path(entrada.name).suffix.lower() not in EXTENSIONES_COMPRIMIDAS:
            continue
        if es_archivo_sistema(entrada.name):
            continue
        try:
            if not entrada.is_file():
                continue
            st = entrada.stat()
        except OSError:
            continue
        if not incluir_ocultos and _es_oculto(entrada.name, st):
            continue
        encontrados.append(Path(entrada.path))
    return sorted(encontrados, key=lambda p: p.name.lower())


def descomprimir_en_raiz(
    raiz: Path,
    *,
    incluir_ocultos: bool = False,
    resultado: ResultadoDescompresion | None = None,
) -> ResultadoDescompresion:
    """Extrae cada comprimido de la raiz en una carpeta con su mismo nombre.

    Si esa carpeta ya existe se da por descomprimido y no se toca: volver a
    correr el inventario no duplica ni pisa nada. Ningun fallo interrumpe el
    proceso; los motivos quedan en `resultado.fallidos`.
    """
    resultado = resultado if resultado is not None else ResultadoDescompresion()
    for archivo in comprimidos_en_raiz(raiz, incluir_ocultos=incluir_ocultos):
        destino = archivo.with_suffix("")
        if destino.exists():
            resultado.ya_existian.append(archivo.name)
            continue
        try:
            descomprimir(archivo, destino)
        except ErrorDescompresion as error:
            resultado.fallidos.append((archivo.name, str(error)))
        else:
            resultado.descomprimidos.append(archivo.name)
    return resultado


def descomprimir(archivo: Path, destino: Path) -> None:
    """Extrae `archivo` en `destino`, que no debe existir todavia."""
    temporal = destino.with_name(destino.name + _SUFIJO_TEMPORAL)
    if temporal.exists():
        # Restos de una corrida anterior interrumpida.
        shutil.rmtree(con_prefijo_largo(temporal), ignore_errors=True)
    temporal.mkdir()

    try:
        if archivo.suffix.lower() == ".zip":
            _extraer_zip(archivo, temporal)
        else:
            _extraer_con_herramienta(archivo, temporal)
        _aplanar_carpeta_unica(temporal, destino.name)
        temporal.rename(destino)
    except BaseException:
        shutil.rmtree(con_prefijo_largo(temporal), ignore_errors=True)
        raise


def _codificacion_oem() -> str:
    """Codificacion con la que el Explorador de Windows escribe nombres en un .zip.

    Los .zip sin la marca UTF-8 guardan los nombres en la pagina de codigos OEM
    del equipo que los creo (cp850 en Windows en espanol). Leerlos como cp437,
    lo que hace Python por defecto, estropea las tildes y las enes.
    """
    if ES_WINDOWS:
        import ctypes

        return f"cp{ctypes.windll.kernel32.GetOEMCP()}"
    return "cp437"


def _extraer_zip(archivo: Path, temporal: Path) -> None:
    try:
        with zipfile.ZipFile(archivo, metadata_encoding=_codificacion_oem()) as zf:
            if any(info.flag_bits & 0x1 for info in zf.infolist()):
                raise ErrorDescompresion("Protegido con contrasena")
            zf.extractall(con_prefijo_largo(temporal))
            _restaurar_fechas_zip(zf, temporal)
    except zipfile.BadZipFile as error:
        raise ErrorDescompresion(f"Comprimido danado o no es un .zip ({error})") from error
    except NotImplementedError:
        # Metodos que zipfile no conoce, como Deflate64, que el Explorador usa
        # en archivos grandes. Las herramientas externas si los manejan.
        shutil.rmtree(con_prefijo_largo(temporal))
        temporal.mkdir()
        _extraer_con_herramienta(archivo, temporal)
    except OSError as error:
        raise ErrorDescompresion(f"No se pudo extraer: {error}") from error


def _restaurar_fechas_zip(zf: zipfile.ZipFile, temporal: Path) -> None:
    """Devuelve a cada archivo la fecha de modificacion que trae el .zip.

    `extractall` deja la hora de la extraccion, y el inventario perderia la
    fecha real de ultima modificacion de todo lo que venia comprimido.
    """
    for info in zf.infolist():
        if info.is_dir():
            continue
        try:
            marca = time.mktime((*info.date_time, 0, 0, -1))
        except (OverflowError, ValueError):
            continue
        ruta = temporal.joinpath(*info.filename.split("/"))
        try:
            os.utime(con_prefijo_largo(ruta), (marca, marca))
        except OSError:
            pass


def _herramientas(archivo: Path, temporal: Path) -> list[list[str]]:
    """Comandos disponibles para extraer `archivo`, en orden de preferencia."""
    comandos: list[list[str]] = []
    es_rar = archivo.suffix.lower() == ".rar"
    salida = str(temporal) + os.sep

    if es_rar:
        # -p- no pregunta contrasena: si la pide, falla en vez de quedarse esperando.
        for nombre in ("UnRAR.exe", "Rar.exe", "unrar", "rar"):
            ejecutable = _buscar(nombre, _RUTAS_WINRAR)
            if ejecutable:
                comandos.append([ejecutable, "x", "-y", "-p-", "-idq", str(archivo), salida])
                break

    siete = _buscar("7z.exe", _RUTAS_7ZIP) or _buscar("7z", ())
    if siete:
        comandos.append([siete, "x", "-y", "-p", "-bso0", "-bsp0", f"-o{temporal}", str(archivo)])

    # El tar de Windows es bsdtar (libarchive) y lee .rar y .zip. El de Git para
    # Windows es GNU tar y no sirve, por eso se prefiere el de System32.
    tar_sistema = Path(os.environ.get("SystemRoot", r"C:\Windows")) / "System32" / "tar.exe"
    tar = str(tar_sistema) if tar_sistema.is_file() else shutil.which("bsdtar")
    if tar:
        comandos.append([tar, "-xf", str(archivo), "-C", str(temporal)])

    return comandos


def _buscar(nombre: str, carpetas: tuple[Path, ...]) -> str | None:
    encontrado = shutil.which(nombre)
    if encontrado:
        return encontrado
    for carpeta in carpetas:
        candidato = carpeta / nombre
        if candidato.is_file():
            return str(candidato)
    return None


def _extraer_con_herramienta(archivo: Path, temporal: Path) -> None:
    comandos = _herramientas(archivo, temporal)
    if not comandos:
        raise ErrorDescompresion(
            "No hay con que descomprimirlo: instale WinRAR o 7-Zip"
        )

    ultimo_error = ""
    for comando in comandos:
        try:
            proceso = subprocess.run(
                comando,
                stdin=subprocess.DEVNULL,
                capture_output=True,
                text=True,
                errors="replace",
            )
        except OSError as error:
            ultimo_error = str(error)
            continue
        if proceso.returncode == 0:
            return
        detalle = (proceso.stderr or proceso.stdout).strip().splitlines()
        ultimo_error = detalle[-1] if detalle else f"codigo de salida {proceso.returncode}"
        # Cada herramienta pudo dejar archivos a medias antes de fallar.
        shutil.rmtree(con_prefijo_largo(temporal), ignore_errors=True)
        temporal.mkdir()

    raise ErrorDescompresion(f"No se pudo extraer ({ultimo_error})")


def _aplanar_carpeta_unica(temporal: Path, nombre: str) -> None:
    """Evita la carpeta duplicada `informe\\informe\\...`.

    Es muy comun comprimir una carpeta entera, de modo que el comprimido trae
    una sola carpeta con su mismo nombre. En ese caso se sube su contenido un
    nivel, que es lo que hace "Extraer aqui" en el Explorador.
    """
    entradas = list(temporal.iterdir())
    if len(entradas) != 1 or not entradas[0].is_dir():
        return
    unica = entradas[0]
    if unica.name.lower() != nombre.lower():
        return
    hijos = list(unica.iterdir())
    if any(h.name.lower() == unica.name.lower() for h in hijos):
        # Subir un hijo con el mismo nombre chocaria con la propia carpeta.
        return
    for hijo in hijos:
        hijo.rename(temporal / hijo.name)
    unica.rmdir()
