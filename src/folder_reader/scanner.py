"""Recorrido del arbol de carpetas y filtrado de archivos del sistema."""

from __future__ import annotations

import os
import stat
import sys
from collections.abc import Iterator
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

ES_WINDOWS = sys.platform == "win32"

PREFIJO_LARGO = "\\\\?\\"
PREFIJO_LARGO_UNC = "\\\\?\\UNC\\"

# Archivos que el sistema operativo crea por su cuenta y que nunca interesan
# en un inventario documental.
ARCHIVOS_SISTEMA = frozenset(
    {
        "desktop.ini",
        "thumbs.db",
        "ehthumbs.db",
        "ehthumbs_vista.db",
        ".ds_store",
        "autorun.inf",
        "pagefile.sys",
        "hiberfil.sys",
        "swapfile.sys",
        "icon\r",
    }
)

# Prefijos de nombres de archivo del sistema u ocultos de Office
# (ntuser.dat.log1, ~$informe.docx, ...).
PREFIJOS_SISTEMA = ("ntuser.dat", "ntuser.ini", "~$")

CARPETAS_SISTEMA = frozenset(
    {
        "$recycle.bin",
        "system volume information",
        "$windows.~bt",
        "$windows.~ws",
        "$winreagent",
        "recycler",
        ".git",
        "__pycache__",
        ".venv",
    }
)

# Solo el atributo "oculto" sirve para descartar. El atributo "sistema" NO: Windows
# se lo pone a cualquier carpeta personalizada con desktop.ini (un icono propio basta)
# y OneDrive se lo pone a todo lo que sincroniza, incluido el Escritorio. Filtrar por
# el dejaria fuera carpetas de trabajo perfectamente normales. Lo verdaderamente
# intocable ($Recycle.Bin, System Volume Information) es oculto ademas de sistema.
_ATRIBUTO_OCULTO = getattr(stat, "FILE_ATTRIBUTE_HIDDEN", 0x2)


@dataclass(slots=True)
class ArchivoInfo:
    """Metadatos crudos de un archivo encontrado durante el recorrido."""

    ruta: Path
    ruta_visible: str
    ruta_relativa: str
    carpeta: str
    carpeta_nivel1: str
    nombre: str
    extension: str
    tamano_bytes: int
    creado: datetime | None
    modificado: datetime | None


@dataclass(slots=True)
class ResultadoRecorrido:
    """Contadores y advertencias acumulados durante el recorrido."""

    omitidos_sistema: int = 0
    omitidos_extension: int = 0
    carpetas_inaccesibles: list[str] = field(default_factory=list)
    archivos_inaccesibles: list[str] = field(default_factory=list)


def normalizar_raiz(ruta: str | os.PathLike[str]) -> Path:
    """Devuelve la ruta absoluta de la raiz, tolerando comillas y espacios sobrantes.

    Arrastrar una carpeta a la consola de Windows pega la ruta entre comillas.
    """
    texto = str(ruta).strip().strip('"').strip("'").strip()
    return Path(os.path.abspath(os.path.expanduser(texto)))


def con_prefijo_largo(ruta: Path | str) -> str:
    """Antepone el prefijo de ruta extendida en Windows (limite de 260 caracteres)."""
    texto = str(ruta)
    if not ES_WINDOWS or texto.startswith(PREFIJO_LARGO):
        return texto
    if texto.startswith("\\\\"):
        return PREFIJO_LARGO_UNC + texto[2:]
    return PREFIJO_LARGO + texto


def sin_prefijo_largo(texto: str) -> str:
    """Quita el prefijo de ruta extendida para mostrar la ruta que el usuario reconoce."""
    if texto.startswith(PREFIJO_LARGO_UNC):
        return "\\\\" + texto[len(PREFIJO_LARGO_UNC) :]
    if texto.startswith(PREFIJO_LARGO):
        return texto[len(PREFIJO_LARGO) :]
    return texto


def _es_oculto(nombre: str, st: os.stat_result | None) -> bool:
    """True si el elemento esta marcado como oculto."""
    if ES_WINDOWS:
        atributos = getattr(st, "st_file_attributes", 0) if st is not None else 0
        return bool(atributos & _ATRIBUTO_OCULTO)
    return nombre.startswith(".")


def es_archivo_sistema(nombre: str) -> bool:
    """True si el nombre corresponde a un archivo generado por el sistema operativo."""
    minuscula = nombre.lower()
    return minuscula in ARCHIVOS_SISTEMA or minuscula.startswith(PREFIJOS_SISTEMA)


def _fecha(marca: float | None) -> datetime | None:
    if marca is None:
        return None
    try:
        return datetime.fromtimestamp(marca)
    except (OverflowError, OSError, ValueError):
        return None


def _carpeta_relativa(carpeta: str, raiz: Path) -> str:
    r"""Ruta de la carpeta relativa a la raiz, incluyendo el nombre de la raiz.

    Para una raiz "MES 7" devuelve "MES 7" en el primer nivel y "MES 7\A" dentro
    de la subcarpeta A. Asi la columna identifica la ubicacion sin ambiguedad
    cuando varias carpetas se llaman igual, y sin arrastrar la ruta absoluta.
    """
    # Una raiz que es la propia unidad ("D:\\") no tiene nombre.
    etiqueta = raiz.name or str(raiz)
    relativa = os.path.relpath(carpeta, str(raiz))
    return etiqueta if relativa == os.curdir else os.path.join(etiqueta, relativa)


def _carpeta_nivel1(carpeta: str, raiz: Path) -> str:
    """Primera subcarpeta bajo la raiz, o cadena vacia si el archivo esta en la raiz.

    Para una raiz "MES 7", tanto "MES 7/A" como "MES 7/A/1.Ajuste" devuelven "A".
    Sirve para agrupar el inventario por la rama de primer nivel.
    """
    relativa = os.path.relpath(carpeta, str(raiz))
    if relativa == os.curdir:
        return ""
    return Path(relativa).parts[0]


def _stat_seguro(carpeta: str, nombre: str) -> os.stat_result | None:
    try:
        return os.stat(os.path.join(carpeta, nombre))
    except OSError:
        return None


def recorrer(
    raiz: Path,
    *,
    incluir_ocultos: bool = False,
    profundidad: int | None = None,
    extensiones: set[str] | None = None,
    resultado: ResultadoRecorrido | None = None,
) -> Iterator[ArchivoInfo]:
    """Recorre `raiz` y produce un `ArchivoInfo` por cada archivo que pase los filtros.

    Los errores de permisos no interrumpen el recorrido: se acumulan en `resultado`.
    """
    resultado = resultado if resultado is not None else ResultadoRecorrido()
    raiz = normalizar_raiz(raiz)
    raiz_walk = con_prefijo_largo(raiz)
    base_len = len(raiz_walk.rstrip("\\/"))

    def _al_fallar(error: OSError) -> None:
        resultado.carpetas_inaccesibles.append(sin_prefijo_largo(str(error.filename or "")))

    for carpeta_actual, subcarpetas, archivos in os.walk(
        raiz_walk, topdown=True, onerror=_al_fallar
    ):
        relativo = carpeta_actual[base_len:].replace("\\", "/").strip("/")
        nivel_actual = len(relativo.split("/")) if relativo else 0

        # profundidad=1 significa "solo la raiz", asi que se compara el nivel
        # que tendrian las subcarpetas, no el de la carpeta actual.
        if profundidad is not None and nivel_actual + 1 >= profundidad:
            subcarpetas[:] = []
        else:
            subcarpetas[:] = [
                d
                for d in subcarpetas
                if d.lower() not in CARPETAS_SISTEMA
                and (incluir_ocultos or not _es_oculto(d, _stat_seguro(carpeta_actual, d)))
            ]

        carpeta_visible = sin_prefijo_largo(carpeta_actual)
        # Se calcula una vez por carpeta, no una vez por archivo.
        carpeta_relativa = _carpeta_relativa(carpeta_visible, raiz)
        nivel1 = _carpeta_nivel1(carpeta_visible, raiz)

        for nombre in sorted(archivos):
            if es_archivo_sistema(nombre):
                resultado.omitidos_sistema += 1
                continue

            ruta_completa = os.path.join(carpeta_actual, nombre)
            try:
                st = os.stat(ruta_completa)
            except OSError:
                resultado.archivos_inaccesibles.append(sin_prefijo_largo(ruta_completa))
                continue

            if not incluir_ocultos and _es_oculto(nombre, st):
                resultado.omitidos_sistema += 1
                continue

            extension = os.path.splitext(nombre)[1].lower()
            if extensiones is not None and extension not in extensiones:
                resultado.omitidos_extension += 1
                continue

            ruta_visible = os.path.join(carpeta_visible, nombre)

            yield ArchivoInfo(
                ruta=Path(ruta_completa),
                ruta_visible=ruta_visible,
                ruta_relativa=os.path.relpath(ruta_visible, str(raiz)),
                carpeta=carpeta_relativa,
                carpeta_nivel1=nivel1,
                nombre=nombre,
                extension=extension,
                tamano_bytes=st.st_size,
                creado=_fecha(getattr(st, "st_birthtime", st.st_ctime)),
                modificado=_fecha(st.st_mtime),
            )
