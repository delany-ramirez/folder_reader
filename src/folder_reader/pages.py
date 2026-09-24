"""Conteo de paginas para PDF y documentos de Word."""

from __future__ import annotations

import logging
import sys
import warnings
import xml.etree.ElementTree as ET
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from types import TracebackType

from folder_reader.scanner import sin_prefijo_largo

ES_WINDOWS = sys.platform == "win32"

EXTENSIONES_OOXML = frozenset({".docx", ".docm", ".dotx"})
EXTENSIONES_CON_PAGINAS = frozenset({".pdf", ".doc", ".dot"}) | EXTENSIONES_OOXML

# Para la fecha de guardado el alcance es mayor: solo Word pagina documentos, pero
# Excel y PowerPoint guardan la fecha en los mismos contenedores (zip OOXML el
# formato nuevo, OLE2 el heredado), y el Explorador la muestra para todos ellos.
EXTENSIONES_OOXML_TODAS = EXTENSIONES_OOXML | frozenset(
    {
        ".dotm",
        ".xlsx",
        ".xlsm",
        ".xltx",
        ".xltm",
        ".pptx",
        ".pptm",
        ".ppsx",
        ".ppsm",
        ".potx",
        ".potm",
    }
)
EXTENSIONES_OLE = frozenset({".doc", ".dot", ".xls", ".xlt", ".ppt", ".pot", ".pps"})

# Si Word falla tantas veces seguidas, se deja de intentar: normalmente significa
# que el formato esta bloqueado o que Word murio, y reintentarlo en cada archivo
# solo hace lento el inventario sin aportar nada.
_FALLOS_SEGUIDOS_PARA_RENDIRSE = 5

# Namespaces de los metadatos OOXML: app.xml lleva el numero de paginas y
# core.xml la fecha de ultimo guardado.
_NS_EXTENDIDO = "{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}"
_NS_TERMINOS = "{http://purl.org/dc/terms/}"

# Constante wdStatisticPages de la API de automatizacion de Word.
_WD_STATISTIC_PAGES = 2

# Contrasena falsa: si el documento esta protegido, Word lanza una excepcion
# en vez de quedarse esperando en un dialogo modal que nadie puede cerrar.
_PASSWORD_CENTINELA = "__sin_password__"

ESTADO_OK = "OK"
ESTADO_NO_APLICA = "No aplica"
# El conteo no se calculo ahora: es el que el programa dejo guardado al guardar
# el archivo, y puede no corresponder al contenido actual.
ESTADO_METADATO = "OK (metadato guardado, sin abrir en Word)"


def _detalle_com(error: Exception) -> str:
    """Extrae el mensaje que Word puso dentro del com_error.

    Sin esto el usuario solo ve "com_error" y no puede saber si el problema es
    un documento protegido o el bloqueo de archivos heredados del Centro de
    confianza, que es una configuracion suya y tiene arreglo.
    """
    info = getattr(error, "excepinfo", None)
    if isinstance(info, tuple) and len(info) > 2 and info[2]:
        return " ".join(str(info[2]).split())[:200]
    return f"{type(error).__name__}: {error}"[:200]


class WordCounter:
    """Instancia unica de Word reutilizada para contar paginas de documentos.

    Arrancar Word una vez por archivo es ordenes de magnitud mas lento, asi que la
    aplicacion se abre al entrar al contexto y se cierra al salir, pase lo que pase.
    """

    def __init__(self) -> None:
        self._word = None
        self._pythoncom = None
        self._fallos_seguidos = 0
        self._no_disponible: str | None = (
            None if ES_WINDOWS else "Word solo esta disponible en Windows"
        )

    def __enter__(self) -> WordCounter:
        return self

    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc: BaseException | None,
        tb: TracebackType | None,
    ) -> None:
        self.cerrar()

    def _asegurar_word(self):
        if self._word is not None or self._no_disponible is not None:
            return self._word
        try:
            import pythoncom  # type: ignore[import-not-found]
            import win32com.client  # type: ignore[import-not-found]

            pythoncom.CoInitialize()
            self._pythoncom = pythoncom
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            self._word = word
        except Exception as error:  # pragma: no cover - depende de Word instalado
            self._no_disponible = f"Word no disponible ({type(error).__name__})"
        return self._word

    def contar(self, ruta: Path) -> tuple[int | None, str]:
        """Abre el documento en Word y devuelve su numero real de paginas."""
        word = self._asegurar_word()
        if word is None:
            return None, self._no_disponible or "Word no disponible"

        # La API COM de Word no entiende el prefijo de ruta extendida: si se lo
        # pasamos, Open devuelve None en vez de lanzar un error.
        ruta_word = sin_prefijo_largo(str(ruta))

        documento = None
        try:
            documento = word.Documents.Open(
                ruta_word,
                ConfirmConversions=False,
                ReadOnly=True,
                AddToRecentFiles=False,
                PasswordDocument=_PASSWORD_CENTINELA,
                Visible=False,
            )
            if documento is None:
                raise RuntimeError("Word no devolvio el documento (ruta demasiado larga?)")
            paginas = int(documento.ComputeStatistics(_WD_STATISTIC_PAGES))
            self._fallos_seguidos = 0
            return paginas, ESTADO_OK
        except Exception as error:  # pragma: no cover - depende de Word instalado
            detalle = _detalle_com(error)
            self._fallos_seguidos += 1
            if self._fallos_seguidos >= _FALLOS_SEGUIDOS_PARA_RENDIRSE:
                self._no_disponible = f"Word desactivado tras varios fallos: {detalle}"
                self.cerrar()
            return None, f"Word no pudo abrir el documento: {detalle}"
        finally:
            if documento is not None:
                try:
                    documento.Close(SaveChanges=0)
                except Exception:
                    pass

    def cerrar(self) -> None:
        if self._word is not None:
            try:
                self._word.Quit()
            except Exception:  # pragma: no cover - Word ya pudo haber muerto
                pass
            self._word = None
        if self._pythoncom is not None:
            try:
                self._pythoncom.CoUninitialize()
            except Exception:  # pragma: no cover
                pass
            self._pythoncom = None


def paginas_pdf(ruta: Path) -> tuple[int | None, str]:
    """Numero de paginas de un PDF, leyendo la tabla de referencias cruzadas."""
    from pypdf import PdfReader

    # pypdf avisa por logging y por warnings de PDFs mal formados que de todos
    # modos consigue leer; en un inventario de miles de archivos eso sepulta el
    # progreso bajo cientos de lineas de ruido.
    logging.getLogger("pypdf").setLevel(logging.ERROR)

    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        lector = PdfReader(str(ruta), strict=False)
        if lector.is_encrypted:
            # Muchos PDF solo llevan contrasena de propietario: se abren con clave vacia.
            try:
                if not lector.decrypt(""):
                    return None, "PDF cifrado (requiere contrasena)"
            except Exception:
                return None, "PDF cifrado (requiere contrasena)"
        return len(lector.pages), ESTADO_OK


def paginas_ooxml(ruta: Path) -> tuple[int | None, str]:
    """Numero de paginas de un .docx leido de docProps/app.xml.

    El formato OOXML no pagina el documento: guarda el conteo que Word calculo
    la ultima vez que lo guardo. Un archivo generado por una libreria y nunca
    abierto en Word no trae el dato.
    """
    with zipfile.ZipFile(ruta) as zf:
        try:
            crudo = zf.read("docProps/app.xml")
        except KeyError:
            return None, "Sin metadatos de paginas (nunca guardado por Word)"

    raiz = ET.fromstring(crudo)
    elemento = raiz.find(f"{_NS_EXTENDIDO}Pages")
    if elemento is None or not (elemento.text or "").strip():
        return None, "Sin metadatos de paginas (nunca guardado por Word)"

    paginas = int(elemento.text.strip())
    return (paginas, ESTADO_METADATO) if paginas > 0 else (None, "Sin metadatos de paginas")


def paginas_doc_ole(ruta: Path) -> tuple[int | None, str]:
    """Numero de paginas de un .doc leido del stream SummaryInformation (OLE2)."""
    import olefile

    if not olefile.isOleFile(str(ruta)):
        return None, "No es un documento Word valido (OLE2)"

    with olefile.OleFileIO(str(ruta)) as ole:
        metadatos = ole.get_metadata()

    paginas = getattr(metadatos, "num_pages", None)
    if paginas:
        return int(paginas), ESTADO_METADATO
    return None, "Sin metadatos de paginas"


def _a_hora_local(valor: datetime | None) -> datetime | None:
    """Pasa una fecha con zona horaria a hora local sin zona.

    Excel no admite fechas con zona horaria, y Windows muestra "Guardado el" en
    hora local. Una fecha sin zona se deja como esta: es lo que ocurre con los PDF,
    cuyo /ModDate puede venir sin desfase y entonces ya se refiere a hora local.
    """
    if valor is None or valor.tzinfo is None:
        return valor
    return valor.astimezone().replace(tzinfo=None)


def _guardado_ooxml(ruta: Path) -> datetime | None:
    """Lee dcterms:modified de docProps/core.xml (en UTC)."""
    with zipfile.ZipFile(ruta) as zf:
        try:
            crudo = zf.read("docProps/core.xml")
        except KeyError:
            return None

    elemento = ET.fromstring(crudo).find(f"{_NS_TERMINOS}modified")
    texto = (elemento.text or "").strip() if elemento is not None else ""
    if not texto:
        return None
    return datetime.fromisoformat(texto)


def _guardado_ole(ruta: Path) -> datetime | None:
    """Lee la fecha de ultimo guardado del stream SummaryInformation (en UTC)."""
    import olefile

    if not olefile.isOleFile(str(ruta)):
        return None

    with olefile.OleFileIO(str(ruta)) as ole:
        valor = ole.get_metadata().last_saved_time

    # olefile devuelve la fecha sin zona, pero el formato OLE la guarda en UTC.
    return valor.replace(tzinfo=timezone.utc) if valor else None


def _guardado_pdf(ruta: Path) -> datetime | None:
    """Lee /ModDate del diccionario de informacion del PDF."""
    from pypdf import PdfReader

    logging.getLogger("pypdf").setLevel(logging.ERROR)
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        metadatos = PdfReader(str(ruta), strict=False).metadata
    return metadatos.modification_date if metadatos else None


def fecha_guardado(ruta: Path, extension: str) -> datetime | None:
    """Fecha que el documento registra como su ultimo guardado, en hora local.

    Es la columna "Guardado el" del Explorador de Windows. No tiene por que
    coincidir con la fecha de modificacion del disco: copiar un archivo actualiza
    la del disco y deja intacta la interna.
    """
    try:
        if extension in EXTENSIONES_OOXML_TODAS:
            return _a_hora_local(_guardado_ooxml(ruta))
        if extension in EXTENSIONES_OLE:
            return _a_hora_local(_guardado_ole(ruta))
        if extension == ".pdf":
            return _a_hora_local(_guardado_pdf(ruta))
    except Exception:
        # Es un dato accesorio: si el documento esta corrupto, la columna queda
        # vacia y el inventario sigue. El fallo ya se reporta en Estado.
        return None
    return None


def contar_paginas(
    ruta: Path,
    extension: str,
    *,
    word: WordCounter | None = None,
) -> tuple[int | None, str]:
    """Devuelve (paginas, estado) para el archivo indicado.

    Ningun archivo corrupto debe detener el inventario: cualquier error se
    traduce en `paginas = None` y un estado legible.
    """
    if extension not in EXTENSIONES_CON_PAGINAS:
        return None, ESTADO_NO_APLICA

    try:
        if extension == ".pdf":
            return paginas_pdf(ruta)

        if extension in EXTENSIONES_OOXML:
            # Word abre el documento y lo repagina: es lento, pero el metadato
            # de app.xml solo dice cuantas paginas tenia la ultima vez que Word
            # lo guardo, y puede estar desactualizado o no existir.
            if word is not None:
                paginas, estado = word.contar(ruta)
                if paginas is not None:
                    return paginas, estado
                respaldo, _ = paginas_ooxml(ruta)
                if respaldo is not None:
                    return respaldo, f"{ESTADO_METADATO}; {estado}"
                return None, estado
            return paginas_ooxml(ruta)

        # .doc / .dot: el metadato OLE es instantaneo y esta presente casi siempre.
        paginas, estado = paginas_doc_ole(ruta)
        if paginas is not None:
            return paginas, estado
        if word is None:
            return None, estado
        return word.contar(ruta)
    except Exception as error:
        return None, f"Error al leer ({type(error).__name__}: {error})"[:250]
