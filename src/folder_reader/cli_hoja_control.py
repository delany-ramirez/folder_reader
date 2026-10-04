"""Comando hoja-control: una hoja de control documental por cada carpeta madre."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from folder_reader import __version__
from folder_reader.cli import _normalizar_extensiones
from folder_reader.comprimidos import (
    ResultadoDescompresion,
    comprimidos_en_raiz,
    descomprimir_en_raiz,
)
from folder_reader.hoja_control import (
    Documento,
    ErrorPlantilla,
    Grupo,
    clave_natural,
    escribir_hoja_control,
)
from folder_reader.pages import (
    EXTENSIONES_CON_PAGINAS,
    WordCounter,
    contar_paginas,
    fecha_guardado,
)
from folder_reader.progreso import Progreso
from folder_reader.scanner import ArchivoInfo, ResultadoRecorrido, normalizar_raiz, recorrer

PREFIJO_SALIDA = "HOJA DE CONTROL"

# Documentos de una carpeta madre, por subcarpeta relativa a ella (() = la propia madre).
_Subcarpetas = dict[tuple[str, ...], list[Documento]]


def construir_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="hoja-control",
        description=(
            "Genera una hoja de control documental (.xlsx, a partir de una plantilla) "
            "por cada carpeta madre: cada subcarpeta de primer nivel de la ruta."
        ),
        epilog='Ejemplo: hoja-control "D:\\Contratos 2022" -o "D:\\Hojas de control"',
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("ruta", help="Carpeta que contiene las carpetas madre.")
    parser.add_argument(
        "-o",
        "--salida",
        help="Carpeta donde se guardan los .xlsx (por defecto: el directorio actual).",
    )
    parser.add_argument(
        "--plantilla",
        help="Plantilla .xlsx a usar en lugar de la incluida.",
    )
    parser.add_argument(
        "--carpeta-madre",
        action="store_true",
        help="La ruta es ella misma una carpeta madre: genera un solo archivo.",
    )
    parser.add_argument(
        "--incluir-ocultos",
        action="store_true",
        help="Incluye archivos y carpetas marcados como ocultos o de sistema.",
    )
    parser.add_argument(
        "--sin-word",
        action="store_true",
        help="No abre los documentos en Word: usa el conteo guardado en los metadatos.",
    )
    parser.add_argument(
        "--sin-descomprimir",
        action="store_true",
        help="No descomprime los .zip y .rar de la carpeta raiz.",
    )
    parser.add_argument(
        "--ext",
        nargs="+",
        metavar="EXT",
        help="Incluye solo estas extensiones. Ejemplo: --ext .pdf .docx",
    )
    parser.add_argument("--version", action="version", version=f"hoja-control {__version__}")
    return parser


def _ubicar(
    archivo: ArchivoInfo, raiz: Path, una_sola: bool
) -> tuple[str, tuple[str, ...]] | None:
    """Devuelve (carpeta madre, subcarpetas dentro de ella), o None si no tiene madre."""
    partes = Path(archivo.ruta_relativa).parts[:-1]
    if una_sola:
        return raiz.name or str(raiz), partes
    if not partes:
        return None
    return partes[0], partes[1:]


def _grupos(subcarpetas: _Subcarpetas) -> list[Grupo]:
    """Ordena las carpetas y sus archivos como el Explorador (orden natural).

    Los archivos de la propia carpeta madre van primero y sin fila de carpeta.
    """
    return [
        Grupo(
            "\\".join(partes),
            sorted(subcarpetas[partes], key=lambda doc: clave_natural(doc.nombre)),
        )
        for partes in sorted(subcarpetas, key=lambda p: [clave_natural(x) for x in p])
    ]


def nombre_archivo(carpeta: str) -> str:
    # Los nombres de carpeta ya son validos en Windows, asi que sirven tal cual.
    return f"{PREFIJO_SALIDA} {carpeta}.xlsx"


def main(argv: list[str] | None = None) -> int:
    args = construir_parser().parse_args(argv)

    raiz = normalizar_raiz(args.ruta)
    if not raiz.is_dir():
        print(f"No existe o no es una carpeta: {raiz}", file=sys.stderr)
        return 2

    plantilla = None
    if args.plantilla:
        plantilla = normalizar_raiz(args.plantilla)
        if not plantilla.is_file():
            print(f"No existe la plantilla: {plantilla}", file=sys.stderr)
            return 2

    salida = normalizar_raiz(args.salida) if args.salida else Path.cwd()

    descompresion = ResultadoDescompresion()
    if not args.sin_descomprimir:
        pendientes = comprimidos_en_raiz(raiz, incluir_ocultos=args.incluir_ocultos)
        if pendientes:
            print(f"Revisando {len(pendientes)} comprimidos de la raiz...", file=sys.stderr)
        descomprimir_en_raiz(raiz, incluir_ocultos=args.incluir_ocultos, resultado=descompresion)

    # Primero se listan los archivos (rapido: solo lee los directorios) para conocer
    # el total; el trabajo lento, contar paginas, va despues con porcentaje.
    print(f"Buscando archivos en: {raiz}", file=sys.stderr)
    resultado = ResultadoRecorrido()
    sueltos: list[str] = []
    por_procesar: list[tuple[ArchivoInfo, str, tuple[str, ...]]] = []
    for archivo in recorrer(
        raiz,
        incluir_ocultos=args.incluir_ocultos,
        extensiones=_normalizar_extensiones(args.ext),
        resultado=resultado,
    ):
        ubicacion = _ubicar(archivo, raiz, args.carpeta_madre)
        if ubicacion is None:
            # Archivos sueltos en la raiz (incluidos los .zip ya extraidos):
            # no pertenecen a ninguna carpeta madre.
            sueltos.append(archivo.nombre)
        else:
            por_procesar.append((archivo, *ubicacion))

    aviso_word = "" if args.sin_word else " (los Word se abren uno por uno, puede tardar)"
    print(f"Leyendo {len(por_procesar)} archivos{aviso_word}...", file=sys.stderr)
    madres: dict[str, _Subcarpetas] = {}
    sin_folios: list[str] = []
    sin_fecha = 0
    progreso = Progreso(len(por_procesar))

    with WordCounter() as word:
        for hechos, (archivo, nombre_madre, subcarpetas) in enumerate(por_procesar):
            # Se informa antes de procesar: si un documento se queda pegado en
            # Word, la linea muestra cual es.
            progreso.avanzar(hechos, archivo.ruta_relativa)

            paginas, _ = contar_paginas(
                archivo.ruta, archivo.extension, word=None if args.sin_word else word
            )
            if paginas is None and archivo.extension in EXTENSIONES_CON_PAGINAS:
                sin_folios.append(archivo.ruta_relativa)

            # La fecha de la hoja es la de guardado del documento. Los formatos que
            # no la registran (imagenes, .txt, PDF sin /ModDate) la dejan en blanco.
            fecha = fecha_guardado(archivo.ruta, archivo.extension)
            if fecha is None:
                sin_fecha += 1

            madres.setdefault(nombre_madre, {}).setdefault(subcarpetas, []).append(
                Documento(archivo.nombre, fecha.date() if fecha else None, paginas)
            )

    progreso.terminar()

    if not madres:
        print("No se encontro ninguna carpeta madre con archivos.", file=sys.stderr)
        return 1

    generados: list[tuple[Path, int]] = []
    fallidos: list[tuple[str, str]] = []
    ordenadas = sorted(madres, key=clave_natural)
    for numero, nombre_madre in enumerate(ordenadas, start=1):
        destino = salida / nombre_archivo(nombre_madre)
        print(f"Escribiendo hoja {numero}/{len(ordenadas)}: {destino.name}", file=sys.stderr)
        try:
            documentos = escribir_hoja_control(_grupos(madres[nombre_madre]), destino, plantilla)
        except ErrorPlantilla as error:
            print(f"Plantilla no valida: {error}", file=sys.stderr)
            return 2
        except PermissionError:
            fallidos.append((destino.name, "esta abierto en Excel o no hay permiso"))
            continue
        except OSError as error:
            fallidos.append((destino.name, str(error)))
            continue
        generados.append((destino, documentos))

    print(f"\nHojas de control generadas: {len(generados)} en {salida}")
    for destino, documentos in generados:
        print(f"  - {destino.name}: {documentos} documentos")
    if fallidos:
        print(f"No se pudieron escribir: {len(fallidos)}")
        for nombre, motivo in fallidos:
            print(f"  - {nombre}: {motivo}")
    if sueltos:
        print(f"Archivos sueltos en la raiz, fuera de toda carpeta madre: {len(sueltos)}")
    if sin_folios:
        print(f"Documentos sin numero de folios: {len(sin_folios)}")
        for ruta in sin_folios[:5]:
            print(f"  - {ruta}")
        if len(sin_folios) > 5:
            print(f"  ... y {len(sin_folios) - 5} mas")
    if sin_fecha:
        print(f"Documentos sin fecha de guardado (quedan en blanco): {sin_fecha}")
    if descompresion.fallidos:
        print(f"Comprimidos que no se pudieron descomprimir: {len(descompresion.fallidos)}")
        for nombre, motivo in descompresion.fallidos:
            print(f"  - {nombre}: {motivo}")
    if resultado.carpetas_inaccesibles:
        print(f"Carpetas sin permiso de lectura: {len(resultado.carpetas_inaccesibles)}")

    return 1 if fallidos else 0


if __name__ == "__main__":
    raise SystemExit(main())
