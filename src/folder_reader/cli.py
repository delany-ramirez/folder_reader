"""Interfaz de linea de comandos del inventario de carpetas."""

from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path

from folder_reader import __version__
from folder_reader.comprimidos import ResultadoDescompresion, descomprimir_en_raiz
from folder_reader.excel import Fila, escribir_excel
from folder_reader.pages import (
    ESTADO_NO_APLICA,
    EXTENSIONES_CON_PAGINAS,
    WordCounter,
    contar_paginas,
    fecha_guardado,
)
from folder_reader.scanner import ResultadoRecorrido, normalizar_raiz, recorrer


def construir_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="folder-reader",
        description=(
            "Recorre una carpeta y genera un Excel con la ruta, el nombre, las fechas "
            "y el numero de paginas (PDF y Word) de cada archivo encontrado."
        ),
        epilog='Ejemplo: folder-reader "D:\\Documentos" -o inventario.xlsx',
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "ruta",
        nargs="?",
        help="Carpeta a inventariar. Si se omite, se pregunta por consola.",
    )
    parser.add_argument(
        "-o",
        "--salida",
        help="Archivo .xlsx de salida (por defecto: inventario_<carpeta>_<fecha>.xlsx).",
    )
    parser.add_argument(
        "--incluir-ocultos",
        action="store_true",
        help="Incluye archivos y carpetas marcados como ocultos o de sistema.",
    )
    parser.add_argument(
        "--sin-paginas",
        action="store_true",
        help="No cuenta paginas (mucho mas rapido).",
    )
    parser.add_argument(
        "--sin-word",
        action="store_true",
        help=(
            "No abre los documentos en Word: usa el conteo que quedo guardado "
            "en los metadatos. Mucho mas rapido, pero puede estar desactualizado."
        ),
    )
    parser.add_argument(
        "--sin-descomprimir",
        action="store_true",
        help=(
            "No descomprime los .zip y .rar de la carpeta raiz. Por defecto se "
            "extraen en una carpeta con su mismo nombre y se inventaria su contenido."
        ),
    )
    parser.add_argument(
        "--profundidad",
        type=int,
        metavar="N",
        help="Limita el recorrido a N niveles (1 = solo la carpeta raiz).",
    )
    parser.add_argument(
        "--ext",
        nargs="+",
        metavar="EXT",
        help="Inventaria solo estas extensiones. Ejemplo: --ext .pdf .docx",
    )
    parser.add_argument("-v", "--verbose", action="store_true", help="Muestra el progreso.")
    parser.add_argument("--version", action="version", version=f"folder-reader {__version__}")
    return parser


def _preguntar_ruta() -> Path:
    """Pide la carpeta por consola hasta que sea valida."""
    while True:
        respuesta = input("Carpeta a inventariar: ").strip()
        if not respuesta:
            print("  Escriba una ruta o pulse Ctrl+C para salir.", file=sys.stderr)
            continue
        ruta = normalizar_raiz(respuesta)
        if ruta.is_dir():
            return ruta
        print(f"  No existe o no es una carpeta: {ruta}", file=sys.stderr)


def _preguntar_salida(predeterminado: str) -> str:
    respuesta = input(f"Archivo de salida [{predeterminado}]: ").strip().strip('"')
    return respuesta or predeterminado


def _nombre_predeterminado(raiz: Path) -> str:
    etiqueta = raiz.name or raiz.drive.replace(":", "").replace("\\", "") or "raiz"
    marca = datetime.now().strftime("%Y%m%d_%H%M%S")
    return f"inventario_{etiqueta}_{marca}.xlsx"


def _normalizar_extensiones(valores: list[str] | None) -> set[str] | None:
    if not valores:
        return None
    return {v if v.startswith(".") else f".{v}" for v in (x.lower() for x in valores)}


def main(argv: list[str] | None = None) -> int:
    args = construir_parser().parse_args(argv)

    if args.ruta:
        raiz = normalizar_raiz(args.ruta)
    elif sys.stdin is not None and sys.stdin.isatty():
        try:
            raiz = _preguntar_ruta()
        except (KeyboardInterrupt, EOFError):
            print("\nCancelado.", file=sys.stderr)
            return 130
    else:
        print(
            "Falta la carpeta a inventariar. Use: folder-reader <ruta> [-o salida.xlsx]",
            file=sys.stderr,
        )
        return 2

    if not raiz.is_dir():
        print(f"No existe o no es una carpeta: {raiz}", file=sys.stderr)
        return 2

    salida = args.salida
    if salida is None:
        predeterminado = _nombre_predeterminado(raiz)
        if not args.ruta and sys.stdin is not None and sys.stdin.isatty():
            try:
                salida = _preguntar_salida(predeterminado)
            except (KeyboardInterrupt, EOFError):
                print("\nCancelado.", file=sys.stderr)
                return 130
        else:
            salida = predeterminado

    destino = Path(salida).expanduser()
    if destino.suffix.lower() != ".xlsx":
        destino = destino.with_suffix(".xlsx")
    destino = destino.resolve()

    descompresion = ResultadoDescompresion()
    if not args.sin_descomprimir:
        descomprimir_en_raiz(
            raiz, incluir_ocultos=args.incluir_ocultos, resultado=descompresion
        )
        if args.verbose and descompresion.descomprimidos:
            print(
                f"Descomprimidos: {len(descompresion.descomprimidos)}", file=sys.stderr
            )

    print(f"Escaneando: {raiz}", file=sys.stderr)

    resultado = ResultadoRecorrido()
    filas: list[Fila] = []
    con_error = 0

    usar_word = not args.sin_paginas and not args.sin_word
    with WordCounter() as word:
        contador_word = word if usar_word else None
        for archivo in recorrer(
            raiz,
            incluir_ocultos=args.incluir_ocultos,
            profundidad=args.profundidad,
            extensiones=_normalizar_extensiones(args.ext),
            resultado=resultado,
        ):
            if args.sin_paginas:
                paginas, estado = None, ESTADO_NO_APLICA
            else:
                paginas, estado = contar_paginas(
                    archivo.ruta, archivo.extension, word=contador_word
                )

            if paginas is None and archivo.extension in EXTENSIONES_CON_PAGINAS:
                con_error += 1

            filas.append(
                Fila(
                    ruta=archivo.ruta_visible,
                    ruta_relativa=archivo.ruta_relativa,
                    carpeta=archivo.carpeta,
                    carpeta_nivel1=archivo.carpeta_nivel1,
                    nombre=archivo.nombre,
                    extension=archivo.extension,
                    tamano_bytes=archivo.tamano_bytes,
                    creado=archivo.creado,
                    modificado=archivo.modificado,
                    guardado=fecha_guardado(archivo.ruta, archivo.extension),
                    paginas=paginas,
                    estado=estado,
                )
            )

            if args.verbose and len(filas) % 100 == 0:
                print(f"\r  {len(filas)} archivos...", end="", file=sys.stderr, flush=True)

    if args.verbose:
        print(f"\r  {len(filas)} archivos.      ", file=sys.stderr)

    if not filas:
        print("No se encontro ningun archivo que inventariar.", file=sys.stderr)
        return 1

    try:
        escritas = escribir_excel(filas, destino)
    except PermissionError:
        print(
            f"No se pudo escribir {destino}. Cierre el archivo en Excel e intente de nuevo.",
            file=sys.stderr,
        )
        return 1
    except OSError as error:
        print(f"No se pudo escribir {destino}: {error}", file=sys.stderr)
        return 1

    _resumen(destino, filas, escritas, resultado, descompresion, con_error, args)
    return 0


def _resumen(
    destino: Path,
    filas: list[Fila],
    escritas: int,
    resultado: ResultadoRecorrido,
    descompresion: ResultadoDescompresion,
    con_error: int,
    args: argparse.Namespace,
) -> None:
    print(f"\nArchivos inventariados: {len(filas)}")
    if escritas < len(filas):
        print(f"  Truncado a {escritas} filas por el limite de Excel.")
    print(f"Archivos de sistema u ocultos omitidos: {resultado.omitidos_sistema}")
    if resultado.omitidos_extension:
        print(f"Archivos omitidos por el filtro --ext: {resultado.omitidos_extension}")
    if not args.sin_paginas:
        print(f"Documentos sin numero de paginas: {con_error} (vea la columna Estado)")
    if resultado.carpetas_inaccesibles:
        print(f"Carpetas sin permiso de lectura: {len(resultado.carpetas_inaccesibles)}")
        for carpeta in resultado.carpetas_inaccesibles[:5]:
            print(f"  - {carpeta}")
        if len(resultado.carpetas_inaccesibles) > 5:
            print(f"  ... y {len(resultado.carpetas_inaccesibles) - 5} mas")
    if descompresion.descomprimidos:
        print(f"Comprimidos descomprimidos en la raiz: {len(descompresion.descomprimidos)}")
    if descompresion.ya_existian:
        print(
            "Comprimidos no extraidos porque ya existe su carpeta: "
            f"{len(descompresion.ya_existian)}"
        )
    if descompresion.fallidos:
        print(f"Comprimidos que no se pudieron descomprimir: {len(descompresion.fallidos)}")
        for nombre, motivo in descompresion.fallidos:
            print(f"  - {nombre}: {motivo}")
    if resultado.archivos_inaccesibles:
        print(f"Archivos sin permiso de lectura: {len(resultado.archivos_inaccesibles)}")
    print(f"\nExcel generado: {destino}")


if __name__ == "__main__":
    raise SystemExit(main())
