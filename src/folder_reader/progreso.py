"""Indicador de progreso en una sola linea de consola."""

from __future__ import annotations

import shutil
import sys
import time
from typing import TextIO

# Redibujar la linea mas seguido no aporta nada y, con miles de archivos rapidos,
# la consola de Windows se vuelve el cuello de botella.
_INTERVALO = 0.2


def _duracion(segundos: float) -> str:
    if segundos < 60:
        return f"{segundos:.0f} s"
    if segundos < 3600:
        return f"{segundos / 60:.0f} min"
    return f"{segundos // 3600:.0f} h {segundos % 3600 / 60:.0f} min"


class Progreso:
    """Muestra `[ 45%] 123/274  ~2 min restantes  carpeta\\archivo.pdf`.

    En una consola la linea se reescribe en el sitio. Si la salida esta redirigida
    a un archivo, se escribe una linea cada 10 % para no llenarlo de basura.
    """

    def __init__(self, total: int, *, salida: TextIO | None = None) -> None:
        self.total = total
        self.salida = salida if salida is not None else sys.stderr
        self.interactiva = self.salida.isatty()
        self.inicio = time.monotonic()
        self._ultimo_dibujo = 0.0
        self._ultimo_decil = -1
        self._ancho_previo = 0

    def avanzar(self, hechos: int, actual: str = "") -> None:
        """Informa que van `hechos` terminados y que se esta procesando `actual`."""
        porcentaje = 100 * hechos // self.total if self.total else 100
        if not self.interactiva:
            decil = porcentaje // 10
            if decil != self._ultimo_decil:
                self._ultimo_decil = decil
                print(f"  {porcentaje}% ({hechos}/{self.total})", file=self.salida, flush=True)
            return

        ahora = time.monotonic()
        if hechos < self.total and ahora - self._ultimo_dibujo < _INTERVALO:
            return
        self._ultimo_dibujo = ahora

        texto = f"  [{porcentaje:3d}%] {hechos}/{self.total}"
        transcurrido = ahora - self.inicio
        # Con pocos archivos hechos la estimacion salta demasiado para ser util.
        if 0 < hechos < self.total and transcurrido > 3:
            restante = transcurrido / hechos * (self.total - hechos)
            texto += f"  ~{_duracion(restante)} restantes"
        if actual:
            texto += f"  {actual}"
        self._dibujar(texto)

    def terminar(self) -> None:
        self.avanzar(self.total)
        if self.interactiva:
            transcurrido = time.monotonic() - self.inicio
            self._dibujar(f"  [100%] {self.total}/{self.total} en {_duracion(transcurrido)}")
            print(file=self.salida, flush=True)

    def _dibujar(self, texto: str) -> None:
        ancho = shutil.get_terminal_size((100, 20)).columns - 1
        if len(texto) > ancho:
            # Se recorta por el principio de la ruta: el nombre del archivo es lo
            # que mas interesa cuando algo se queda pegado.
            texto = texto[: ancho // 2] + "…" + texto[-(ancho - ancho // 2 - 1) :]
        relleno = " " * max(0, self._ancho_previo - len(texto))
        self._ancho_previo = len(texto)
        print(f"\r{texto}{relleno}", end="", file=self.salida, flush=True)
