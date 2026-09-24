# folder-reader

Recorre una carpeta y genera un Excel con la ruta, el nombre, las fechas y el número de
páginas de cada archivo encontrado. Los archivos del sistema (`desktop.ini`, `Thumbs.db`,
temporales `~$…` de Office, elementos ocultos) se ignoran.

## Instalación

El proyecto usa [uv](https://docs.astral.sh/uv/). No hay que crear ni activar el entorno a
mano: `uv run` lo prepara la primera vez a partir de `pyproject.toml` y `uv.lock`.

```bash
uv sync
```

## Uso

```bash
# Forma habitual
uv run folder-reader "D:\Documentos" -o inventario.xlsx

# Sin argumentos: pregunta la carpeta y el archivo de salida por consola
uv run folder-reader
```

Si se omite `-o`, el archivo se llama `inventario_<carpeta>_<AAAAMMDD_HHMMSS>.xlsx` en el
directorio actual, de modo que una corrida nunca pisa a la anterior.

### Opciones

| Opción | Para qué sirve |
| --- | --- |
| `-o`, `--salida` | Archivo `.xlsx` de salida. |
| `--incluir-ocultos` | Incluye archivos y carpetas marcados como ocultos o de sistema. |
| `--sin-paginas` | No cuenta páginas. Mucho más rápido si solo interesan las fechas. |
| `--sin-word` | No abre los documentos en Word: usa el conteo guardado en los metadatos. Mucho más rápido, pero puede estar desactualizado. |
| `--profundidad N` | Limita el recorrido a N niveles (`1` = solo la carpeta raíz). |
| `--ext .pdf .docx` | Inventaria solo esas extensiones. |
| `-v`, `--verbose` | Muestra el avance del escaneo. |

## Qué se ignora

- Archivos que crea el sistema operativo por su cuenta: `desktop.ini`, `Thumbs.db`,
  `.DS_Store`, los temporales `~$…` de Office, `pagefile.sys`, `ntuser.dat*`.
- Carpetas del sistema: `$Recycle.Bin`, `System Volume Information`, `.git`, `__pycache__`.
- Todo lo marcado como **oculto**, salvo que se use `--incluir-ocultos`.

El atributo **«sistema» por sí solo no descarta nada**, y es importante: Windows se lo pone
a cualquier carpeta personalizada con un icono propio, y OneDrive se lo pone a todo lo que
sincroniza, Escritorio incluido. Filtrar por él dejaría fuera carpetas de trabajo normales.
Lo verdaderamente intocable (`$Recycle.Bin`, `System Volume Information`) está marcado como
oculto además de sistema, así que sigue quedando fuera.

## Columnas del Excel

Ruta completa · Ruta relativa · Carpeta contenedora · **Carpeta nivel 1** ·
Nombre del archivo · Extensión ·
Tamaño (bytes) · Tamaño (MB) · Fecha de creación · Fecha de última modificación ·
**Guardado el (documento)** · Número de páginas · Estado

La hoja sale con fila de encabezado fija, autofiltro y las fechas como fechas reales de
Excel, para poder ordenar y filtrar directamente.

**Carpeta contenedora** es la ruta relativa completa de la carpeta, empezando por el nombre
de la carpeta escaneada: `MES 7`, `MES 7\A`, `MES 7\A\1.Ajuste informe 2 proyecto`. Así
identifica la ubicación sin ambigüedad cuando varias subcarpetas se llaman igual, y permite
agrupar por rama en una tabla dinámica sin arrastrar la ruta absoluta entera.

**Carpeta nivel 1** es solo la primera subcarpeta bajo la raíz: para `MES 7\A\1.Ajuste
informe 2 proyecto` devuelve `A`. Sirve para agrupar o filtrar el inventario por rama
principal sin tocar el resto del árbol. Los archivos que están en la propia raíz la dejan
vacía, porque no cuelgan de ninguna subcarpeta.

La columna **Estado** dice qué pasó con el conteo de páginas: de dónde salió la cifra, o el
motivo concreto del fallo. Ningún archivo ilegible interrumpe el inventario.

### «Guardado el (documento)»

Es la columna *Guardado el* del Explorador de Windows: la fecha que **el propio documento**
registra como su último guardado, no la que tiene el archivo en el disco. Las dos suelen
coincidir, pero no siempre: copiar o descargar un archivo actualiza la fecha del disco y deja
intacta la interna, así que esta columna dice cuándo se editó el contenido de verdad.

Sale de los metadatos del documento, convertida de UTC a hora local:

| Formato | De dónde |
| --- | --- |
| `.docx`, `.xlsx`, `.pptx` y demás OOXML | `dcterms:modified` de `docProps/core.xml` |
| `.doc`, `.xls`, `.ppt` (heredados) | stream `SummaryInformation` del contenedor OLE2 |
| `.pdf` | `/ModDate` del diccionario de información |

Queda vacía si el documento no trae el dato, y en los formatos que no lo tienen (`.txt`,
imágenes, `.zip`…). En los PDF el Explorador la muestra siempre vacía; aquí sí se rellena
cuando el archivo la trae.

## De dónde sale el número de páginas

La columna **Estado** distingue de dónde salió cada cifra, porque no todas tienen la misma
fiabilidad:

- `OK` — el número se calculó ahora, leyendo o repaginando el documento.
- `OK (metadato guardado, sin abrir en Word)` — es el conteo que el programa dejó escrito
  la última vez que se guardó el archivo. Puede no corresponder al contenido actual.

Por formato:

- **`.pdf`** — se cuentan las páginas reales del documento con `pypdf`. Un PDF protegido con
  contraseña de usuario queda con estado `PDF cifrado`.
- **`.docx` / `.docm`** — el documento **se abre en Word**, que lo repagina y devuelve el
  número real. El formato OOXML no pagina por sí mismo: el metadato de `docProps/app.xml`
  solo dice cuántas páginas tenía cuando Word lo guardó, y a menudo está desactualizado o
  directamente ausente. Si Word no está disponible o no puede abrir el archivo, se cae al
  metadato y el estado lo indica.
- **`.doc`** (Word 97-2003) — se lee el contador del stream `SummaryInformation` del archivo
  OLE2 con `olefile`, que está presente casi siempre y es instantáneo. Si falta, se abre el
  documento en Word.

Con `--sin-word` nada se abre en Word y todo sale de metadatos: el inventario es unas veinte
veces más rápido, a cambio de fiabilidad.

### Rendimiento

Abrir cada documento en Word cuesta alrededor de **1 segundo por archivo**. En una medición
real sobre 507 archivos con 28 `.docx`, el inventario completo tardó 22 s; con `--sin-word`,
1,5 s. Vale la pena: en esa misma carpeta, 3 documentos no tenían ningún metadato de páginas
—uno de ellos de 436 páginas— y otro lo tenía desactualizado.

Si Word falla cinco veces seguidas (formato bloqueado, Word cerrado a la fuerza), el programa
deja de intentarlo y termina el inventario con metadatos, en vez de arrastrar el coste en
cada archivo restante.

> **Nota sobre el Centro de confianza.** Word puede tener bloqueada la apertura de archivos
> `.doc` heredados; en ese caso la columna Estado lo dice textualmente. Se habilita en
> *Word → Opciones → Centro de confianza → Configuración del Centro de confianza →
> Configuración de bloqueo de archivos*. En la práctica casi ningún `.doc` lo necesita,
> porque el metadato OLE suele estar presente.

## Desarrollo

```bash
uv run pytest
```

Los tests construyen un árbol de prueba con PDFs reales de 1, 3 y 10 páginas, un `.docx`
mínimo, basura del sistema y un PDF deliberadamente corrupto.
