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
| `--sin-descomprimir` | No extrae los `.zip` y `.rar` de la carpeta raíz (ver abajo). |
| `--profundidad N` | Limita el recorrido a N niveles (`1` = solo la carpeta raíz). |
| `--ext .pdf .docx` | Inventaria solo esas extensiones. |
| `-v`, `--verbose` | Muestra el avance del escaneo. |

## Archivos comprimidos (`.zip`, `.rar`)

Antes de escanear, cada `.zip` o `.rar` que esté **directamente en la carpeta raíz** se
extrae en una carpeta hermana con su mismo nombre (`Anexos.zip` → `Anexos\`), y su
contenido se inventaria como el de cualquier otra subcarpeta: con páginas, fechas y
**Carpeta nivel 1** = `Anexos`. El comprimido original no se toca y también aparece en el
inventario como un archivo más.

- Los comprimidos que están en subcarpetas, o dentro de otro comprimido, **no** se extraen.
- Si la carpeta de destino ya existe, se da por descomprimido y no se toca nada. Por eso
  volver a correr el inventario no duplica ni pisa archivos.
- Si el comprimido trae una sola carpeta con su mismo nombre (lo típico al comprimir una
  carpeta entera), se evita el `Anexos\Anexos\…`, igual que «Extraer aquí».
- En los `.zip` se conserva la fecha de modificación original de cada archivo, y los
  nombres con tildes o eñes creados por el Explorador de Windows se leen bien. La fecha de
  creación, en cambio, es la de la extracción.
- Los `.zip` se extraen con Python. Los `.rar` (y los `.zip` con métodos que Python no
  soporta, como Deflate64) se extraen con **WinRAR** o **7-Zip** si están instalados, y
  si no, con el `tar.exe` que trae Windows 10/11.
- Un comprimido dañado o con contraseña no detiene nada: el resumen final lo lista con el
  motivo, y no quedan carpetas a medio extraer.

## Hojas de control (`hoja-control`)

Un segundo comando vuelca el contenido de las carpetas en el formato institucional
**Hoja de control** (Gestión de Documentos, código 1122 - F15). Se genera **un archivo por
carpeta madre**: cada subcarpeta de primer nivel de la ruta indicada.

```bash
uv run hoja-control "D:\Contratos 2022" -o "D:\Hojas de control"
```

Con una carpeta `6170-2022` dentro de `D:\Contratos 2022` se obtiene
`HOJA DE CONTROL 6170-2022.xlsx`. Dentro de cada hoja:

- Cada subcarpeta abre con una fila en **negrilla** con su nombre (`1. PRECONTRACTUAL`), y
  debajo van sus archivos. Las subcarpetas anidadas muestran la ruta relativa a la carpeta
  madre (`2. CONTRACTUAL\5. INFORMES`). Los archivos que están en la propia carpeta madre
  van primero, sin fila de carpeta.
- Carpetas y archivos se ordenan como en el Explorador, comparando los números por su
  valor: `1, 2, 3, 10, 11` (no `1, 10, 11, 2`), `2. RP` antes que `10. CEDULA` y `2.9`
  antes que `2.10`.
- **ITEM** numera todas las filas, también las de carpeta, como en la plantilla.
- **FECHA DOCUMENTO** y **FECHA REGISTRO** llevan las dos la fecha de guardado del
  documento (la de [«Guardado el»](#guardado-el-documento)), sin hora y en `dd/mm/aaaa`.
  Los archivos que no registran esa fecha (imágenes, `.txt`, PDF sin `/ModDate`) las dejan
  **en blanco**.
- **CANTIDAD DE FOLIOS** es el número de páginas, calculado igual que en el inventario. Queda
  vacía en los formatos sin páginas (`.xlsx`, imágenes…). **FOLIO** se deja vacía para
  diligenciarla a mano.
- La fila **TOTAL** suma los folios y, debajo, se conserva el pie de firmas.

Los `.zip` y `.rar` de la raíz se descomprimen antes, igual que en el inventario, así que
cada comprimido se convierte en una carpeta madre más. Los archivos sueltos en la raíz no
pertenecen a ninguna carpeta madre: se cuentan en el resumen y no van en ninguna hoja.

| Opción | Para qué sirve |
| --- | --- |
| `-o`, `--salida` | Carpeta donde se guardan los `.xlsx` (por defecto, la actual). Si ya existe una hoja con el mismo nombre, se reemplaza. |
| `--carpeta-madre` | La ruta es ella misma una carpeta madre: se genera una sola hoja. |
| `--plantilla` | Usa otra plantilla `.xlsx` en lugar de la incluida. |
| `--sin-word`, `--sin-descomprimir`, `--incluir-ocultos`, `--ext` | Igual que en `folder-reader`. |

### La plantilla

La plantilla incluida está en `src/folder_reader/plantillas/hoja_control.xlsx`. El programa
no reconstruye el formato: copia el de la plantilla. Busca la fila de encabezado (la que
tiene `ITEM`), toma como modelo la primera fila con el nombre en negrilla (carpeta) y la
primera sin negrilla (documento), y replica esas filas tantas veces como haga falta. Todo lo
que hay desde la fila `TOTAL` hacia abajo se reubica debajo de los datos, y la fórmula
`=SUM(…)` se ajusta al rango real.

Por eso una plantilla propia (`--plantilla`), incluso una hoja ya diligenciada, funciona
siempre que tenga las columnas `ITEM`, `FECHA DOCUMENTO`, `TIPO DOCUMENTAL`,
`CANTIDAD DE FOLIOS` y `FECHA REGISTRO`, y una fila `TOTAL`.

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
