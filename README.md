# Lector Docs

App Android para leer PDF, Word (.docx) y PowerPoint (.pptx) sin internet.

- Abre archivos desde el botón, desde "Abrir con…" del gestor de archivos o al compartirlos hacia la app.
- Lista de recientes, zoom con botones, pellizco o doble toque, ir a página y modo nocturno.
- Todo se procesa en el teléfono; ningún archivo se envía a internet.

## Generar el APK

### Opción A: GitHub (sin instalar nada)

1. Crea un repositorio nuevo en GitHub y sube el contenido de esta carpeta (incluida la carpeta oculta `.github`).
2. Entra a la pestaña **Actions**. El flujo **Generar APK** se ejecuta solo con cada push (o con *Run workflow*).
3. Cuando termine (unos 3–5 minutos), abre la ejecución y descarga el artefacto **LectorDocs-apk**. Dentro está `LectorDocs.apk`.

### Opción B: Android Studio

1. *File → Open* y selecciona esta carpeta. Acepta la sincronización de Gradle.
2. *Build → Build App Bundle(s) / APK(s) → Build APK(s)*.
3. El APK queda en `app/build/outputs/apk/debug/app-debug.apk`.

## Instalar en el teléfono

Copia el APK al teléfono, ábrelo y permite "Instalar apps de fuentes desconocidas" cuando Android lo pida.

## Estructura

| Ruta | Qué contiene |
|---|---|
| `app/src/main/java/.../MainActivity.java` | Contenedor nativo: WebView, selector de archivos, "Abrir con", compartir |
| `app/src/main/assets/www/app.js` | Interfaz, visor PDF (pdf.js) y Word (docx-preview), zoom, recientes |
| `app/src/main/assets/www/pptx.js` | Renderizador propio de PowerPoint |
| `app/src/main/assets/www/lib/` | Librerías incluidas (pdf.js, docx-preview, JSZip) con sus licencias |

Para cambiar el nombre visible edita `res/values/strings.xml`; para cambiar el identificador del paquete, `applicationId` en `app/build.gradle`.

## Limitaciones

- Formatos antiguos `.doc` y `.ppt` no se abren: hay que guardarlos como `.docx` / `.pptx`.
- En PowerPoint, los gráficos (charts) y animaciones no se dibujan; SmartArt, tablas, imágenes, formas y texto sí.
- El APK de este flujo es de depuración (firmado automáticamente, instalable directo). Para publicarlo en Play Store hay que firmar una versión *release* con tu propia llave.
