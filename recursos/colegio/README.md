# Datos del colegio

```
colegio/
├─ institucion.json              nombre, lema, membretes y rutas del escudo
├─ perfil-docente.ejemplo.json   formato del perfil del docente (se reutiliza entre sesiones)
└─ escudos/                      escudo institucional en tres tamaños
```

## institucion.json
Guarda los datos del colegio que usan **todos** los prompts: nombre, ciudad, lema, rutas del escudo y membrete de los documentos (título, código de formato, versión y fecha).
- Los generadores lo buscan solos (o con `--institucion otro.json`).
- Las rutas de `escudo` son relativas a este archivo.
- Para usar el repositorio con otra institución, edita este archivo y cambia las imágenes con `herramientas/preparar-escudo.py`.
- Un `config.membrete` dentro de un `exam.json` tiene prioridad sobre este archivo (útil para un formato especial de una sola prueba).

## Escudo del colegio (`escudos/`)
| Archivo | Para qué |
|---|---|
| `escudo_colegio_simon_bolivar_2025.png` | Maestro recortado (alto 2400 px): afiches, diplomas, impresión grande |
| `escudo_colegio_simon_bolivar_2025_encabezado.png` | 360 px de ancho: encabezado de los documentos Word |
| `escudo_colegio_simon_bolivar_2025_miniatura.png` | 120 px de ancho: tablas, web, README |

**Si cambia el escudo:** `python3 herramientas/preparar-escudo.py NUEVO.png --nombre escudo_colegio_simon_bolivar_2026 --salida recursos/colegio/escudos` y actualiza las rutas en `institucion.json`. Ya no hay que copiarlo a cada prompt: `preparar-prompt.sh` lo coloca solo.

Reglas: nombres sin tildes ni espacios; preferir **SVG** para gráficos (se ve nítido al imprimir); no subir imágenes con derechos de autor ajenos ni fotos de estudiantes.
