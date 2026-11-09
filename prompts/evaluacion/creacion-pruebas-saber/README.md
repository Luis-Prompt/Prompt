# Creación de pruebas tipo SABER (v2.4)

Genera **dos documentos Word** a partir de las respuestas del docente:

- `…_ESTUDIANTE.docx`: encabezado con escudo, datos del estudiante, preguntas a dos columnas y hoja de respuestas con óvalos.
- `…_DOCENTE_CLAVES.docx`: la misma prueba con la clave resaltada, plantilla de calificación, justificación de cada respuesta, error que representa cada distractor y **matriz de especificaciones** (competencias × niveles).

**Versiones A y B (opcional):** con `"versiones": 2` en el `config` del JSON, o `--versiones 2` al ejecutar, se generan cuatro Word: estudiante y docente de la versión A y de la versión B (otro orden de preguntas y otras letras de clave) para evitar la copia.

## Contenido de la carpeta

| Archivo | Función |
|---|---|
| `SKILL.md` | El prompt: preguntas iniciales, reglas de diseño de preguntas y pasos de entrega. |
| `scripts/build_exam.js` | Construye los Word, valida las preguntas y verifica la maquetación. |
| `assets/escudo_header.png` | Escudo 2025 del encabezado (copia del banco de imágenes). |
| `examples/exam_ejemplo.json` | Ejemplo completo (14 preguntas de selección múltiple + 2 abiertas, 6 gráficos). |
| `package.json` | Dependencias (`docx`, `sharp`). |
| `CHANGELOG.md` | Historial de versiones. |
| `ROADMAP.md` | Mejoras pendientes, por prioridad. |

## Probar en el computador

Requisitos: Node 18+, LibreOffice y `poppler-utils` (para medir la maquetación).

```bash
cd prompts/evaluacion/creacion-pruebas-saber
npm install
npm run ejemplo          # genera los Word en ./salida
node scripts/build_exam.js examples/exam_ejemplo.json --out salida --versiones 2   # versiones A y B
```

## Uso con Claude

Entrega a Claude el enlace de `SKILL.md`. Él hará las preguntas iniciales (tema, grado, número de preguntas, dificultad, apoyos visuales), construirá el `exam.json`, ejecutará el script y entregará los dos Word.

## Datos de la institución

El nombre, el escudo y el membrete (título, código de formato, versión y fecha) viven en [`config/institucion.json`](../../../config/institucion.json), compartido con los demás prompts. El script lo encuentra solo, o se indica con `--institucion archivo.json`. Para una sola prueba con otro membrete, se puede poner `config.membrete` dentro del `exam.json`.

**Escudo:** acepta PNG, JPG o WebP (aunque la extensión no coincida) y recorta los márgenes blancos. Para cambiarlo, usa `herramientas/preparar-escudo.py` (ver `banco-imagenes/README.md`).

**Banco de preguntas:** `node ../../../herramientas/guardar-en-banco.js exam.json` guarda las preguntas de una prueba en `banco-preguntas/`.
