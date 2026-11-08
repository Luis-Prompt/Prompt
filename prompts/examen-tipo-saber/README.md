# Generador de examen tipo SABER (v2.1)

Genera **dos documentos Word** a partir de las respuestas del docente:

- `…_ESTUDIANTE.docx`: encabezado con escudo, datos del estudiante, preguntas a dos columnas y hoja de respuestas con óvalos.
- `…_DOCENTE_CLAVES.docx`: la misma prueba con la clave resaltada, plantilla de calificación, justificación de cada respuesta y error que representa cada distractor.

## Contenido de la carpeta

| Archivo | Función |
|---|---|
| `SKILL.md` | El prompt: preguntas iniciales, reglas de diseño de preguntas y pasos de entrega. |
| `scripts/build_exam.js` | Construye los Word, valida las preguntas y verifica la maquetación. |
| `assets/escudo_header.png` | Escudo institucional del encabezado. |
| `examples/exam_ejemplo.json` | Ejemplo completo (14 preguntas de selección múltiple + 2 abiertas, 6 gráficos). |
| `package.json` | Dependencias (`docx`, `sharp`). |
| `CHANGELOG.md` | Historial de versiones. |
| `ROADMAP.md` | Mejoras pendientes, por prioridad. |

## Probar en el computador

Requisitos: Node 18+, LibreOffice y `poppler-utils` (para medir la maquetación).

```bash
cd prompts/examen-tipo-saber
npm install
npm run ejemplo          # genera los Word en ./salida
```

## Uso con Claude

Entrega a Claude el enlace de `SKILL.md`. Él hará las preguntas iniciales (tema, grado, número de preguntas, dificultad, apoyos visuales), construirá el `exam.json`, ejecutará el script y entregará los dos Word.

## Datos de la institución

El membrete (título, código, versión y fecha del formato) se puede cambiar en el `exam.json`:

```json
"config": { "membrete": { "tituloMembrete": ["…"], "codFormato": "GA-F16", "codVersion": "Versión 1.0", "codFecha": "04-04-2016" } }
```

El escudo se reemplaza cambiando `assets/escudo_header.png` (PNG con proporción similar a 200×242 px).
