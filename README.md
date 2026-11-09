# Biblioteca de Prompts — Mg. Luis Alberto Avellaneda C.

Repositorio central para **crear, ordenar y mejorar día a día** los prompts (skills) que uso con Claude para producir material educativo. Cada prompt vive en su propia carpeta con todo lo necesario: instrucciones, scripts, ejemplos, imágenes y registro de cambios.

## Menú de prompts

| Prompt | Para qué sirve | Versión | Carpeta |
|---|---|---|---|
| **Creación de pruebas tipo SABER** | Crea la prueba del estudiante y la del docente (con claves) en Word, con preguntas de selección múltiple verificadas, gráficos y hoja de respuestas. | 2.4 | [`prompts/evaluacion/creacion-pruebas-saber`](prompts/evaluacion/creacion-pruebas-saber) |
| *(próximo)* Planes de clase | Planeación de clases, planeadores, guías y talleres. | — | [`prompts/planeacion`](prompts/planeacion) |

## Cómo está organizado

```
Prompt/
├─ prompts/            Los prompts, agrupados por categoría
│  ├─ evaluacion/      Pruebas, exámenes, rúbricas  →  creacion-pruebas-saber/
│  ├─ planeacion/      Planes de clase, planeadores, guías (próximamente)
│  └─ _plantilla/      Molde para crear un prompt nuevo con el mismo orden
├─ config/             Datos de la institución (nombre, escudo, membrete) que usan todos los prompts
├─ banco-imagenes/     Escudos, plantillas gráficas y recursos visuales reutilizables
├─ banco-preguntas/    Preguntas aprobadas por área y grado (se reutilizan y mejoran)
├─ herramientas/       Utilidades compartidas: preparar-escudo, guardar-en-banco…
├─ docs/               Guías y catálogo
└─ .github/            Plantillas para anotar ideas y errores (Issues)
```

## Cómo se usa un prompt

1. Entra a la carpeta del prompt y lee su `README.md`.
2. Entrega a Claude el enlace de su `SKILL.md` (o súbelo como Skill).
3. Claude hace las preguntas iniciales, guía el proceso y entrega los archivos.

## Cómo se mejora un prompt (flujo diario)

1. Anota la idea en **Issues** (hay plantilla) o en el `ROADMAP.md` del prompt.
2. Haz el cambio en una **rama**, nunca directo en `main`.
3. Actualiza el `CHANGELOG.md` del prompt y súbele la versión.
4. Prueba que el ejemplo siga generándose (`npm run ejemplo`). La verificación automática está lista en `herramientas/github-actions/`, pendiente de activar.
5. Se revisa y se une a `main` con un *Pull Request*.

## Reglas de orden

- Cada prompt va en su propia carpeta, dentro de una categoría: `prompts/<categoria>/<nombre-del-prompt>`.
- Nombres de carpeta en minúsculas, sin tildes ni espacios, con guiones: `creacion-pruebas-saber`.
- Un `SKILL.md` **corto y claro**; el código va en `scripts/`, nunca pegado dentro del prompt.
- Cada prompt trae un ejemplo funcional en `examples/`.
- No subir claves, tokens ni datos personales de estudiantes.
