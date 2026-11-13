# Biblioteca de Prompts — Mg. Luis Alberto Avellaneda C.

Repositorio central para **crear, ordenar y mejorar día a día** los prompts (skills) que se usan con Claude para producir material educativo del **Colegio Integrado Simón Bolívar (Cúcuta)**. Cada prompt tiene su propia carpeta con todo lo que necesita: instrucciones, scripts, ejemplos y registro de cambios.

## 🚀 Empezar (docentes)
Abre Claude, copia **un enlace** de la tabla y escribe: *«Usa este prompt: [enlace]»*. Claude te guiará paso a paso.

| Prompt | Para qué sirve | Versión | Enlace para copiar |
|---|---|---|---|
| 📝 **Creación de pruebas tipo SABER** | Prueba del estudiante y del docente (con claves) en Word para **20 áreas**, con textos compartidos, gráficos, ilustraciones con ChatGPT y hoja de respuestas | 2.6 | [`SKILL.md`](https://github.com/Luis-Prompt/Prompt/blob/main/prompts/evaluacion/creacion-pruebas-saber/SKILL.md) |
| 📅 **Planeaciones de clase** | Planeación completa del periodo (9 semanas) en **Word horizontal y PDF**, con tutorial por rondas | 1.0 | [`SKILL.md`](https://github.com/Luis-Prompt/Prompt/blob/main/prompts/planeacion/planeaciones-de-clase/SKILL.md) |

## 🗺️ Mapa del repositorio
```
Prompt/
├─ README.md                  ← este índice
├─ prompts/                   Los prompts, uno por carpeta y agrupados por categoría
│  ├─ evaluacion/
│  │  └─ creacion-pruebas-saber/     SKILL.md · scripts · areas (20) · examples
│  ├─ planeacion/
│  │  └─ planeaciones-de-clase/      SKILL.md · scripts · catalogo · examples
│  └─ _plantilla/                    molde para un prompt nuevo
├─ config/                    Datos del colegio (nombre, escudo, membretes) compartidos por todos
├─ banco-imagenes/            Escudo en 3 tamaños y recursos visuales
├─ banco-preguntas/           Preguntas aprobadas por área y grado
├─ herramientas/              Utilidades compartidas (preparar escudo, guardar en banco, CI)
└─ docs/                      Guías de uso
```

## 🧩 Qué comparten los prompts
- **`config/institucion.json`**: nombre del colegio, lema, rutas del escudo y membretes (de pruebas y de planeación). Cambiarlo aquí cambia todos los documentos.
- **`banco-imagenes/escudos/`**: el escudo 2025 en maestro, encabezado y miniatura.
- **`herramientas/`**: `preparar-escudo.py` y `guardar-en-banco.js`.

## 🛠️ Cómo se mejora un prompt (flujo diario)
1. Anota la idea en **Issues** (hay plantillas) o en el `ROADMAP.md` del prompt.
2. Haz el cambio en una **rama**, nunca directo en `main`.
3. Actualiza el `CHANGELOG.md` del prompt y sube su versión.
4. Prueba con el ejemplo del prompt (`npm run ejemplo`).
5. Únelo a `main` con un *Pull Request*.

## 📐 Reglas de orden
- Cada prompt va en `prompts/<categoria>/<nombre-del-prompt>` (minúsculas, sin tildes ni espacios).
- El `SKILL.md` es **corto y claro**: el conocimiento va en `catalogo/` o `areas/` y el código en `scripts/`.
- Cada prompt trae un ejemplo funcional en `examples/`.
- No subir claves, tokens ni datos personales de estudiantes.

Guía rápida para docentes: [`docs/GUIA-RAPIDA.md`](docs/GUIA-RAPIDA.md).
