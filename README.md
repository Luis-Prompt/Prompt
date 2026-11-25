# Biblioteca de Prompts — Mg. Luis Alberto Avellaneda C.

[![CI](https://github.com/Luis-Prompt/Prompt/actions/workflows/ci.yml/badge.svg)](https://github.com/Luis-Prompt/Prompt/actions/workflows/ci.yml)
![Node](https://img.shields.io/badge/node-%E2%89%A520-339933)
![Prompts](https://img.shields.io/badge/prompts-2-blue)

Repositorio central para **crear, ordenar y mejorar día a día** los prompts (skills) que se usan con Claude para producir material educativo del **Colegio Integrado Simón Bolívar (Cúcuta)**. Cada prompt tiene su propia carpeta con todo lo que necesita: instrucciones, scripts, ejemplos y registro de cambios.

## 🤖 Si le pasas este repositorio a Claude
**Para Claude:** lee primero [`INDICE-PARA-CLAUDE.md`](INDICE-PARA-CLAUDE.md) (sección 0: qué hacer, en orden) o el mapa completo con enlaces directos en [`llms.txt`](llms.txt). Prepara el prompt elegido con un solo comando:
`curl -fsSL https://raw.githubusercontent.com/Luis-Prompt/Prompt/main/herramientas/preparar-prompt.sh | bash -s -- <pruebas-saber|planeaciones-de-clase>`

Este repositorio tiene **prompts independientes**; cada uno hace **una sola cosa** y **no se mezcla con los demás**. Claude elige uno según lo que pidas, lee solo el `SKILL.md` de ese prompt y ignora el resto. Las reglas están en [`INDICE-PARA-CLAUDE.md`](INDICE-PARA-CLAUDE.md) (versión para máquinas: [`prompts/INDICE.json`](prompts/INDICE.json)).

## 🚀 Empezar (docentes)
Abre Claude y copia **una** de estas frases (según lo que necesites):

| Quiero… | Escribe en Claude |
|---|---|
| 📝 una **prueba de comprensión** | *Usa el repositorio https://github.com/Luis-Prompt/Prompt para crear una prueba de comprensión.* |
| 📅 una **planeación de clase** | *Usa el repositorio https://github.com/Luis-Prompt/Prompt para crear una planeación de clase.* |

También puedes pegar directamente el enlace de un solo prompt:

| Prompt | Para qué sirve (y para qué NO) | Versión | Enlace para copiar |
|---|---|---|---|
| 📝 **Creación de pruebas tipo SABER** | Prueba del estudiante y del docente (con claves) en Word para **20 áreas**, con textos compartidos, gráficos, ilustraciones con ChatGPT y hoja de respuestas. *No crea planeaciones.* | 2.9 | [`SKILL.md`](https://github.com/Luis-Prompt/Prompt/blob/main/prompts/evaluacion/creacion-pruebas-saber/SKILL.md) |
| 📅 **Planeaciones de clase** | Planeación completa del periodo (9 semanas) en **Word horizontal y PDF**, con tutorial por rondas. *No crea pruebas.* | 1.3 | [`SKILL.md`](https://github.com/Luis-Prompt/Prompt/blob/main/prompts/planeacion/planeaciones-de-clase/SKILL.md) |

## 🗺️ Qué hay en este repositorio
```
Prompt/
├─ README.md                 ← estás aquí
├─ INDICE-PARA-CLAUDE.md     ← para Claude: qué hacer al recibir el enlace y cómo elegir UN prompt
├─ llms.txt                  ← para Claude: mapa completo con enlaces directos (se genera solo)
├─ CLAUDE.md · package.json  ← acceso rápido para Claude Code · comando `npm test`
│
├─ prompts/                  🎯 LO IMPORTANTE: los prompts, cada uno independiente
│  ├─ evaluacion/creacion-pruebas-saber/   📝 pruebas tipo SABER (20 áreas)
│  ├─ planeacion/planeaciones-de-clase/    📅 planeaciones del periodo (Word + PDF)
│  └─ _plantilla/                          molde para crear un prompt nuevo
│
├─ recursos/                 🏫 Datos compartidos (no son instrucciones)
│  ├─ colegio/               institucion.json (nombre, membretes) · escudos/ · perfil-docente.ejemplo.json
│  └─ banco-preguntas/       preguntas aprobadas por área y grado
│
├─ herramientas/             🔧 Utilidades: preparar-prompt.sh · guardar-en-banco · plan-a-borrador · preparar-escudo
│  └─ mantenimiento/         verificadores y generador de llms.txt (los usa el CI)
│
├─ docs/                     📚 Guía rápida para docentes · ideas · historial del repositorio
└─ .github/                  ⚙️ CI, plantillas de issues y PR, CONTRIBUTING y SECURITY
```
Cada carpeta tiene su propio `README.md`. Cada prompt trae: `SKILL.md` (las instrucciones), `manifest.json` (mapa de sus archivos), `scripts/`, sus catálogos (`areas/` o `catalogo/`), `examples/`, `CHANGELOG.md` y `ROADMAP.md`.

## 🧩 Qué comparten los prompts
- **`recursos/colegio/institucion.json`**: nombre del colegio, lema, rutas del escudo y membretes (de pruebas y de planeación). Cambiarlo aquí cambia todos los documentos.
- **`recursos/colegio/escudos/`**: el escudo 2025 en maestro, encabezado y miniatura (una sola copia).
- **`herramientas/`**: `preparar-prompt.sh` (deja listo un prompt con un comando), `guardar-en-banco.js`, `plan-a-borrador.js` y `preparar-escudo.py`.

## 🛠️ Cómo se mejora un prompt (flujo diario)
1. Anota la idea en **Issues** (hay plantillas) o en el `ROADMAP.md` del prompt.
2. Haz el cambio en una **rama**, nunca directo en `main`.
3. Actualiza el `CHANGELOG.md` del prompt y sube su versión.
4. Corre **`npm test`** (revisa la calidad del repo, verifica el aislamiento de los prompts, genera los ejemplos y comprueba la calificación). El **CI** lo repite solo en cada Pull Request.
5. Únelo a `main` con un *Pull Request*.

## 📐 Reglas de orden
- Cada prompt va en `prompts/<categoria>/<nombre-del-prompt>` (minúsculas, sin tildes ni espacios).
- El `SKILL.md` es **corto y claro**: el conocimiento va en `catalogo/` o `areas/` y el código en `scripts/`.
- Cada prompt trae un ejemplo funcional en `examples/` y un `manifest.json` que mapea **todos** sus archivos (el CI lo comprueba).
- Si agregas o quitas archivos o prompts, ejecuta `node herramientas/mantenimiento/generar-mapa.js` para actualizar `llms.txt`.
- No subir claves, tokens ni datos personales de estudiantes.

Cómo contribuir: [`.github/CONTRIBUTING.md`](.github/CONTRIBUTING.md) · Seguridad: [`.github/SECURITY.md`](.github/SECURITY.md) · Historial: [`docs/CHANGELOG.md`](docs/CHANGELOG.md)

Guía rápida para docentes: [`docs/GUIA-RAPIDA.md`](docs/GUIA-RAPIDA.md) · Ideas para seguir mejorando: [`docs/IDEAS.md`](docs/IDEAS.md).
