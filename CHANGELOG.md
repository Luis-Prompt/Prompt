# Historial de cambios del repositorio

Cada prompt lleva además su propio `CHANGELOG.md`.

## 1.2.0 — Claude lee todo el repositorio con un enlace
- **`herramientas/preparar-prompt.sh`:** un solo comando descarga y deja listo el prompt elegido (scripts, catálogos, 20 áreas, escudo), verifica que no falte nada ni haya páginas «404» e instala `docx`/`sharp` si hacen falta. Reemplaza la descarga archivo por archivo, que queda como respaldo.
- **`manifest.json` por prompt:** mapa completo (qué leer siempre, qué leer según la necesidad, qué se ejecuta). El verificador exige que todo archivo de la carpeta esté mapeado y que existan los listados.
- **`llms.txt` autogenerado** (`node herramientas/generar-mapa.js`) con enlaces directos a todo, para Claude sin terminal; el CI falla si está desactualizado.
- **`INDICE-PARA-CLAUDE.md` sección 0** (qué hacer en orden al recibir el enlace), bloque «Para Claude» al inicio del README y `CLAUDE.md`.
- Prueba nueva `npm run test:preparar`: prepara ambos prompts y genera una prueba y una planeación desde la carpeta de trabajo, igual que lo haría Claude.
- Pruebas SABER 2.9.0 y Planeaciones 1.3.0 (solo el arranque; sin cambios en los generadores). Se corrige la versión que mostraba el `SKILL.md` de planeaciones.

## 1.1.0 — CI activo y verificación de calidad
- **El CI ahora corre de verdad:** el workflow pasó de `herramientas/github-actions/` a `.github/workflows/ci.yml`. Se ejecuta en cada push a `main` y en cada Pull Request, con caché de npm, `npm ci`, cancelación de ejecuciones repetidas y permisos de solo lectura.
- El CI ahora también prueba la **plantilla de calificación en Excel** y guarda la planeación de ejemplo (Word y PDF) como artefacto.
- Nuevo `herramientas/verificar-calidad.js` (`npm run verificar:calidad`): JSON válidos, rutas del escudo, enlaces de los `.md`, detección de secretos y límite de tamaño de archivos. Incluido en `npm test`.
- Nuevos: `CONTRIBUTING.md`, `SECURITY.md`, `.editorconfig`, `.gitattributes`, `.nvmrc`, `engines` de Node y Dependabot (Actions y npm).

## 1.0.0 — Biblioteca inicial
- Dos prompts independientes (pruebas tipo SABER y planeaciones de clase), enrutador para Claude, configuración institucional, banco de imágenes y de preguntas, y herramientas compartidas.
