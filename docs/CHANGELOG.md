# Historial de cambios del repositorio

Cada prompt lleva además su propio `CHANGELOG.md`.

## 1.4.0 — Pruebas armadas desde el banco de preguntas
- **`herramientas/armar-desde-banco.js`:** arma un `exam.json` con preguntas **aprobadas** del banco, repartidas entre competencias y niveles, repetible con `--semilla`, sin repetir las de una prueba anterior (`--excluir`) y con `--listar` para ver qué hay y en qué estado. Las preguntas con texto compartido van juntas.
- **Corrección en `guardar-en-banco.js`:** si una pregunta traía su propio `id`, este podía pisar la huella estable del banco y volver a duplicarla al guardar de nuevo. Ahora la huella manda y el id original se conserva como `idEnLaPrueba`.
- Prueba nueva `npm run test:banco` (también en el CI): guardar sin duplicar, armar, excluir, exigir «aprobada» y generar el Word.

## 1.3.0 — Repositorio ordenado
- **La raíz pasa de 6 carpetas y 8 archivos a 4 carpetas y 5 archivos** (`prompts/`, `recursos/`, `herramientas/`, `docs/`).
- **`recursos/`** reúne los datos compartidos: `colegio/` (antes `config/` y `banco-imagenes/`: institucion.json, perfil de ejemplo y escudos) y `banco-preguntas/`. Un solo README por carpeta en vez de tres sueltos.
- **Una sola copia del escudo:** se eliminan las copias de `assets/` dentro de cada prompt; `preparar-prompt.sh` coloca el escudo en la carpeta de trabajo y los generadores lo buscan en `recursos/colegio/escudos/`.
- **`herramientas/mantenimiento/`** agrupa los verificadores y el generador de `llms.txt` (los usa el CI); en `herramientas/` quedan solo las utilidades de uso.
- `CONTRIBUTING.md` y `SECURITY.md` pasan a `.github/` (GitHub los reconoce ahí) y este historial a `docs/`.
- README con el mapa de carpetas y un README en cada una; `docs/README.md` como índice.
- Rutas que NO cambian: `prompts/...`, `herramientas/preparar-prompt.sh`, `llms.txt` e `INDICE-PARA-CLAUDE.md`, para que los enlaces ya compartidos sigan funcionando.

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
