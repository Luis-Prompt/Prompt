# Historial de cambios del repositorio

Cada prompt lleva además su propio `CHANGELOG.md`.

## 1.1.0 — CI activo y verificación de calidad
- **El CI ahora corre de verdad:** el workflow pasó de `herramientas/github-actions/` a `.github/workflows/ci.yml`. Se ejecuta en cada push a `main` y en cada Pull Request, con caché de npm, `npm ci`, cancelación de ejecuciones repetidas y permisos de solo lectura.
- El CI ahora también prueba la **plantilla de calificación en Excel** y guarda la planeación de ejemplo (Word y PDF) como artefacto.
- Nuevo `herramientas/verificar-calidad.js` (`npm run verificar:calidad`): JSON válidos, rutas del escudo, enlaces de los `.md`, detección de secretos y límite de tamaño de archivos. Incluido en `npm test`.
- Nuevos: `CONTRIBUTING.md`, `SECURITY.md`, `.editorconfig`, `.gitattributes`, `.nvmrc`, `engines` de Node y Dependabot (Actions y npm).

## 1.0.0 — Biblioteca inicial
- Dos prompts independientes (pruebas tipo SABER y planeaciones de clase), enrutador para Claude, configuración institucional, banco de imágenes y de preguntas, y herramientas compartidas.
