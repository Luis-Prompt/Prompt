# Historial de cambios — Creación de pruebas tipo SABER

## 2.2.0 — estructura pedagógica
- Cada pregunta admite `competencia`, `componente`, `nivel` (1–4) y `dba`.
- El Word del docente incluye la **matriz de especificaciones** (competencias × niveles de desempeño) y la clasificación bajo cada clave.
- El script valida `nivel` y avisa si faltan competencias.
- Entrevista inicial **por etapas** (tutorial) con resumen de confirmación.
- Autoverificación ciega con rúbrica de 5 puntos, obligatoria en áreas no numéricas.
- La carpeta pasa a `prompts/evaluacion/creacion-pruebas-saber`.

## 2.1.0 — reorganización
- El repositorio pasa de un `.zip` a carpetas versionables.
- `SKILL.md` baja de 86 KB a ~11 KB: se elimina el Anexo A, que era una copia idéntica de `scripts/build_exam.js`.
- Si el script no está instalado como skill, el prompt lo descarga del repositorio.
- Se agregan `package.json`, `README.md`, `ROADMAP.md` y verificación automática (GitHub Actions) lista en `herramientas/github-actions/`, pendiente de activar.
- Sin cambios en la lógica del script: el resultado es idéntico al de la v2.0.

## 2.0.0
- Versión original: dos Word, escudo garantizado, alineación, validaciones y maquetación sin huecos.
