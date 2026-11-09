# Herramientas compartidas

Utilidades que sirven a más de un prompt. Cada una trae su modo de uso al inicio del archivo.

| Herramienta | Qué hace | Ejemplo |
|---|---|---|
| `preparar-escudo.py` | Recorta el escudo y genera maestro, encabezado y miniatura | `python3 herramientas/preparar-escudo.py ESCUDO.png --nombre escudo_colegio_simon_bolivar_2025 --salida banco-imagenes/escudos` |
| `guardar-en-banco.js` | Guarda las preguntas de un `exam.json` en `banco-preguntas/` sin duplicarlas | `node herramientas/guardar-en-banco.js exam.json` |
| `github-actions/` | Verificación automática (pendiente de activar; ver su `LEEME.md`) | — |
