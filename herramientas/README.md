# Herramientas compartidas

Utilidades que sirven a más de un prompt. Cada una trae su modo de uso al inicio del archivo.

| Herramienta | Qué hace | Ejemplo |
|---|---|---|
| `preparar-escudo.py` | Recorta el escudo y genera maestro, encabezado y miniatura | `python3 herramientas/preparar-escudo.py ESCUDO.png --nombre escudo_colegio_simon_bolivar_2025 --salida banco-imagenes/escudos` |
| `guardar-en-banco.js` | Guarda las preguntas de un `exam.json` en `banco-preguntas/` sin duplicarlas | `node herramientas/guardar-en-banco.js exam.json` |
| `plan-a-borrador.js` | **Puente de datos:** convierte las preguntas Saber de una planeación en un borrador de preguntas para el prompt de pruebas | `node herramientas/plan-a-borrador.js plan.json` |
| `verificar-indice.js` | Comprueba que los prompts estén **aislados** (cada uno con su frontera y sus archivos) y que el índice los enrute bien | `node herramientas/verificar-indice.js` |
| `verificar-calidad.js` | Revisa la salud del repo: JSON válidos, rutas del escudo, enlaces de los `.md`, secretos y archivos demasiado pesados | `node herramientas/verificar-calidad.js` |

La verificación automática (CI) está en [`.github/workflows/ci.yml`](../.github/workflows/ci.yml) y se ejecuta sola en cada cambio.
