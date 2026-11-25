# Herramientas

Utilidades de apoyo. Cada una trae su modo de uso al inicio del archivo.

## Para usar (docentes y Claude)
| Herramienta | Qué hace | Ejemplo |
|---|---|---|
| `preparar-prompt.sh` | **Un comando** que descarga y deja listo todo un prompt (scripts, catálogos, escudo), verifica que no falte nada y dice qué leer | `bash herramientas/preparar-prompt.sh pruebas-saber` |
| `guardar-en-banco.js` | Guarda las preguntas de un `exam.json` en `recursos/banco-preguntas/` sin duplicarlas | `node herramientas/guardar-en-banco.js exam.json` |
| `armar-desde-banco.js` | **Arma una prueba con preguntas ya aprobadas del banco**, equilibrada por competencia y nivel, sin repetir las de una prueba anterior | `node herramientas/armar-desde-banco.js --area matematicas --grado 5 --n 10` |
| `plan-a-borrador.js` | **Puente de datos:** convierte las preguntas Saber de una planeación en un borrador de preguntas para el prompt de pruebas | `node herramientas/plan-a-borrador.js plan.json` |
| `preparar-escudo.py` | Recorta el escudo y genera maestro, encabezado y miniatura | `python3 herramientas/preparar-escudo.py ESCUDO.png --nombre escudo_colegio_simon_bolivar_2025 --salida recursos/colegio/escudos` |

## Para mantener el repositorio (`mantenimiento/`)
Los ejecuta el CI en cada cambio; también puedes correrlos tú con `npm test`.

| Herramienta | Qué comprueba o genera | Comando |
|---|---|---|
| `verificar-calidad.js` | JSON válidos, rutas del escudo, enlaces de los `.md`, secretos y archivos demasiado pesados | `npm run verificar:calidad` |
| `verificar-indice.js` | Prompts **aislados**, índice bien enrutado y `manifest.json` completo de cada prompt | `npm run verificar` |
| `probar-banco.js` | Prueba todo el ciclo del banco: guardar sin duplicar, armar, excluir, exigir «aprobada» y generar el Word | `npm run test:banco` |
| `generar-mapa.js` | Genera `llms.txt` (mapa con enlaces directos); con `--check` falla si está desactualizado | `npm run mapa` |

La verificación automática (CI) está en [`.github/workflows/ci.yml`](../.github/workflows/ci.yml).
