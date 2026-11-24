# Cómo contribuir

¡Gracias por ayudar a mejorar la biblioteca! Este flujo mantiene `main` siempre funcionando.

## Flujo de trabajo
1. Anota la idea en **Issues** (hay plantillas) o en el `ROADMAP.md` del prompt.
2. Crea una **rama** (`mejora/<tema>`); nunca trabajes directo en `main`.
3. Haz commits pequeños con mensajes claros (`Pruebas: agrega plantilla de histograma`).
4. Actualiza el `CHANGELOG.md` del prompt y sube su versión (SemVer: `MAYOR.MENOR.PARCHE`).
5. Corre **`npm test`** en la raíz (requiere Node 20+, Python 3 con `openpyxl` y LibreOffice).
6. Abre un **Pull Request**: el CI debe quedar en verde antes de unirlo.

## Reglas de orden
- Cada prompt vive en `prompts/<categoria>/<nombre>` (minúsculas, sin tildes ni espacios) y hace **una sola cosa**.
- El `SKILL.md` debe ser corto: el conocimiento va en `catalogo/`, `areas/` o `referencia/`, y el código en `scripts/`.
- Todo prompt trae un ejemplo funcional en `examples/` y un `CHANGELOG.md`.
- Los datos del colegio van en `config/institucion.json`, no dentro del código.
- Una sola copia de cada recurso: el escudo vive en `banco-imagenes/escudos/`.
- Preguntas nuevas: guárdalas en `banco-preguntas/` con estado `borrador` hasta que un docente las revise.

## Antes de subir
```bash
npm run verificar:calidad   # JSON, enlaces, rutas, secretos y tamaños
npm run verificar           # aislamiento de los prompts
npm test                    # todo lo anterior + generación de ejemplos
```
Ver también [`SECURITY.md`](SECURITY.md).
