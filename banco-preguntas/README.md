# Banco de preguntas

Preguntas guardadas en el mismo formato JSON que usa el generador de pruebas, para reutilizarlas y mejorarlas.

```
banco-preguntas/<area>/<grado>/<tema>.json      ej.: matematicas/5/fracciones.json
```

## Cómo se llena

Después de generar una prueba con Claude, guarda el `exam.json` que te entregue y ejecuta:

```bash
node herramientas/guardar-en-banco.js exam.json                  # entran como "borrador"
node herramientas/guardar-en-banco.js exam.json --estado aprobada
```

- Cada pregunta recibe un **id estable**: si la guardas dos veces no se duplica.
- Estados: `borrador` (recién generada) → `revisada` → `aprobada` (la revisó el docente). Cambia el campo `estado` a mano al revisar.
- Cada pregunta conserva su clasificación (`competencia`, `componente`, `nivel`, `dba`), la justificación y el error que representa cada distractor.

## Regla de calidad
Solo se usan para armar pruebas las preguntas en estado `aprobada`.
