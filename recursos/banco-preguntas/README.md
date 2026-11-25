# Banco de preguntas

Preguntas guardadas en el mismo formato JSON que usa el generador de pruebas, para reutilizarlas y mejorarlas.

```
recursos/banco-preguntas/<area>/<grado>/<tema>.json      ej.: matematicas/5/fracciones.json
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

## Cómo se aprueban
Abre el archivo del tema (por ejemplo `matematicas/5/fracciones.json`), lee cada pregunta (enunciado, clave, justificación y errores de los distractores) y cambia su `"estado"` a `"revisada"` o `"aprobada"`. Puedes ver cuántas hay en cada estado con:
```bash
node herramientas/armar-desde-banco.js --listar
```

## Cómo se arma una prueba desde el banco
```bash
node herramientas/armar-desde-banco.js --area matematicas --grado 5 --n 10 --abiertas 1 --out exam.json
node prompts/evaluacion/creacion-pruebas-saber/scripts/build_exam.js exam.json --out salida
```
- Reparte las preguntas entre las **competencias** y los **niveles** disponibles.
- `--tema fracciones` limita a un tema; `--semilla 7` hace el resultado repetible; sin semilla, cada ejecución da una prueba distinta (útil para versiones A y B de otro grupo).
- `--excluir exam_anterior.json` evita repetir las preguntas de una prueba ya aplicada.
- Las preguntas que comparten un texto o gráfico quedan juntas y el texto viaja con ellas.
- Si no hay suficientes preguntas aprobadas, la herramienta lo dice con claridad (no rellena con borradores). `--estado revisada` o `--estado borrador` relajan la regla.

## Regla de calidad
Por defecto solo se usan para armar pruebas las preguntas en estado `aprobada`.
