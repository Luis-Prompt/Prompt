# Catálogo de áreas

Una carpeta de conocimiento por cada área o asignatura. El prompt (`SKILL.md`) lee **solo el archivo del área elegida** por el docente.

| Clave | Área | Marco de referencia |
|---|---|---|
| `matematicas` | Matemáticas | 1-11 |
| `geometria` | Geometría | 1-11 |
| `estadistica` | Estadística y probabilidad | 1-11 |
| `lenguaje` | Lenguaje (Lengua Castellana) | 1-11 |
| `lectura-critica` | Lectura crítica | 9-11 |
| `ingles` | Inglés | 1-11 |
| `ciencias-naturales` | Ciencias naturales | 1-11 |
| `biologia` | Biología | 6-11 |
| `fisica` | Física | 9-11 |
| `quimica` | Química | 9-11 |
| `ciencias-sociales` | Ciencias sociales (historia y geografía) | 1-11 |
| `constitucion-democracia` | Constitución política y democracia | 6-11 |
| `competencias-ciudadanas` | Competencias ciudadanas | 1-11 |
| `filosofia` | Filosofía | 10-11 |
| `etica-valores` | Ética y valores humanos | 1-11 |
| `religion` | Educación religiosa | 1-11 |
| `artistica` | Educación artística | 1-11 |
| `educacion-fisica` | Educación física, recreación y deporte | 1-11 |
| `tecnologia-informatica` | Tecnología e informática | 1-11 |
| `ciencias-economicas-politicas` | Ciencias económicas y políticas | 10-11 |

## Qué trae cada archivo `<clave>.json`
- `marcos`: competencias, componentes y niveles de desempeño según el grado (Saber 11.°, Saber 3.°-5.°-9.°, orientaciones del MEN…).
- `tipos_de_estimulo` y `plantillas_recomendadas`: qué apoyos visuales usar.
- `reglas_de_calidad` y `errores_tipicos_para_distractores`: cómo escribir buenas preguntas en esa área.
- `preguntas_iniciales`: lo que el tutorial le pregunta al docente en esa área.
- `frase_final`, `verificacion` y `ejemplo` (una prueba funcional en `examples/areas/`).

El script valida que `competencia` y `componente` de cada pregunta existan en el catálogo del área (con aviso, no con error).

## Fuentes y advertencia
Las competencias de Matemáticas, Lectura crítica, Ciencias naturales, Sociales y ciudadanas e Inglés salen de las guías del ICFES (Saber 11.° y Saber 3.°, 5.°, 7.° y 9.°); las demás áreas, de las orientaciones del MEN. **Inglés:** los niveles vigentes son Pre A1, A1, A2 y B1 (hasta 2021 se reportaban A-, A1, A2, B1 y B+). Educación religiosa, Ética y Educación artística/física/tecnología usan competencias **adaptadas** para evaluación escrita: ajústalas al plan de área y al PEI de tu colegio.

## Cómo mejorar un área
Edita su `.json` (competencias, reglas, errores típicos, preguntas iniciales), prueba con `npm run probar:areas` y anota el cambio en el `CHANGELOG.md`.
