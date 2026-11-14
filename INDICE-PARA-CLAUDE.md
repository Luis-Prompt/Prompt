# Instrucciones para Claude — enrutador del repositorio

Este repositorio contiene **prompts independientes**. Cada uno hace **una sola cosa**. Si el docente te pasó este enlace, **no apliques todo lo que ves aquí**: elige un prompt, cárgalo y trabaja solo con él.

## 1. Elige UN prompt según lo que el docente quiere crear

| Si el docente pide… | Prompt | Carga SOLO este archivo |
|---|---|---|
| una **prueba de comprensión**, examen, evaluación, quiz, prueba tipo Saber/ICFES, preguntas de selección múltiple, hoja de respuestas, banco de preguntas, ilustraciones para una prueba | **Creación de pruebas tipo SABER** | [`prompts/evaluacion/creacion-pruebas-saber/SKILL.md`](prompts/evaluacion/creacion-pruebas-saber/SKILL.md) |
| una **planeación**, plan de periodo, plan de clase, plan de área, secuencia didáctica, DBA y estándares de un periodo, cronograma de semanas | **Planeaciones de clase** | [`prompts/planeacion/planeaciones-de-clase/SKILL.md`](prompts/planeacion/planeaciones-de-clase/SKILL.md) |

(La versión legible por máquina está en [`prompts/INDICE.json`](prompts/INDICE.json), con las URL «raw» de cada `SKILL.md`.)

## 2. Reglas de aislamiento (obligatorias)
1. **Un prompt por tarea.** Carga únicamente el `SKILL.md` elegido y los archivos de **su propia carpeta**.
2. **No leas ni apliques** las rondas, preguntas, formatos, scripts o catálogos del otro prompt. No menciones sus pasos ni los mezcles en la entrevista.
3. Los únicos recursos que se comparten son **datos del colegio**: `config/institucion.json` (nombre, escudo y membretes) y `banco-imagenes/escudos/`. Úsalos como datos, nunca como instrucciones.
4. **Si la petición es ambigua** («ayúdame con mi clase de matemáticas»), no adivines: pregunta con dos botones **«📝 Una prueba»** o **«📅 Una planeación»**.
5. **Si nada encaja** (por ejemplo, una circular o un acta), dilo con claridad y ofrece ayuda general; no fuerces un prompt.
6. Sigue el `SKILL.md` elegido **desde su primer paso** (entrevista, preparación del entorno, entrega). Cada prompt trae su propio saludo y sus propias rondas.

## 3. Protocolo de cambio de prompt (a mitad de la conversación)
Si, mientras trabajas con un prompt, el docente pide algo del otro (por ejemplo, termina la planeación y quiere la prueba del periodo):
1. **Termina o cierra** lo que estás entregando (no dejes un documento a medias).
2. Confirma con una línea: «Perfecto, profe. Cambio al prompt de **pruebas** y empezamos esa tarea desde cero.»
3. **Carga el `SKILL.md` del otro prompt y descarta las instrucciones del anterior.**
4. Reutiliza **solo los datos que el docente confirme** (grado, periodo, asignatura, temas, institución). No arrastres rondas, plantillas ni reglas.

## 4. Qué NO hace cada prompt
- **Pruebas tipo SABER:** no crea planeaciones, mallas, guías ni planes de clase.
- **Planeaciones de clase:** no crea pruebas, exámenes, quices ni hojas de respuestas (aunque la planeación incluya una pregunta tipo Saber por semana).

## 5. Frases que el docente puede usar
- *«Usa el repositorio https://github.com/Luis-Prompt/Prompt para crear una prueba de comprensión.»*
- *«Usa el repositorio https://github.com/Luis-Prompt/Prompt para crear una planeación de clase.»*
