# Guía rápida para docentes

## 1. Usar un prompt
1. Entra a Claude (claude.ai o la app).
2. Escribe **una** de estas frases:
   - *Usa el repositorio https://github.com/Luis-Prompt/Prompt para crear una **prueba de comprensión**.*
   - *Usa el repositorio https://github.com/Luis-Prompt/Prompt para crear una **planeación de clase**.*
3. Claude elige el prompt correcto y **no mezcla** las instrucciones de uno con las del otro. Si tu petición no es clara, te preguntará «¿una prueba o una planeación?».
4. Responde con clics. Al final recibirás tus archivos.

(También puedes pegar el enlace directo del prompt, de la tabla del [README](../README.md).)

## 2. ¿Qué prompt necesito?
| Quiero… | Usa |
|---|---|
| Una **planeación** del periodo | Planeaciones de clase |
| Una **prueba tipo Saber** de cualquier área | Creación de pruebas tipo SABER |
| Una prueba con **ilustraciones** | Pruebas tipo SABER → te entrega un prompt para ChatGPT y luego recorta las imágenes |
| **Devolver los resultados** a cada estudiante (retroalimentación) | Pruebas tipo SABER → con el Excel de calificación ya lleno |

## 3. Después de aplicar la prueba: retroalimentación por estudiante
1. Al crear la prueba, acepta la **plantilla de calificación en Excel**.
2. Cuando los estudiantes presenten la prueba, llena el Excel: nombre, versión (A o B) y las letras que marcó cada uno. Las preguntas abiertas se califican a mano.
3. Vuelve a Claude, sube **ese mismo Excel** y escribe, por ejemplo: *Usa el repositorio https://github.com/Luis-Prompt/Prompt y hazme la **retroalimentación por estudiante** con este Excel.*
4. Recibirás **una hoja por estudiante** (Word y PDF): su nota, sus fortalezas y, en cada pregunta fallada, qué marcó, qué error representa y por qué la correcta lo es. Sirve para devolver la prueba y para reuniones con familias.

Los nombres y resultados de los estudiantes solo se usan para generar esos archivos; no se guardan en el repositorio.

## 4. Consejos
- Ten a mano: grado, periodo, asignatura, los temas del periodo y tu nombre.
- Si Claude te muestra un DBA con **†**, confírmalo con el documento oficial del MEN.
- La ponderación de la evaluación es una propuesta: ajústala al SIEE del colegio.
- Si algo sale mal, dile a Claude qué pasó; también puedes anotar la idea en **Issues** del repositorio.
