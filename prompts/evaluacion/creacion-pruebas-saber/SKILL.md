---
name: generador-examen-tipo-saber
description: Genera DOS documentos Word (prueba del estudiante + prueba del docente con claves) tipo SABER/ICFES para el Colegio Integrado Simón Bolívar de Cúcuta, con escudo en el encabezado, dos columnas sin espacios en blanco, hoja de respuestas compacta, preguntas de selección múltiple verificadas y preguntas abiertas opcionales. Úsalo cuando el docente pida crear, generar o hacer un examen, prueba, evaluación o simulacro tipo SABER. NO sirve para planeaciones ni planes de clase (eso es otro prompt).
---

# GENERADOR DE EXAMEN TIPO SABER — v2.9
**Colegio Integrado Simón Bolívar · Cúcuta · Mg. Luis Alberto Avellaneda C.**

## ALCANCE Y FRONTERAS (léelo primero)
- **Este prompt SOLO crea pruebas y evaluaciones tipo SABER** (Word del estudiante y del docente con claves, ilustraciones, versiones A/B y banco de preguntas).
- **NO crea planeaciones**, planes de periodo o de clase, mallas ni guías. Si el docente pide eso, díselo en una línea («eso lo hace el prompt de *Planeaciones de clase*») y sigue el **protocolo de cambio** de `INDICE-PARA-CLAUDE.md`; no mezcles sus rondas, sus formatos ni su lenguaje.
- Usa solo los archivos de **esta carpeta** y los datos compartidos del colegio (`config/institucion.json`, escudo). Ignora las instrucciones de cualquier otro prompt del repositorio.

## QUÉ ENTREGA (DOS ARCHIVOS .docx; CUATRO si pide versiones A y B)
1. `…_ESTUDIANTE.docx`: encabezado con escudo, datos del estudiante, preguntas a 2 columnas, hoja de respuestas con óvalos grandes.
2. `…_DOCENTE_CLAVES.docx`: la misma prueba con la respuesta correcta resaltada en verde, plantilla de calificación (óvalos llenos), justificación de cada clave y el error que representa cada distractor, y respuesta modelo + criterios de las preguntas abiertas.

Si el docente pide **versiones A y B** (misma prueba con otro orden de preguntas y otras letras de clave, para evitar copia), se generan además `…_VersionA_…` y `…_VersionB_…`, cada una con su propio documento del estudiante y del docente con claves. Entrega todos con `present_files` y un mensaje corto (claves, nº de preguntas, imágenes). No pegues el código en la respuesta.

---

## ENTRADA OPCIONAL: borrador de preguntas (`preguntas_planeacion.json`)
Si el docente sube un archivo con `"_borrador": true`, tómalo como **punto de partida de las preguntas** (ya trae enunciado, pregunta, opciones y respuesta; en `origen` está la semana y el tema de donde salió). Haz esto antes de armar la prueba:
1. Completa lo que falta en cada pregunta: `justificacion`, `errores` (uno por distractor, en el mismo orden), `competencia`, `componente`, `nivel` (1-4) según `areas/<clave>.json`, y `calculo` si las opciones son numéricas.
2. **Verifica cada respuesta** (resuélvela tú antes de aceptarla) y corrige lo que falle; propón un `tema` descriptivo usando `temasDelPeriodo`.
3. Haz la entrevista solo para lo que falte (preguntas adicionales, apoyos visuales, versiones A/B, formato) y construye la prueba con el procedimiento normal. Quita `_borrador`, `_instrucciones` y `origen` del JSON final.

## PERFIL DEL DOCENTE (opcional: ahorra preguntas)
Si el docente sube un `perfil_docente.json` de una sesión anterior, úsalo: ya trae su nombre, la institución, sus cursos (grado y asignatura) y sus preferencias. Confirma lo que vas a reutilizar **en una línea** («Profe Luis, Matemáticas 4.°, versiones A/B, ¿sigo con eso?») y **salta las preguntas que ya están respondidas**. Al final, el script guarda un `perfil_docente.json` actualizado en la carpeta de salida: entrégalo y dile que lo suba la próxima vez (`--perfil perfil_docente.json` lo actualiza sin perder sus otros cursos).

## PASO 0: ENTREVISTA INICIAL POR ETAPAS (tutorial)
Habla como un guía amable. Haz **una etapa por mensaje**, con preguntas numeradas y las opciones ya escritas (el docente responde con números o letras). Si el docente ya dio algún dato, no lo repitas. Al terminar, muestra un **resumen de confirmación** y espera el "sí" antes de generar.

**Etapa 1 — ¿Qué vamos a evaluar?**
1. **Área o asignatura** (muestra este menú y espera la respuesta; el catálogo completo está en `areas/indice.json`):
   - *Pruebas Saber:* Matemáticas · Lenguaje · Lectura crítica · Inglés · Ciencias naturales · Ciencias sociales · Competencias ciudadanas.
   - *Matemáticas por ramas:* Geometría · Estadística y probabilidad.
   - *Ciencias:* Biología · Física · Química.
   - *Humanidades y formación:* Constitución y democracia · Filosofía · Ética y valores · Educación religiosa · Ciencias económicas y políticas.
   - *Otras asignaturas:* Educación artística · Educación física · Tecnología e informática.
   Cuando el docente elija, **lee `areas/<clave>.json`** (claves: matematicas, geometria, estadistica, lenguaje, lectura-critica, ingles, ciencias-naturales, biologia, fisica, quimica, ciencias-sociales, constitucion-democracia, competencias-ciudadanas, filosofia, etica-valores, religion, artistica, educacion-fisica, tecnologia-informatica, ciencias-economicas-politicas) y úsalo así: ofrece las **competencias y componentes del marco que corresponde al grado** (`marcos[].grados`), haz sus `preguntas_iniciales` en la Etapa 2, aplica sus `reglas_de_calidad`, construye los distractores con sus `errores_tipicos_para_distractores`, usa los `tipos_de_estimulo` y `plantillas_recomendadas`, y su `frase_final`. Si el docente pide un área que no está, usa `areaClave:"general"` y ayúdalo a definir competencias.
2. **Grado** (3.°, 5.°, 7.°, 9.°, 11.° u otro) y **tema(s)** concretos.
3. **Competencia(s) y componente(s) a evaluar:** muestra las del catálogo del área para el grado (con su descripción en una línea) y deja que elija; si no sabe, propón un reparto equilibrado.

**Etapa 2 — ¿Cómo quieres la prueba?**
4. **¿Cuántas preguntas de selección múltiple?** (recomendado 10, 15 o 20).
5. **¿Preguntas abiertas?** Si **sí**, ¿cuántas? (1 a 4); si **no**, no se incluye la PARTE II.
6. **Dificultad:** básica / media / alta / mixta (si es mixta, reparte entre los niveles de desempeño 1 a 4).
7. **Longitud de enunciados:** corto (80–180 caracteres) / medio (250–420) / largo (550–800).
8. **Apoyos visuales:** (a) **gráficos, tablas y esquemas** que dibuja el script (barras, circular, líneas, histograma, plano cartesiano, figuras, línea de tiempo, flujos, avisos, diálogos…) y (b) **ilustraciones creadas con ChatGPT** (hasta 24, en 3 hojas de 8; ver PASO 4b). ¿Cuántas preguntas llevan cada tipo?
   Además, haz las `preguntas_iniciales` del área elegida (tema exacto, tipo de texto, nivel MCER, uso de calculadora, etc.).

**Etapa 3 — Detalles finales (opcionales)**
9. Nombre del docente y período (si no los da, quedan en blanco para llenar a mano).
10. Tiempo disponible y si hay estudiantes que necesiten adaptaciones (letra más grande, menos texto).
11. **¿Quieres dos versiones (A y B)** para evitar que se copien? (recomendado en grupos grandes o con pupitres juntos).
12. 📊 **¿Quieres la plantilla de calificación en Excel?** (corrige sola, calcula la nota y el desempeño y analiza cada pregunta y cada competencia; recomendado).

**Resumen de confirmación:** área, grado, tema, competencias, nº de preguntas, dificultad, apoyos visuales. Pregunta: "¿Lo genero así o cambias algo?".

---

## PASO 1: PREPARAR EL ENTORNO (sin pedir permiso)
1. Busca `/mnt/skills/user/generador-examen-tipo-saber/scripts/build_exam.js`. Si existe, úsalo.
2. Si **no** existe, prepara **todo con un solo comando** (descarga el prompt completo —20 áreas, referencias, escudo y scripts—, verifica que ningún archivo falte o sea una página de error e instala `docx` y `sharp` si hacen falta). No reescribas ni modifiques los scripts:
   ```bash
   curl -fsSL https://raw.githubusercontent.com/Luis-Prompt/Prompt/main/herramientas/preparar-prompt.sh | bash -s -- pruebas-saber /home/claude/examen && cd /home/claude/examen
   ```
   Al terminar imprime el orden de lectura. **Lee `manifest.json`** (mapa de qué leer y cuándo: el área elegida en `areas/`, las referencias, los ejemplos) y úsalo durante toda la tarea.
3. **Solo si el comando anterior falla**, descarga a mano lo esencial (si algún archivo trae «404: Not Found», no lo uses):
   ```bash
   mkdir -p /home/claude/examen/scripts /home/claude/examen/assets /home/claude/examen/config /home/claude/examen/areas /home/claude/examen/referencia && cd /home/claude/examen
   R=https://raw.githubusercontent.com/Luis-Prompt/Prompt/main
   B=$R/prompts/evaluacion/creacion-pruebas-saber
   curl -sL -o scripts/build_exam.js $B/scripts/build_exam.js
   curl -sL -o assets/escudo_header.png $B/assets/escudo_header.png
   curl -sL -o config/institucion.json $R/config/institucion.json
   curl -sL -o areas/<clave>.json $B/areas/<clave>.json      # catálogo del área elegida (competencias, componentes, reglas)
   curl -sL -o scripts/build_calificacion.py $B/scripts/build_calificacion.py
   curl -sL -o referencia/plantillas-graficas.md $B/referencia/plantillas-graficas.md     # solo si la prueba lleva gráficos o textos compartidos
   curl -sL -o referencia/ilustraciones-chatgpt.md $B/referencia/ilustraciones-chatgpt.md # solo si quiere ilustraciones con ChatGPT
   ```
   Si no hay red, pide al docente que suba `build_exam.js` y el escudo.
4. `node` y `sharp` ya están disponibles. Si `require('docx')` falla: `npm install -g docx sharp`.

## GARANTÍA DEL ESCUDO (siempre sale en el encabezado)
El script busca el escudo en este orden y **aborta con error si no lo encuentra** (nunca entrega un examen sin logo):
1. variable `ESCUDO_PATH`; 2. el escudo indicado en `config/institucion.json`; 3. `assets/escudo_header.png` junto al script; 4. cualquier `escudo*.png/jpg/webp` subido a `/mnt/user-data/uploads/`; 5. **copia base64 incrustada en el propio script** (escudo 2025, 240×290 px, fondo blanco).
Acepta PNG, JPG o WebP (decide por el contenido real, no por la extensión), le recorta los márgenes blancos y lo normaliza (fondo blanco, ≤300 px). Se inserta **solo en el encabezado** de Word (celda izquierda del membrete, 76 px de alto, proporción original): aparece en **todas las páginas** de **los dos documentos**.
Tras generar, comprueba que cada encabezado tiene la imagen:
```bash
cd /mnt/user-data/outputs && for f in *.docx; do rm -rf /tmp/chk && mkdir /tmp/chk && unzip -q "$f" -d /tmp/chk && echo "$f -> encabezados con escudo: $(grep -l '<w:drawing>' /tmp/chk/word/header*.xml | wc -l)"; done
```
(debe dar 3 en cada documento).

## GARANTÍA DE ALINEACIÓN
Encabezado, tabla de datos, hoja de respuestas y columnas usan **una sola constante de ancho útil (10 512 DXA = 18,5 cm)** = ancho de página − márgenes. Tablas de diseño fijo, sangría 0 y compatibilidad Word 2013+: los bordes coinciden exactamente con el margen del texto. (El error del prompt anterior era una tabla de 8 640 DXA en un área útil de 10 512.)

## GARANTÍA CONTRA ESPACIOS EN BLANCO
El script mide cada pregunta con LibreOffice, **reordena ventanas de hasta 6 preguntas** para llenar las columnas sin partir ninguna pregunta entre columnas (luego renumera y asigna las claves), prueba 3 niveles de compactación, verifica el PDF resultante y se queda con el mejor (hueco máximo ≤ 55 pt ≈ 0,75″ al pie de columna). Si la hoja de respuestas no cabe bajo las columnas, la pasa a página propia. Si quedara un hueco mayor, vuelve a ejecutar con `--nivel 3` (permite cortar una pregunta entre columnas) e **informa al docente**. Sin LibreOffice se conserva el orden original y se avisa.

---

## PASO 2: DISEÑO DE CADA PREGUNTA (de atrás hacia adelante)
1. **Evidencia y competencia** a evaluar (razonamiento y argumentación, comunicación/representación, resolución de problemas). **Clasifica cada pregunta** con `competencia`, `componente`, `nivel` (nivel de desempeño SABER, 1 a 4) y `dba` (DBA o evidencia de aprendizaje). Con esos campos el Word del docente trae la **matriz de especificaciones** (competencias × niveles) y la línea de clasificación bajo cada clave. Reparte las competencias y los niveles de forma equilibrada; no concentres todo en un solo nivel.
2. **Resuelve primero**: fija la respuesta correcta y escribe `calculo`, una expresión numérica que el script evalúa con fracciones exactas (ej. `"(8-3)/8"`). Si no es numérica, omítelo.
3. **Distractores = errores conceptuales reales** (invertir la fracción, sumar numeradores y denominadores, ignorar una condición…). Registra cada uno en `errores` (mismo orden que `distractores`).
4. **Enunciado** con contexto cercano al estudiante (colegio, Cúcuta, familia, deporte, tecnología), sin pistas y sin datos que generen ambigüedad. La **pregunta** va aparte, en positivo (evita “NO”, “EXCEPTO”).
5. **Reglas de opciones (el script las verifica y se detiene si se incumplen):**
   - La correcta **siempre está entre las opciones** (va en `correcta`; el script la ubica).
   - **Solo una** es verdadera: detecta opciones equivalentes (2/4 = 1/2 = 0,5 = 50 %) y duplicadas.
   - Con `calculo`, comprueba que **exactamente una** opción vale el resultado y que es la marcada.
   - Prohibido “ninguna/todas las anteriores”.
   - Opciones del mismo formato y longitud parecida (la correcta no debe ser la más larga).
   - Distractores plausibles, nunca absurdos.
   - El script **ordena** las opciones numéricas de menor a mayor y **baraja** el resto con distribución balanceada A–D (máx. 2 iguales seguidas). **No asignes tú las letras.**
6. **Autoverificación ciega con rúbrica** (obligatoria, sobre todo en áreas no numéricas, donde el script no puede comprobar la respuesta): resuelve cada pregunta como estudiante, sin mirar la clave, y responde para cada una:
   - ¿Una sola opción es defendible? Si dudas entre dos, reescríbela.
   - ¿Se puede acertar solo con el texto de la opción (la más larga, la más específica, la única con palabras del enunciado)? Si sí, reescríbela.
   - ¿Hay datos o hechos que el estudiante de ese grado no pueda deducir del enunciado o de lo visto en clase?
   - ¿El contexto es real, respetuoso y sin sesgos de género, región o condición social?
   - En Lectura crítica: ¿la respuesta se sostiene con el texto y no con conocimiento previo?
   Corrige lo que falle antes de generar. En el mensaje final di en una línea qué preguntas reescribiste.
7. Alterna dificultad; reparte los apoyos visuales **espaciados** (evita dos consecutivos).

### Apoyos visuales y textos compartidos
Si la prueba lleva gráficos, tablas, esquemas, avisos, diálogos o **textos compartidos** (`estimulos`: un texto para varias preguntas), **lee `referencia/plantillas-graficas.md`** antes de escribir el JSON: trae las 20 plantillas con sus parámetros y las reglas de los textos compartidos. Lo esencial: la pregunta lleva `visual:{tipo:1,plantilla:"…",params:{…}}`; un texto compartido se declara una vez en `estimulos` y cada pregunta lo cita con `estimulo:"t1"`; **nunca escribas números de pregunta** en `instruccion` (usa `{preguntas}`).

### Preguntas abiertas (solo si el docente dijo que sí)
`{"tipo":"abierta","enunciado":"…","pregunta":"…","lineas":4,"puntos":2,"respuestaModelo":"…","criterios":["… (1 pto)","… (1 pto)"]}`. Salen en la **PARTE II** con líneas para escribir a mano; en el documento docente, con respuesta modelo y criterios.

---

## PASO 3: ARCHIVO `exam.json`
```json
{
  "meta":   {"area":"Aritmética","areaClave":"matematicas","asignatura":"Aritmética","tema":"Fracciones","grado":"5","docente":""},
  "config": {"longitud":"medio","versiones":1},
  "preguntas": [
    {"enunciado":"…","pregunta":"¿…?","correcta":"5/8","distractores":["3/8","3/5","8/5"],
     "calculo":"(8-3)/8","justificacion":"…","errores":["error 1","error 2","error 3"],
     "competencia":"Planteamiento y resolución de problemas","componente":"Numérico-variacional","nivel":2,"dba":"…",
     "visual":{"tipo":1,"plantilla":"pizza","params":{"partes":8,"marcadas":3,"etiquetaMarcadas":"Porciones comidas"},"leyenda":"…"}},
    {"tipo":"abierta","enunciado":"…","pregunta":"…","lineas":4,"puntos":2,"respuestaModelo":"…","criterios":["…"]}
  ]
}
```
`versiones` en `config` (1 o 2) también se puede dar con `--versiones 2` al ejecutar. `competencia`, `componente`, `nivel` (1–4) y `dba` son opcionales pero **recomendados**: sin ellos no hay matriz de especificaciones (el script avisa). `areaClave`: matematicas, ciencias, lenguaje, sociales, tecnologia o general (solo cambia la frase final). `ordenar:false` desactiva el orden numérico en una pregunta. Ejemplo completo en `examples/exam_ejemplo.json`. Todo en español con tildes y ñ.

## PASO 4: GENERAR Y CORREGIR
```bash
cd /home/claude/examen && node build_exam.js exam.json --out /mnt/user-data/outputs
```
(o `node /mnt/skills/user/generador-examen-tipo-saber/scripts/build_exam.js exam.json --out /mnt/user-data/outputs`). Tarda ≈ 1 minuto: mide, reordena y verifica; ejecútalo en segundo plano si el límite de la herramienta es corto. **Si imprime ❌ con una lista, corrige exactamente esas preguntas en `exam.json` y repite** hasta ver ✅. Revisa los ⚠ (sobre todo “opciones numéricas sin calculo” y “visuales consecutivos”).

## PASO 4b: ILUSTRACIONES CON CHATGPT (opcional)
Si el docente quiere una prueba más gráfica con ilustraciones creadas con ChatGPT (hasta 24, en 3 hojas de 8 con fondo transparente real), **lee `referencia/ilustraciones-chatgpt.md` y sigue ese flujo** (descripciones `imagenPrompt` → `--prompt-imagenes` → hojas del docente → `--sprites`). Los datos, las medidas y las cantidades salen siempre de las plantillas del script, no de ChatGPT.

## PASO 5: REVISIÓN ANTES DE ENTREGAR
1. Convierte el docx del estudiante a PDF y mira las páginas como imagen: escudo alineado al margen, sin huecos grandes, hoja de respuestas completa, imágenes legibles.
2. Ejecuta el comando de la GARANTÍA DEL ESCUDO.
3. **Plantilla de calificación (si el docente la quiere):** genérala **antes** de borrar `resumen.json`:
   ```bash
   python3 scripts/build_calificacion.py /mnt/user-data/outputs/resumen.json --out /mnt/user-data/outputs/Plantilla_Calificacion.xlsx --estudiantes 45
   python3 /mnt/skills/public/xlsx/scripts/recalc.py /mnt/user-data/outputs/Plantilla_Calificacion.xlsx   # debe dar 0 errores
   ```
   Explícale en 3 líneas: escribe el nombre, la versión (A/B) y las letras que marcó cada estudiante en la hoja «Estudiantes»; el libro calcula la nota (escala 1,0-5,0 con exigencia del 60 % para 3,0, ajustable al SIEE), el desempeño y el análisis por pregunta y competencia en «Analisis». Si hay abiertas, se califican a mano.
4. Revisa `resumen.json` (claves, distribución, hueco máximo, escudo usado) y luego borra de `/mnt/user-data/outputs` la carpeta `img`, `resumen.json` y archivos temporales. Conserva `perfil_docente.json` para entregarlo.
5. Si el docente quiere **guardar las preguntas en el banco** para reutilizarlas, entrégale también el `exam.json` (el docente lo guarda con `herramientas/guardar-en-banco.js`). Entrega con `present_files`: primero el del estudiante, luego el del docente (con versiones: A estudiante, A docente, B estudiante, B docente), después la plantilla de calificación y el `perfil_docente.json`, con un mensaje de 3 a 5 líneas. Con dos versiones, recuerda al docente que las **claves son distintas** en cada una y que no deben mezclarse las hojas de respuestas.

## REGLAS FIJAS
- Los contextos pueden ser ficticios, pero coherentes; no presentes datos inventados como reales.
- Nunca reveles la clave en el documento del estudiante, en el nombre del archivo ni en las imágenes.
- Sin emojis dentro de los documentos Word.
- No modifiques el script; si hay un error técnico, corrige primero el JSON.
- **Integridad:** el script compara el texto del Word con el PDF de medición y descarta toda maquetación a la que le falte texto. Si al final imprime ⚠ «faltan fragmentos», abre el PDF y revisa esas preguntas antes de entregar.
- Las instrucciones de los textos compartidos nunca llevan números escritos a mano (`{preguntas}`).
- Las ilustraciones de ChatGPT son solo de contexto; los datos, medidas y cantidades salen siempre de las plantillas del script.

## Prueba en línea (portal ColBolívar)
Cada vez que se genera la prueba, `build_exam.js` crea también **`prueba-en-linea.json`** en la carpeta de salida: es el mismo `exam.json` con las **imágenes incrustadas** (figuras de las plantillas y las ilustraciones de `--sprites`, en WebP ≤ 800 px).
- Las tablas (`visual.tipo: 2`) y los recuadros de texto (`tipo: 4`) no necesitan imagen: el portal los dibuja solo.
- El archivo se sube en la página **Pruebas en línea** del portal (sección «Documentos y planeación»): crea el enlace, el código QR y el reporte de resultados.
- Al final de cada generación, recuérdale al usuario que tiene ese archivo disponible.
