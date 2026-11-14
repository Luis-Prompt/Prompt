---
name: generador-examen-tipo-saber
description: Genera DOS documentos Word (prueba del estudiante + prueba del docente con claves) tipo SABER/ICFES para el Colegio Integrado Simón Bolívar de Cúcuta, con escudo en el encabezado, dos columnas sin espacios en blanco, hoja de respuestas compacta, preguntas de selección múltiple verificadas y preguntas abiertas opcionales. Úsalo cuando el docente pida crear, generar o hacer un examen, prueba, evaluación o simulacro tipo SABER. NO sirve para planeaciones ni planes de clase (eso es otro prompt).
---

# GENERADOR DE EXAMEN TIPO SABER — v2.6
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

**Resumen de confirmación:** área, grado, tema, competencias, nº de preguntas, dificultad, apoyos visuales. Pregunta: "¿Lo genero así o cambias algo?".

---

## PASO 1: PREPARAR EL ENTORNO (sin pedir permiso)
1. Busca `/mnt/skills/user/generador-examen-tipo-saber/scripts/build_exam.js`. Si existe, úsalo.
2. Si **no** existe, descarga el script oficial del repositorio (no lo reescribas ni lo modifiques):
   ```bash
   mkdir -p /home/claude/examen/scripts /home/claude/examen/assets /home/claude/examen/config /home/claude/examen/areas && cd /home/claude/examen
   R=https://raw.githubusercontent.com/Luis-Prompt/Prompt/main
   B=$R/prompts/evaluacion/creacion-pruebas-saber
   curl -sL -o scripts/build_exam.js $B/scripts/build_exam.js
   curl -sL -o assets/escudo_header.png $B/assets/escudo_header.png
   curl -sL -o config/institucion.json $R/config/institucion.json
   curl -sL -o areas/<clave>.json $B/areas/<clave>.json      # catálogo del área elegida (competencias, componentes, reglas)
   ```
   Si no hay red, pide al docente que suba `build_exam.js` y el escudo.
3. `node` y `sharp` ya están disponibles. Si `require('docx')` falla: `npm install -g docx sharp`.

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

### Apoyos visuales (campo `visual`)
- **tipo 1: imagen PNG de alta resolución** (1400 px, degradados y sombras, rasterizada). Usa `plantilla` + `params`; las cifras salen de los mismos números del enunciado. **La imagen no debe revelar la respuesta.**

| plantilla | params |
|---|---|
| `pizza` | `partes`, `marcadas`, `estilo` ("torta" opcional), `etiquetaMarcadas`, `etiquetaRestantes?` |
| `barra` (chocolate) | `cols`, `filas`, `marcadas`, `etiquetaMarcadas?`, `etiquetaRestantes?` |
| `cuadricula` | `filas`, `cols`, `sombreadas`, `etiquetaSombreadas?`, `etiquetaRestantes?` |
| `recta` | `divisiones`, `marcas:[{pos,etiqueta}]`, `punto:{pos,etiqueta?}`, `etiquetaInicio`, `etiquetaFin`, `nota?` |
| `vaso` | `nivel:[num,den]`, `etiquetaCapacidad?`, `etiquetaNivel?` |
| `barras` | `titulo`, `categorias:[{nombre,valor,oculto?}]`, `max`, `paso`, `ejeY?` (`oculto:true` dibuja una barra “?” sin revelar el valor) |
| `cancha` | `estilo` (`cancha`/`piscina`/`lote`), `largo`, `ancho` (textos) |
| `huerta` | `surcos`, `conPlantas`, `etiquetaPlantas?`, `etiquetaVacios?` |
| `circular` | `titulo`, `categorias:[{nombre,valor,oculto?,color?}]`, `mostrar` (`porcentaje`/`valor`), `leyendaValores?` (nombres cortos, ≤ 18 caracteres) |
| `lineas` | `titulo`, `x:[…]`, `series:[{nombre,valores:[…]}]`, `min?`, `max?`, `paso?`, `ejeX?`, `ejeY?` (varias series; `null` deja un hueco) |
| `histograma` | `titulo`, `intervalos:[{etiqueta,frecuencia,oculto?}]`, `max?`, `paso?`, `mostrarValores?`, `ejeX?`, `ejeY?` |
| `plano` | `titulo`, `xmin,xmax,ymin,ymax`, `puntos:[{x,y,etiqueta?}]`, `segmentos:[[i,j]]`, `poligono:[i,j,k…]` |
| `figura` | `tipo` (`triangulo_rectangulo`, `triangulo`, `rectangulo`, `cuadrado`, `circulo`, `trapecio`, `paralelogramo`, `poligono`+`lados`), `etiquetas:{…}` (usa `"?"` en la medida que se pide) |
| `angulo` | `rayos:[0,35,90]` (grados, antihorario desde la derecha), `etiquetas:["35°","?"]`, `nota?` |
| `lineaTiempo` | `titulo`, `eventos:[{anio,texto}]` (hasta 6) |
| `flujo` | `titulo`, `nodos:["…",{texto,oculto:true}]`, `modo` (`fila`/`ciclo`/`columna`), `porFila?`, `flechas?`, `cerrar?` (cadenas, ciclos, procesos, algoritmos) |
| `aviso` | `estilo` (`prohibido`/`advertencia`/`informativo`/`seguridad`/`invitacion`/`etiqueta`), `titulo`, `lineas:[…]` (texto discontinuo, señales, carteles) |
| `dialogo` | `mensajes:[{quien,texto,lado?}]` (hasta 5; conversaciones en inglés, lenguaje, ciudadanas) |
| `circuloCromatico` | `titulo`, `ocultos:[índices]`, `nombres?`, `centro?` (con `ocultos` no actives `nombres`: revelaría la respuesta) |
| `svg` | `svg` completo (último recurso; también se rasteriza a PNG) |
- **tipo 2: tabla** `{titulo, headers[], rows[][]}`.
- **tipo 3: gráfico de barras** (`plantilla:"barras"`).
- **tipo 4: cuadro de texto** `{titulo, texto, fuente?}`.
`leyenda` opcional como pie de figura. No se descargan imágenes de internet (en este entorno el acceso externo suele estar bloqueado); se generan localmente en PNG.

### Textos compartidos (`estimulos`) — Lectura, Inglés, Sociales, Ciencias…
Un texto, tabla, gráfico o aviso puede servir a varias preguntas (como en la prueba SABER). Se declara una vez y cada pregunta lo cita:
```json
"estimulos": [{"id":"t1","titulo":"…","texto":"Párrafo 1\nPárrafo 2","fuente":"Texto original","numerarParrafos":true,
               "tipo":"discontinuo","visual":{…},"instruccion":"Responde {preguntas} de acuerdo con el siguiente texto."}],
"preguntas": [{"estimulo":"t1","pregunta":"¿…?","correcta":"…","distractores":["…","…","…"], …}]
```
- La pregunta con `estimulo` **no necesita `enunciado`** (el texto es su contexto). El texto y sus preguntas **siempre van juntos y en orden**, también al armar la versión B.
- **Nunca escribas números de pregunta** en `instruccion`: usa `{preguntas}` (o `{questions}` en inglés). Si no pones `instruccion`, el script la escribe solo («Responde las preguntas 2 a 4 de acuerdo con el siguiente texto.»).
- 2 a 5 preguntas por texto (máx. 6); texto de hasta ~500 palabras; propio o adaptado, sin copiar textos con derechos de autor. Toda respuesta debe poder sostenerse con el texto.
- `visual` puede ser cualquier plantilla (aviso, diálogo, gráfico) y `tipo:"discontinuo"` cambia la frase por «la siguiente información».

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

## PASO 4b: ILUSTRACIONES CON CHATGPT (opcional, cuando el docente quiere una prueba más gráfica)
Las plantillas del script dibujan datos y esquemas exactos. Para **ilustraciones de contexto** (animales, objetos, escenas, personas, lugares) el docente las crea con ChatGPT en **hojas de 8 imágenes** (2 filas × 4), hasta **3 hojas = 24 imágenes**, con **fondo transparente real**. Tú las recortas y las acomodas en la prueba.

**A. Describe cada ilustración.** En cada pregunta (o texto compartido) que necesite una, agrega `"imagenPrompt"`: una descripción **muy detallada** para generarla: qué objetos o personas, cuántos, colores, postura, vista (frontal, lateral, tres cuartos), estilo, y qué NO debe aparecer. Reglas: sin texto dentro de la imagen; **no debe revelar la respuesta**; una sola escena clara; termina con «Fondo transparente». No uses ChatGPT para datos, cantidades exactas ni medidas: eso va en las plantillas del script.

**B. Genera el prompt para el docente:**
```bash
node build_exam.js exam.json --out /mnt/user-data/outputs --prompt-imagenes
```
Crea `prompt_imagenes_chatgpt.md` (y guarda el número de cada imagen en `exam.json`). **Muestra al docente en bloques de código separados** (para copiar con un clic): el **PROMPT GENERAL** y luego **HOJA 1**, **HOJA 2** y **HOJA 3**, y explícale los pasos: 1) pegar el general en un chat nuevo de ChatGPT con imágenes, 2) pegar cada hoja y descargar el PNG, 3) si sale fondo blanco o cuadrícula gris, pedirle «quita el fondo y dame el PNG con transparencia real (canal alfa)», 4) subir las hojas aquí en orden. Adjunta el archivo `.md` con `present_files`.

**C. Cuando el docente suba las hojas** (están en `/mnt/user-data/uploads`, en orden Hoja 1, 2, 3):
```bash
node build_exam.js exam.json --out /mnt/user-data/outputs --sprites hoja1.png hoja2.png hoja3.png
```
El script detecta cada ilustración (aunque no estén perfectamente alineadas), la **recorta al borde exacto sin márgenes**, conserva la transparencia (si la hoja trae fondo blanco o cuadrícula gris, lo quita y avisa), la numera y genera `sprites_previsualizacion.png`. **Mira esa previsualización** y confirma que las 8 imágenes de cada hoja se recortaron completas y en el orden correcto; si alguna falla, dile al docente cuál hoja repetir. Cada pregunta con ilustración queda como un **bloque indivisible**: texto y opciones a la izquierda, ilustración a la derecha (nunca se separan entre columnas ni se desconfigura el texto). Luego haz el PASO 5 y entrega la prueba **final con las imágenes**.

## PASO 5: REVISIÓN ANTES DE ENTREGAR
1. Convierte el docx del estudiante a PDF y mira las páginas como imagen: escudo alineado al margen, sin huecos grandes, hoja de respuestas completa, imágenes legibles.
2. Ejecuta el comando de la GARANTÍA DEL ESCUDO.
3. Revisa `resumen.json` (claves, distribución, hueco máximo, escudo usado) y borra de `/mnt/user-data/outputs` la carpeta `img`, `resumen.json` y archivos temporales.
4. Si el docente quiere **guardar las preguntas en el banco** para reutilizarlas, entrégale también el `exam.json` (el docente lo guarda con `herramientas/guardar-en-banco.js`). Entrega con `present_files`: primero el del estudiante, luego el del docente (con versiones: A estudiante, A docente, B estudiante, B docente), con un mensaje de 3 a 5 líneas. Con dos versiones, recuerda al docente que las **claves son distintas** en cada una y que no deben mezclarse las hojas de respuestas.

## REGLAS FIJAS
- Los contextos pueden ser ficticios, pero coherentes; no presentes datos inventados como reales.
- Nunca reveles la clave en el documento del estudiante, en el nombre del archivo ni en las imágenes.
- Sin emojis dentro de los documentos Word.
- No modifiques el script; si hay un error técnico, corrige primero el JSON.
- **Integridad:** el script compara el texto del Word con el PDF de medición y descarta toda maquetación a la que le falte texto. Si al final imprime ⚠ «faltan fragmentos», abre el PDF y revisa esas preguntas antes de entregar.
- Las instrucciones de los textos compartidos nunca llevan números escritos a mano (`{preguntas}`).
- Las ilustraciones de ChatGPT son solo de contexto; los datos, medidas y cantidades salen siempre de las plantillas del script.
