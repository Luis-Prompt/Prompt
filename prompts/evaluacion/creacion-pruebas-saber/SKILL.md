---
name: generador-examen-tipo-saber
description: Genera DOS documentos Word (prueba del estudiante + prueba del docente con claves) tipo SABER/ICFES para el Colegio Integrado Simón Bolívar de Cúcuta, con escudo en el encabezado, dos columnas sin espacios en blanco, hoja de respuestas compacta, preguntas de selección múltiple verificadas y preguntas abiertas opcionales. Úsalo cuando el docente pida crear, generar o hacer un examen, prueba, evaluación o simulacro tipo SABER.
---

# GENERADOR DE EXAMEN TIPO SABER — v2.2
**Colegio Integrado Simón Bolívar · Cúcuta · Mg. Luis Alberto Avellaneda C.**

## QUÉ ENTREGA (SIEMPRE DOS ARCHIVOS .docx)
1. `…_ESTUDIANTE.docx`: encabezado con escudo, datos del estudiante, preguntas a 2 columnas, hoja de respuestas con óvalos grandes.
2. `…_DOCENTE_CLAVES.docx`: la misma prueba con la respuesta correcta resaltada en verde, plantilla de calificación (óvalos llenos), justificación de cada clave y el error que representa cada distractor, y respuesta modelo + criterios de las preguntas abiertas.

Entrega ambos con `present_files` y un mensaje corto (claves, nº de preguntas, imágenes). No pegues el código en la respuesta.

---

## PASO 0: ENTREVISTA INICIAL POR ETAPAS (tutorial)
Habla como un guía amable. Haz **una etapa por mensaje**, con preguntas numeradas y las opciones ya escritas (el docente responde con números o letras). Si el docente ya dio algún dato, no lo repitas. Al terminar, muestra un **resumen de confirmación** y espera el "sí" antes de generar.

**Etapa 1 — ¿Qué vamos a evaluar?**
1. **Área o prueba:** (a) Matemáticas · (b) Lectura crítica · (c) Ciencias naturales · (d) Sociales y ciudadanas · (e) Inglés · (f) Otra.
2. **Grado** (3.°, 5.°, 7.°, 9.°, 11.° u otro) y **tema(s)** concretos.
3. **Competencia(s) a evaluar** (si no sabe, propón las de la prueba SABER del área y deja que elija).

**Etapa 2 — ¿Cómo quieres la prueba?**
4. **¿Cuántas preguntas de selección múltiple?** (recomendado 10, 15 o 20).
5. **¿Preguntas abiertas?** Si **sí**, ¿cuántas? (1 a 4); si **no**, no se incluye la PARTE II.
6. **Dificultad:** básica / media / alta / mixta (si es mixta, reparte entre los niveles de desempeño 1 a 4).
7. **Longitud de enunciados:** corto (80–180 caracteres) / medio (250–420) / largo (550–800).
8. **Apoyos visuales:** ¿cuántas preguntas llevan imagen, tabla, gráfico o cuadro?

**Etapa 3 — Detalles finales (opcionales)**
9. Nombre del docente y período (si no los da, quedan en blanco para llenar a mano).
10. Tiempo disponible y si hay estudiantes que necesiten adaptaciones (letra más grande, menos texto).

**Resumen de confirmación:** área, grado, tema, competencias, nº de preguntas, dificultad, apoyos visuales. Pregunta: "¿Lo genero así o cambias algo?".

---

## PASO 1: PREPARAR EL ENTORNO (sin pedir permiso)
1. Busca `/mnt/skills/user/generador-examen-tipo-saber/scripts/build_exam.js`. Si existe, úsalo.
2. Si **no** existe, descarga el script oficial del repositorio (no lo reescribas ni lo modifiques):
   ```bash
   mkdir -p /home/claude/examen/scripts /home/claude/examen/assets && cd /home/claude/examen
   B=https://raw.githubusercontent.com/Luis-Prompt/Prompt/main/prompts/evaluacion/creacion-pruebas-saber
   curl -sL -o scripts/build_exam.js $B/scripts/build_exam.js
   curl -sL -o assets/escudo_header.png $B/assets/escudo_header.png
   ```
   Si no hay red, pide al docente que suba `build_exam.js` y el escudo.
3. `node` y `sharp` ya están disponibles. Si `require('docx')` falla: `npm install -g docx sharp`.

## GARANTÍA DEL ESCUDO (siempre sale en el encabezado)
El script busca el escudo en este orden y **aborta con error si no lo encuentra** (nunca entrega un examen sin logo):
1. variable `ESCUDO_PATH`; 2. `assets/escudo_header.png` junto al script; 3. cualquier `escudo*.png` subido a `/mnt/user-data/uploads/`; 4. **copia base64 incrustada en el propio script** (PNG 200×242 px, ≈9 KB, fondo blanco).
Valida que sea un PNG real y lo normaliza (fondo blanco, ≤300 px). Se inserta **solo en el encabezado** de Word (celda izquierda del membrete, 76 px de alto, proporción original): aparece en **todas las páginas** de **los dos documentos**.
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
| `svg` | `svg` completo (último recurso; también se rasteriza a PNG) |
- **tipo 2: tabla** `{titulo, headers[], rows[][]}`.
- **tipo 3: gráfico de barras** (`plantilla:"barras"`).
- **tipo 4: cuadro de texto** `{titulo, texto, fuente?}`.
`leyenda` opcional como pie de figura. No se descargan imágenes de internet (en este entorno el acceso externo suele estar bloqueado); se generan localmente en PNG.

### Preguntas abiertas (solo si el docente dijo que sí)
`{"tipo":"abierta","enunciado":"…","pregunta":"…","lineas":4,"puntos":2,"respuestaModelo":"…","criterios":["… (1 pto)","… (1 pto)"]}`. Salen en la **PARTE II** con líneas para escribir a mano; en el documento docente, con respuesta modelo y criterios.

---

## PASO 3: ARCHIVO `exam.json`
```json
{
  "meta":   {"area":"Aritmética","areaClave":"matematicas","asignatura":"Aritmética","tema":"Fracciones","grado":"5","docente":""},
  "config": {"longitud":"medio"},
  "preguntas": [
    {"enunciado":"…","pregunta":"¿…?","correcta":"5/8","distractores":["3/8","3/5","8/5"],
     "calculo":"(8-3)/8","justificacion":"…","errores":["error 1","error 2","error 3"],
     "competencia":"Planteamiento y resolución de problemas","componente":"Numérico-variacional","nivel":2,"dba":"…",
     "visual":{"tipo":1,"plantilla":"pizza","params":{"partes":8,"marcadas":3,"etiquetaMarcadas":"Porciones comidas"},"leyenda":"…"}},
    {"tipo":"abierta","enunciado":"…","pregunta":"…","lineas":4,"puntos":2,"respuestaModelo":"…","criterios":["…"]}
  ]
}
```
`competencia`, `componente`, `nivel` (1–4) y `dba` son opcionales pero **recomendados**: sin ellos no hay matriz de especificaciones (el script avisa). `areaClave`: matematicas, ciencias, lenguaje, sociales, tecnologia o general (solo cambia la frase final). `ordenar:false` desactiva el orden numérico en una pregunta. Ejemplo completo en `examples/exam_ejemplo.json`. Todo en español con tildes y ñ.

## PASO 4: GENERAR Y CORREGIR
```bash
cd /home/claude/examen && node build_exam.js exam.json --out /mnt/user-data/outputs
```
(o `node /mnt/skills/user/generador-examen-tipo-saber/scripts/build_exam.js exam.json --out /mnt/user-data/outputs`). Tarda ≈ 1 minuto: mide, reordena y verifica; ejecútalo en segundo plano si el límite de la herramienta es corto. **Si imprime ❌ con una lista, corrige exactamente esas preguntas en `exam.json` y repite** hasta ver ✅. Revisa los ⚠ (sobre todo “opciones numéricas sin calculo” y “visuales consecutivos”).

## PASO 5: REVISIÓN ANTES DE ENTREGAR
1. Convierte el docx del estudiante a PDF y mira las páginas como imagen: escudo alineado al margen, sin huecos grandes, hoja de respuestas completa, imágenes legibles.
2. Ejecuta el comando de la GARANTÍA DEL ESCUDO.
3. Revisa `resumen.json` (claves, distribución, hueco máximo, escudo usado) y borra de `/mnt/user-data/outputs` la carpeta `img`, `resumen.json` y archivos temporales.
4. Entrega con `present_files`: primero el del estudiante, luego el del docente, con un mensaje de 3 a 5 líneas.

## REGLAS FIJAS
- Los contextos pueden ser ficticios, pero coherentes; no presentes datos inventados como reales.
- Nunca reveles la clave en el documento del estudiante, en el nombre del archivo ni en las imágenes.
- Sin emojis dentro de los documentos Word.
- No modifiques el script; si hay un error técnico, corrige primero el JSON.
