# Referencia — Apoyos visuales y textos compartidos

Se lee **solo cuando la prueba lleva gráficos, tablas, esquemas, avisos, diálogos o textos compartidos**.

## Apoyos visuales (campo `visual`)
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

## Textos compartidos (`estimulos`) — Lectura, Inglés, Sociales, Ciencias…
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

