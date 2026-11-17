# Referencia — Ilustraciones con ChatGPT (hojas de 8 imágenes)

Se lee **solo cuando el docente quiere ilustraciones creadas con ChatGPT**.

## Flujo completo
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

