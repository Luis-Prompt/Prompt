# Historial de cambios — Creación de pruebas tipo SABER

## 2.6.0 — ilustraciones con ChatGPT
- Nuevo flujo de **hojas de ilustraciones**: cada pregunta (o texto compartido) puede llevar `imagenPrompt`; `--prompt-imagenes` genera el **prompt general + prompts de hasta 3 hojas de 8 imágenes** (2 filas × 4, fondo transparente real) para pegar en ChatGPT.
- `--sprites hoja1.png hoja2.png hoja3.png` **detecta, recorta (sin márgenes) y numera** cada ilustración aunque la hoja no esté perfectamente alineada; conserva la transparencia; si la hoja trae fondo blanco o la falsa cuadrícula gris, lo quita y avisa; genera `sprites_previsualizacion.png`.
- Cada pregunta ilustrada queda como **un bloque indivisible** (texto y opciones a la izquierda, ilustración a la derecha); los textos compartidos llevan su ilustración dentro de la caja.
- Los números de imagen se guardan en el `exam.json`, así no cambian al reordenar.

## 2.5.0 — todas las áreas
- **Catálogo de 20 áreas** (`areas/`): Matemáticas, Geometría, Estadística, Lenguaje, Lectura crítica, Inglés, Ciencias naturales, Biología, Física, Química, Ciencias sociales, Constitución y democracia, Competencias ciudadanas, Filosofía, Ética y valores, Educación religiosa, Educación artística, Educación física, Tecnología e informática, Ciencias económicas y políticas. Cada una con marco oficial por grado, competencias, componentes, reglas de calidad, errores típicos y preguntas iniciales del tutorial. Un ejemplo funcional por área en `examples/areas/`.
- **Textos compartidos** (`estimulos`): un texto, tabla, gráfico o aviso para varias preguntas; siempre juntos y en orden (también en la versión B). Instrucción con numeración automática (`{preguntas}`).
- **11 plantillas gráficas nuevas:** circular, líneas, histograma, plano cartesiano, figuras geométricas, ángulos, línea de tiempo, flujo/ciclo, aviso, diálogo y círculo cromático.
- **Verificación de integridad:** el script compara el texto del Word con el PDF y descarta toda maquetación a la que le falte texto (se detectó un defecto de LibreOffice al partir cajas entre columnas).
- Nuevo `--rapido` (comprobación de contenido sin medir maquetación) y `scripts/probar_areas.sh` (prueba los 20 ejemplos).
- Corrección: en inglés los niveles son Pre A1, A1, A2 y B1 (vigentes).

## 2.4.0 — escudo 2025 y configuración institucional
- Nuevo escudo del colegio (2025) en el encabezado; el respaldo incrustado en el script también se actualizó.
- Nuevo `config/institucion.json` (nombre, escudo y membrete) compartido por todos los prompts; el script lo lee solo.
- **Corrección:** el escudo ya no se rechaza si es un JPG o WebP con extensión `.png` (caso real de la imagen subida por el docente). Además se recortan automáticamente los márgenes blancos.
- Los escudos de los documentos se guardan con paleta de colores: el mismo peso aproximado con mucha mejor calidad.
- Nuevas herramientas: `herramientas/preparar-escudo.py` y `herramientas/guardar-en-banco.js` (banco de preguntas sin duplicados).

## 2.3.0 — versiones A y B
- Nueva opción `"versiones": 2` en `config` (o `--versiones 2`): genera la prueba en **versión A y B**, con otro orden de preguntas y otras letras de clave. Cuatro Word: estudiante y docente de cada versión.
- Cada versión pasa por su propia verificación de maquetación (huecos al pie de columna).
- El título del documento indica "VERSIÓN A" o "VERSIÓN B"; el resumen (`resumen.json`) incluye las claves de la versión B.
- Se elimina el `|| true` que armaba el documento final dos veces sin necesidad.
- Nota: las preguntas con opciones numéricas mantienen las opciones ordenadas de menor a mayor, por lo que su letra puede coincidir entre versiones.

## 2.2.0 — estructura pedagógica
- Cada pregunta admite `competencia`, `componente`, `nivel` (1–4) y `dba`.
- El Word del docente incluye la **matriz de especificaciones** (competencias × niveles de desempeño) y la clasificación bajo cada clave.
- El script valida `nivel` y avisa si faltan competencias.
- Entrevista inicial **por etapas** (tutorial) con resumen de confirmación.
- Autoverificación ciega con rúbrica de 5 puntos, obligatoria en áreas no numéricas.
- La carpeta pasa a `prompts/evaluacion/creacion-pruebas-saber`.

## 2.1.0 — reorganización
- El repositorio pasa de un `.zip` a carpetas versionables.
- `SKILL.md` baja de 86 KB a ~11 KB: se elimina el Anexo A, que era una copia idéntica de `scripts/build_exam.js`.
- Si el script no está instalado como skill, el prompt lo descarga del repositorio.
- Se agregan `package.json`, `README.md`, `ROADMAP.md` y verificación automática (GitHub Actions) lista en `herramientas/github-actions/`, pendiente de activar.
- Sin cambios en la lógica del script: el resultado es idéntico al de la v2.0.

## 2.0.0
- Versión original: dos Word, escudo garantizado, alineación, validaciones y maquetación sin huecos.
