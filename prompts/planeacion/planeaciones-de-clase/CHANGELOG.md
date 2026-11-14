# Historial de cambios — Planeaciones de clase

## 1.0.1 — aislamiento
- Bloque «ALCANCE Y FRONTERAS» y descripción con lo que NO hace; el cierre ya no mezcla la entrevista de pruebas (protocolo de cambio en `INDICE-PARA-CLAUDE.md`).

## 1.0.0 — primera versión
- Base: el prompt «Generador de Planeaciones Académicas en DOCX» (v3), reorganizado y potenciado.
- **Entrevista por rondas gráfica y realmente usable con botones:** la herramienta de botones admite 2 a 4 opciones, así que el grado y la asignatura se eligen en **dos pasos** o con **paginación** (el prompt original pedía 12 y 17 botones). Barra de progreso, mapa del proceso, tarjeta de resumen y **vista previa del cronograma** de las 9 semanas antes de generar. Modo express.
- **Catálogo editable** (`catalogo/`): temas por grado para 16 familias de áreas (el original cubría 5), 20 autores y teorías para el sustento, 6 enfoques, competencias por área y normativa con su vigencia por verificar.
- **Generador `build_plan.js`:** Word horizontal con **membrete real de Word y escudo en todas las páginas** (con «Página X de Y» dinámico), ruta visual del periodo, tabla maestra con encabezado repetido, matriz de seguimiento, bloque de firmas indivisible y **PDF**.
- **Validación y honestidad:** DBA con † cuando no se pudieron verificar; verificación de integridad texto-PDF; sustentos que no se repiten; preguntas Saber con respuesta única.
- Nuevo: sistema de evaluación con ponderación (Saber / Hacer / Ser), atención a la diversidad (DUA/PIAR, Decreto 1421 de 2017), normas de los proyectos transversales y Educación Religiosa en el menú.
- Corrección normativa: Decreto 1290 de 2009 **compilado en el Decreto 1075 de 2015**.
