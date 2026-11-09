# Mejoras pendientes — Creación de pruebas tipo SABER

Ordenadas por prioridad. Marca con `x` lo terminado.

## Alta
- [x] **Estructura pedagógica por pregunta (v2.2):** campos `competencia`, `componente`, `nivel` (1–4) y `dba/evidencia`; matriz de especificaciones y resumen de cobertura en el Word del docente.
- [x] **Versiones A/B (v2.3):** misma prueba con preguntas y opciones en otro orden, con su propia clave.
- [x] **Revisión ciega por rúbrica (v2.2)** para áreas no numéricas (Lectura crítica, Ciencias, Sociales): Claude resuelve cada pregunta sin ver la clave y reporta dudas.
- [x] **Preguntas iniciales por etapas (v2.2)** (tutorial), con opciones numeradas y resumen de confirmación antes de generar. Agregar: tipo de prueba (3.°, 5.°, 9.°, 11.°, Pro), competencia, tiempo, adaptaciones (NEE).

## Media
- [ ] **Más plantillas gráficas:** geometría (ángulos, perímetro, área), plano cartesiano, diagrama circular, histograma, pictograma, línea de tiempo, ciclos (Ciencias), textos discontinuos (Lectura).
- [x] **`config/institucion.json` (v2.4)** con membrete, escudo y nombre del colegio, para no depender del código.
- [x] Exportar cada pregunta a `banco-preguntas/` (v2.4, `herramientas/guardar-en-banco.js`).
- [ ] Armar una prueba nueva **tomando preguntas aprobadas del banco** (filtrar por área, grado, tema, competencia y nivel).
- [x] Quitar el `|| true` del armado final (v2.3).

## Baja
- [ ] Asistente gráfico en GitHub Pages (formulario que arma el `exam.json`).
- [ ] Exportación a PDF y a formato de lectura óptica.
