# Mejoras pendientes — Generador de examen tipo SABER

Ordenadas por prioridad. Marca con `x` lo terminado.

## Alta
- [ ] **Estructura pedagógica por pregunta:** campos `competencia`, `componente`, `nivel` (1–4) y `dba/evidencia`; matriz de especificaciones y resumen de cobertura en el Word del docente.
- [ ] **Versiones A/B:** misma prueba con preguntas y opciones en otro orden, con su propia clave.
- [ ] **Revisión ciega por rúbrica** para áreas no numéricas (Lectura crítica, Ciencias, Sociales): Claude resuelve cada pregunta sin ver la clave y reporta dudas.
- [ ] **Preguntas iniciales por etapas** (tutorial), con opciones numeradas y resumen de confirmación antes de generar. Agregar: tipo de prueba (3.°, 5.°, 9.°, 11.°, Pro), competencia, tiempo, adaptaciones (NEE).

## Media
- [ ] **Más plantillas gráficas:** geometría (ángulos, perímetro, área), plano cartesiano, diagrama circular, histograma, pictograma, línea de tiempo, ciclos (Ciencias), textos discontinuos (Lectura).
- [ ] **`config/institucion.json`** con membrete, escudo y nombre del colegio, para no depender del código.
- [ ] Exportar cada pregunta aprobada a `banco-preguntas/` para reutilizarla.
- [ ] Quitar el `|| true` del armado final (hoy el documento se arma dos veces).

## Baja
- [ ] Asistente gráfico en GitHub Pages (formulario que arma el `exam.json`).
- [ ] Exportación a PDF y a formato de lectura óptica.
