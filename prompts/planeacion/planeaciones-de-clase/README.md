# Planeaciones de clase (v1.0)

Tutor interactivo que guía al docente, **con botones y de forma gráfica**, para crear la planeación completa de un periodo y entregarla en **Word horizontal y PDF**, con el membrete y el escudo del colegio en todas las páginas.

## Qué entrega
- `Plan_<Area>_Grado<G>_Periodo<N>_<Año>.docx` y `.pdf` (7 a 9 páginas).
- Contenido: identificación · **Fase 1** estructura general (meta, pregunta problematizadora, ejes, DBA, estándares, marco legal) · **Fase 2** ruta del periodo y tabla maestra de 9 semanas (secuencia didáctica, **sustento teórico distinto cada semana**, **pregunta tipo Saber con respuesta**, compromiso y recursos) · **Fase 3** competencias, logro integral, desempeños, sistema de evaluación, matriz de seguimiento, apoyos, bibliografía APA 7 y firmas.

## Cómo se usa (docente)
1. Abre Claude y escribe: *«Usa este prompt: https://github.com/Luis-Prompt/Prompt/blob/main/prompts/planeacion/planeaciones-de-clase/SKILL.md»*.
2. Responde con clics las 5 rondas (datos base · asignatura y docente · temas y enfoque · proyectos y contexto · vista previa).
3. Recibe el Word y el PDF.

## Contenido de la carpeta
| Archivo | Función |
|---|---|
| `SKILL.md` | El prompt: entrevista por rondas, reglas de calidad y pasos de entrega. |
| `scripts/build_plan.js` | Valida el plan y construye el Word y el PDF (membrete, tablas, firmas, integridad). |
| `catalogo/` | Conocimiento editable: `asignaturas`, `temas` (16 familias × grado), `teorias` (20 autores), `enfoques`, `competencias`, `normativa`. |
| `assets/escudo_header.png` | Escudo 2025 (copia del banco de imágenes). |
| `examples/plan_ejemplo.json` | Planeación completa de Matemáticas 4.°, periodo 1. |
| `CHANGELOG.md` · `ROADMAP.md` | Historial y mejoras pendientes. |

## Probar en el computador
Requisitos: Node 18+, LibreOffice y `poppler-utils`.
```bash
cd prompts/planeacion/planeaciones-de-clase
npm install
npm run ejemplo        # genera el Word y el PDF en ./salida
```

## Control de calidad que hace el script
- **Valida** el plan: semanas completas, sustentos que no se repiten, preguntas Saber con 4 opciones distintas y `rta` válida, mínimo 2 DBA y 2 referencias, porcentajes de evaluación que suman 100.
- **Integridad:** compara el texto del plan con el PDF y avisa si falta algo.
- **DBA honestos:** los DBA no verificados se imprimen con † para que el docente los confirme.
- El membrete sale de `config/institucion.json` (compartido con los demás prompts).
