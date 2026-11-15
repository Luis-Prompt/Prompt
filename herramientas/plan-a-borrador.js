#!/usr/bin/env node
/* ============================================================================
 *  plan-a-borrador.js — PUENTE DE DATOS entre «Planeaciones de clase» y «Pruebas tipo SABER»
 *  Convierte las preguntas tipo Saber de una planeación (plan.json) en un BORRADOR de preguntas
 *  con la estructura del prompt de pruebas. Solo mueve DATOS: no mezcla instrucciones de ningún prompt.
 *
 *  Uso : node herramientas/plan-a-borrador.js plan.json [--out preguntas_planeacion.json]
 *  Salida: preguntas_planeacion.json  → el docente lo sube al prompt de pruebas, que completa lo que falta
 *          (justificación, errores de cada distractor, competencia, componente y nivel) y lo verifica.
 * ========================================================================== */
const fs = require('fs'), path = require('path');
const args = process.argv.slice(2), jsonPath = args.find(a => a.endsWith('.json') && args[args.indexOf(a) - 1] !== '--out');
if (!jsonPath) { console.error('Uso: node herramientas/plan-a-borrador.js plan.json [--out archivo.json]'); process.exit(2); }
const out = args.includes('--out') ? args[args.indexOf('--out') + 1] : path.join(path.dirname(path.resolve(jsonPath)), 'preguntas_planeacion.json');
const plan = JSON.parse(fs.readFileSync(jsonPath, 'utf8')), M = plan.meta || {}, F1 = plan.fase1 || {};
const nz = (t) => String(t || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
// asignatura de la planeación → clave de área del prompt de pruebas
const AREAS = { 'matematicas': 'matematicas', 'aritmetica': 'matematicas', 'calculo': 'matematicas', 'estadistica': 'estadistica', 'geometria': 'geometria', 'lenguaje': 'lenguaje', 'lengua castellana': 'lenguaje', 'literatura': 'lenguaje',
  'ciencias naturales': 'ciencias-naturales', 'quimica': 'quimica', 'fisica': 'fisica', 'biologia': 'biologia', 'ciencias sociales': 'ciencias-sociales', 'historia': 'ciencias-sociales', 'geografia': 'ciencias-sociales', 'economia': 'ciencias-economicas-politicas',
  'educacion artistica': 'artistica', 'educacion fisica': 'educacion-fisica', 'educacion etica': 'etica-valores', 'educacion religiosa': 'religion', 'tecnologia e informatica': 'tecnologia-informatica', 'ingles': 'ingles', 'filosofia': 'filosofia' };
const areaClave = AREAS[nz(M.asignatura || M.area)] || 'general';
const grado = String(M.grado || '').replace(/[^\d]/g, '') || String(M.grado || '');
// «¿Cuál es…?» → enunciado (contexto) + pregunta (último interrogante). Si no hay contexto, usa un enunciado neutro con el tema.
function partir(stem, tema) {
  const s = String(stem).trim(), m = s.match(/^([\s\S]*?)([^.!?¿]*¿[^?]*\?|[^.!?]*\?)\s*$/);
  let enun = m ? m[1].trim() : '', preg = m ? m[2].trim() : s;
  if (!m) { enun = s; preg = 'Marca la opción correcta.'; }
  if (!enun) enun = `En la clase de ${M.asignatura || M.area} se trabajó el tema «${tema}».`;
  return { enun, preg };
}
const preguntas = (plan.semanas || []).map((s, i) => {
  const q = s.saber || {}, k = 'abcd'.indexOf(String(q.rta || '').toLowerCase()), ops = q.opciones || [], p = partir(q.pregunta, s.tema);
  return { enunciado: p.enun, pregunta: p.preg, correcta: ops[k], distractores: ops.filter((_, j) => j !== k),
    errores: ['', '', ''], justificacion: '', competencia: '', componente: '', nivel: null, dba: s.dba || '', origen: { semana: s.n || i + 1, tema: s.tema } };
}).filter(p => p.correcta);
const tema = `Periodo ${M.periodo}`, temasDelPeriodo = (F1.ejes || []).map(e => e.tema);
const borrador = { _borrador: true,
  _instrucciones: 'BORRADOR para el prompt de pruebas: completa justificacion, errores (uno por distractor, en el mismo orden), competencia, componente y nivel (1-4) de cada pregunta según el catálogo del área; si las opciones son numéricas agrega el campo calculo; verifica la respuesta de cada una y propón un tema descriptivo para la prueba (temasDelPeriodo) antes de armarla.',
  meta: { area: M.area || M.asignatura, areaClave, asignatura: M.asignatura || M.area, tema, temasDelPeriodo, grado, docente: M.docente || '' }, config: {}, preguntas };
fs.writeFileSync(out, JSON.stringify(borrador, null, 2) + '\n');
console.log(`📝 ${preguntas.length} pregunta(s) de la planeación convertidas en borrador → ${out}`);
console.log(`   Área de la prueba: ${areaClave} · grado ${grado} · pendiente por completar: justificación, errores, competencia, componente y nivel.`);
