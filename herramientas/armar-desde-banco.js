#!/usr/bin/env node
/* ============================================================================
 *  armar-desde-banco.js — Arma un exam.json con preguntas YA GUARDADAS en el banco.
 *
 *  Uso:
 *    node herramientas/armar-desde-banco.js --listar
 *    node herramientas/armar-desde-banco.js --area matematicas --grado 5 [--tema fracciones] [--n 10] [--abiertas 1]
 *         [--estado aprobada|revisada|borrador] [--semilla 7] [--excluir ids.json] [--docente "Nombre"]
 *         [--longitud corto|medio|largo] [--banco carpeta] [--out exam.json]
 *    luego:  node prompts/evaluacion/creacion-pruebas-saber/scripts/build_exam.js exam.json --out salida
 *
 *  - Por defecto solo usa preguntas en estado «aprobada» (--estado revisada incluye revisadas y aprobadas; borrador, todas).
 *  - Reparte las preguntas entre las competencias y los niveles de dificultad disponibles (selección equilibrada).
 *  - Con --semilla el resultado es repetible; sin ella cambia cada vez (útil para versiones distintas de una prueba).
 *  - Las preguntas que comparten un texto o gráfico quedan juntas y el texto viaja con ellas.
 *  - --excluir recibe un JSON con una lista de ids (los de una prueba anterior) para no repetir preguntas.
 *  - Deja en el exam.json el campo «origenBanco» con los ids usados (para trazabilidad y para --excluir).
 * ========================================================================== */
const fs = require('fs'), path = require('path');
const args = process.argv.slice(2);
const val = (k, d) => args.includes(k) ? args[args.indexOf(k) + 1] : d;
const banco = val('--banco', path.join(__dirname, '..', 'recursos', 'banco-preguntas'));
const slug = (t) => String(t || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '');
const RANGO = { borrador: 0, revisada: 1, aprobada: 2 };
const fail = (m) => { console.error('❌ ' + m); process.exit(1); };

// ---- leer el banco completo ----
const archivos = [];
(function rec(d) { if (!fs.existsSync(d)) return; for (const n of fs.readdirSync(d, { withFileTypes: true })) { const p = path.join(d, n.name); n.isDirectory() ? rec(p) : (n.name.endsWith('.json') && archivos.push(p)); } })(banco);
const temas = archivos.map(f => { try { return { f, ...JSON.parse(fs.readFileSync(f, 'utf8')), _area: path.basename(path.dirname(path.dirname(f))), _grado: path.basename(path.dirname(f)) }; } catch (e) { fail(`No se pudo leer ${f}: ${e.message}`); } }).filter(t => Array.isArray(t.preguntas));

if (args.includes('--listar')) {
  if (!temas.length) { console.log('📭 El banco está vacío. Guarda preguntas con: node herramientas/guardar-en-banco.js exam.json'); process.exit(0); }
  console.log('📚 Banco de preguntas (borrador / revisada / aprobada):');
  temas.forEach(t => { const c = { borrador: 0, revisada: 0, aprobada: 0 }; t.preguntas.forEach(p => c[p.estado || 'borrador']++); console.log(`   ${t._area}/${t._grado}/${path.basename(t.f, '.json')}: ${c.borrador} / ${c.revisada} / ${c.aprobada}`); });
  process.exit(0);
}

// ---- filtrar ----
const area = slug(val('--area', '')), grado = slug(val('--grado', '')), tema = slug(val('--tema', ''));
if (!area || !grado) fail('Indica --area y --grado (por ejemplo: --area matematicas --grado 5). Usa --listar para ver qué hay en el banco.');
const estadoMin = val('--estado', 'aprobada'); if (!(estadoMin in RANGO)) fail('--estado debe ser borrador, revisada o aprobada');
const n = Number(val('--n', 10)), nAbiertas = Number(val('--abiertas', 0));
if (!Number.isInteger(n) || n < 1 || !Number.isInteger(nAbiertas) || nAbiertas < 0) fail('--n debe ser un entero ≥ 1 y --abiertas un entero ≥ 0');
let excluidos = new Set(); if (val('--excluir', null)) { try { const x = JSON.parse(fs.readFileSync(val('--excluir'), 'utf8')); excluidos = new Set(Array.isArray(x) ? x : (x.origenBanco && x.origenBanco.ids) || []); } catch (e) { fail('No se pudo leer --excluir: ' + e.message); } }
const fuentes = temas.filter(t => slug(t._area) === area && slug(t._grado) === grado && (!tema || slug(path.basename(t.f, '.json')) === tema));
if (!fuentes.length) fail(`No hay preguntas en el banco para ${area}/${grado}${tema ? '/' + tema : ''}. Usa --listar para ver qué hay.`);
const cand = fuentes.flatMap(t => t.preguntas.map(p => ({ p, t }))).filter(({ p }) => RANGO[p.estado || 'borrador'] >= RANGO[estadoMin] && !excluidos.has(p.id));
const mc = cand.filter(({ p }) => p.tipo !== 'abierta'), ab = cand.filter(({ p }) => p.tipo === 'abierta');
if (mc.length < n) fail(`Solo hay ${mc.length} preguntas de selección múltiple en estado «${estadoMin}» o superior y pediste ${n}. Revisa y aprueba más preguntas (campo "estado" del banco) o baja --n / usa --estado revisada.`);
if (ab.length < nAbiertas) fail(`Solo hay ${ab.length} preguntas abiertas disponibles y pediste ${nAbiertas}.`);

// ---- selección equilibrada (competencia y nivel), con semilla opcional ----
let s = val('--semilla', null) !== null ? Number(val('--semilla')) >>> 0 : (Date.now() ^ (process.pid << 8)) >>> 0;
const azar = () => { s = (s + 0x6D2B79F5) >>> 0; let t = s; t = Math.imul(t ^ (t >>> 15), t | 1); t ^= t + Math.imul(t ^ (t >>> 7), t | 61); return ((t ^ (t >>> 14)) >>> 0) / 4294967296; };
function elegir(lista, k) {
  const pool = lista.slice(), sel = [], porComp = {}, porNivel = {};
  while (sel.length < k) {
    let mejor = null, puntaje = Infinity;
    for (const c of pool) { const sc = (porComp[c.p.competencia] || 0) * 10 + (porNivel[c.p.nivel] || 0) + azar() * 0.9; if (sc < puntaje) { puntaje = sc; mejor = c; } }
    sel.push(mejor); pool.splice(pool.indexOf(mejor), 1);
    porComp[mejor.p.competencia] = (porComp[mejor.p.competencia] || 0) + 1; porNivel[mejor.p.nivel] = (porNivel[mejor.p.nivel] || 0) + 1;
  }
  return sel;
}
const elegidas = [...elegir(mc, n), ...elegir(ab, nAbiertas)];

// ---- armar el exam.json (las preguntas con el mismo texto compartido van juntas) ----
const orden = [], vistos = new Set();
for (const c of elegidas) { if (c.p.estimulo) { if (vistos.has(c.p.estimulo)) continue; vistos.add(c.p.estimulo); orden.push(...elegidas.filter(x => x.p.estimulo === c.p.estimulo)); } else orden.push(c); }
const finales = [...orden.filter(c => c.p.tipo !== 'abierta'), ...orden.filter(c => c.p.tipo === 'abierta')];
const QUITAR = new Set(['id', 'estado', 'fechaRegistro', 'idEnLaPrueba', 'estimuloData']);
const estimulos = []; finales.forEach(({ p }) => { if (p.estimuloData && !estimulos.some(e => e.id === p.estimuloData.id)) estimulos.push(p.estimuloData); });
const t0 = fuentes[0], temasUsados = [...new Set(fuentes.map(t => t.tema))];
const ex = {
  meta: { area: t0.area, areaClave: area, asignatura: t0.area, tema: temasUsados.length === 1 ? temasUsados[0] : temasUsados.join(' · '), grado: String(t0.grado || grado), docente: val('--docente', '') },
  config: { longitud: val('--longitud', 'medio') },
  ...(estimulos.length ? { estimulos } : {}),
  preguntas: finales.map(({ p }) => Object.fromEntries(Object.entries(p).filter(([k]) => !QUITAR.has(k)))),
  origenBanco: { generado: new Date().toISOString().slice(0, 10), estadoMinimo: estadoMin, ids: finales.map(({ p }) => p.id) }
};
const out = val('--out', 'exam.json'); fs.writeFileSync(out, JSON.stringify(ex, null, 2) + '\n');

const cuenta = (k) => finales.filter(({ p }) => p.tipo !== 'abierta').reduce((a, { p }) => (a[p[k]] = (a[p[k]] || 0) + 1, a), {});
console.log(`✅ ${out}: ${finales.length - nAbiertas} de selección múltiple + ${nAbiertas} abiertas · disponibles: ${mc.length} / ${ab.length} · estado mínimo: ${estadoMin}`);
console.log('   Competencias:', Object.entries(cuenta('competencia')).map(([k, v]) => `${k} (${v})`).join(', '));
console.log('   Niveles:', Object.entries(cuenta('nivel')).sort().map(([k, v]) => `${k} (${v})`).join(', '));
console.log('   Siguiente paso: node prompts/evaluacion/creacion-pruebas-saber/scripts/build_exam.js ' + out + ' --out salida');
