#!/usr/bin/env node
/* ============================================================================
 *  guardar-en-banco.js — Guarda las preguntas de un exam.json en el banco de preguntas.
 *  Uso:  node herramientas/guardar-en-banco.js exam.json [--banco banco-preguntas] [--estado aprobada]
 *  Efecto: banco-preguntas/<area>/<grado>/<tema>.json  (se crea o se completa)
 *  - Cada pregunta recibe un id estable (huella del enunciado+pregunta+correcta): no se duplica.
 *  - Solo guarda preguntas de selección múltiple y abiertas con su clasificación y justificación.
 *  - Estados: borrador | revisada | aprobada (por defecto: borrador, para que el docente las revise).
 * ========================================================================== */
const fs = require('fs'), path = require('path'), crypto = require('crypto');
const args = process.argv.slice(2), jsonPath = args.find(a => a.endsWith('.json'));
if (!jsonPath) { console.error('Uso: node herramientas/guardar-en-banco.js exam.json [--banco carpeta] [--estado borrador|revisada|aprobada]'); process.exit(2); }
const val = (k, d) => args.includes(k) ? args[args.indexOf(k) + 1] : d;
const banco = val('--banco', path.join(__dirname, '..', 'banco-preguntas')), estado = val('--estado', 'borrador');
if (!['borrador', 'revisada', 'aprobada'].includes(estado)) { console.error('--estado debe ser borrador, revisada o aprobada'); process.exit(2); }

const slug = (t) => String(t || 'general').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '') || 'general';
const norm = (t) => String(t || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/\s+/g, ' ').trim();
const huella = (p) => crypto.createHash('sha1').update([p.tipo === 'abierta' ? 'a' : 'm', norm(p.enunciado), norm(p.pregunta), norm(p.correcta || p.respuestaModelo)].join('|')).digest('hex').slice(0, 10);
const CAMPOS_EFIMEROS = new Set(['num', '_ops', '_clave', '_imgs']);

const ex = JSON.parse(fs.readFileSync(jsonPath, 'utf8')), m = ex.meta || {};
const area = slug(m.areaClave || m.area), grado = slug(m.grado), tema = slug(m.tema);
const dir = path.join(banco, area, grado), file = path.join(dir, tema + '.json');
fs.mkdirSync(dir, { recursive: true });
const base = fs.existsSync(file) ? JSON.parse(fs.readFileSync(file, 'utf8')) : { area: m.area || area, grado: String(m.grado || ''), tema: m.tema || tema, preguntas: [] };
const ids = new Set(base.preguntas.map(p => p.id));
let nuevas = 0, repetidas = 0;
for (const p of ex.preguntas || []) {
  const id = huella(p); if (ids.has(id)) { repetidas++; continue; }
  const limpia = Object.fromEntries(Object.entries(p).filter(([k]) => !CAMPOS_EFIMEROS.has(k)));
  base.preguntas.push({ id, estado, fechaRegistro: new Date().toISOString().slice(0, 10), ...limpia }); ids.add(id); nuevas++;
}
fs.writeFileSync(file, JSON.stringify(base, null, 2) + '\n');
console.log(`📚 ${path.relative(process.cwd(), file) || file}: ${nuevas} nuevas, ${repetidas} repetidas (ya estaban). Total en el archivo: ${base.preguntas.length}.`);
