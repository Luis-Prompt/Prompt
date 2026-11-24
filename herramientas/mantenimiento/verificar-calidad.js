#!/usr/bin/env node
/* verificar-calidad.js — Revisa la salud general del repositorio.
 * Uso: node herramientas/mantenimiento/verificar-calidad.js
 * Comprueba: (1) que todos los .json sean válidos, (2) que las rutas del escudo en config/institucion.json existan,
 * (3) que los enlaces relativos de todos los .md apunten a archivos que existen, (4) que no se haya subido
 * ningún secreto (tokens de GitHub, claves de API, llaves privadas) y (5) que ningún archivo pese más de 5 MB. */
const fs = require('fs'), path = require('path');
const raiz = path.join(__dirname, '..', '..'), E = [];
const IGNORAR = new Set(['.git', 'node_modules', 'salida']);
const todos = [];
(function recorrer(d) {
  for (const n of fs.readdirSync(d, { withFileTypes: true })) {
    if (IGNORAR.has(n.name)) continue;
    const p = path.join(d, n.name);
    n.isDirectory() ? recorrer(p) : todos.push(p);
  }
})(raiz);
const rel = (p) => path.relative(raiz, p).split(path.sep).join('/');

// 1. JSON válidos
const jsons = todos.filter(f => f.endsWith('.json'));
jsons.forEach(f => { try { JSON.parse(fs.readFileSync(f, 'utf8')); } catch (e) { E.push(`JSON inválido en ${rel(f)}: ${e.message}`); } });

// 2. Rutas del escudo
try {
  const cfg = JSON.parse(fs.readFileSync(path.join(raiz, 'recursos/colegio/institucion.json'), 'utf8'));
  Object.entries(cfg.escudo || {}).forEach(([k, v]) => { if (!fs.existsSync(path.resolve(raiz, 'recursos/colegio', v))) E.push(`recursos/colegio/institucion.json: el escudo «${k}» apunta a ${v} y no existe`); });
} catch (e) { E.push('No se pudo leer recursos/colegio/institucion.json: ' + e.message); }

// 3. Enlaces relativos en Markdown (se ignoran los bloques de código)
todos.filter(f => f.endsWith('.md')).forEach(f => {
  const txt = fs.readFileSync(f, 'utf8').replace(/```[\s\S]*?```/g, '').replace(/`[^`\n]*`/g, '');
  (txt.match(/\]\(([^)\s]+)\)/g) || []).map(x => x.slice(2, -1).split('#')[0]).filter(x => x && !/^(https?:|mailto:)/.test(x)).forEach(l => {
    if (!fs.existsSync(path.resolve(path.dirname(f), decodeURIComponent(l)))) E.push(`${rel(f)} enlaza a «${l}» y no existe`);
  });
});

// 4. Secretos
const PATRONES = [/github_pat_[A-Za-z0-9_]{20,}/, /gh[pousr]_[A-Za-z0-9]{30,}/, /sk-ant-[A-Za-z0-9\-_]{20,}/, /-----BEGIN [A-Z ]*PRIVATE KEY-----/, /AKIA[0-9A-Z]{16}/];
todos.filter(f => !/\.(png|jpe?g|gif|docx|xlsx|pdf|zip)$/i.test(f) && !f.endsWith('package-lock.json')).forEach(f => {
  const t = fs.readFileSync(f, 'utf8');
  PATRONES.forEach(r => { if (r.test(t)) E.push(`Posible secreto en ${rel(f)} (patrón ${r})`); });
});

// 5. Tamaño
todos.forEach(f => { const mb = fs.statSync(f).size / 1048576; if (mb > 5) E.push(`${rel(f)} pesa ${mb.toFixed(1)} MB (máximo 5 MB)`); });

if (E.length) { console.error(E.map(e => '❌ ' + e).join('\n')); process.exit(1); }
console.log(`✅ Calidad del repositorio: ${jsons.length} JSON válidos, enlaces y rutas correctos, sin secretos, sin archivos pesados.`);
