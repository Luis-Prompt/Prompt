#!/usr/bin/env node
/* verificar-indice.js — Comprueba que los prompts del repositorio estén aislados y bien enrutados.
 * Uso: node herramientas/mantenimiento/verificar-indice.js           (desde cualquier carpeta)
 * Verifica: cada prompt del INDICE.json existe, tiene nombre y descripción con su límite («NO sirve para»), un bloque
 * «ALCANCE Y FRONTERAS», descarga archivos SOLO de su propia carpeta y las URL del índice apuntan a su SKILL.md. */
const fs = require('fs'), path = require('path');
const raiz = path.join(__dirname, '..', '..'), E = [], ok = [];
const falla = (m) => E.push(m), bien = (m) => ok.push(m);
const leer = (p) => fs.readFileSync(path.join(raiz, p), 'utf8');
let idx; try { idx = JSON.parse(leer('prompts/INDICE.json')); } catch (e) { console.error('❌ prompts/INDICE.json no es válido: ' + e.message); process.exit(1); }
const ids = new Set();
for (const p of idx.prompts) {
  const c = `[${p.id}]`;
  if (ids.has(p.id)) falla(`${c} id repetido`); ids.add(p.id);
  if (!fs.existsSync(path.join(raiz, p.carpeta))) { falla(`${c} no existe la carpeta ${p.carpeta}`); continue; }
  if (!fs.existsSync(path.join(raiz, p.skill))) { falla(`${c} no existe ${p.skill}`); continue; }
  if (!p.skill.startsWith(p.carpeta + '/')) falla(`${c} el SKILL.md no está dentro de su carpeta`);
  for (const k of ['skill_url_raw', 'skill_url_github']) if (!p[k].endsWith(p.skill)) falla(`${c} ${k} no apunta a ${p.skill}`);
  const s = leer(p.skill), fm = (s.match(/^---\n([\s\S]*?)\n---/) || [])[1] || '';
  const nombre = (fm.match(/^name:\s*(.+)$/m) || [])[1], desc = (fm.match(/^description:\s*(.+)$/m) || [])[1] || '';
  if (!nombre) falla(`${c} el SKILL.md no tiene «name» en el encabezado`);
  if (!/NO sirve para/.test(desc)) falla(`${c} la descripción debe decir para qué NO sirve («NO sirve para…»)`);
  if (!/## ALCANCE Y FRONTERAS/.test(s)) falla(`${c} falta el bloque «ALCANCE Y FRONTERAS»`);
  // aislamiento: el SKILL solo descarga archivos de su carpeta (y datos compartidos)
  const B = (s.match(/B=\$R\/(prompts\/[^\s]+)/) || [])[1]; if (B && B !== p.carpeta) falla(`${c} descarga archivos de otra carpeta: ${B}`);
  const otras = idx.prompts.filter(q => q.id !== p.id).map(q => q.carpeta);
  otras.forEach(o => { if (s.includes(o + '/scripts') || s.includes(o + '/catalogo') || s.includes(o + '/areas')) falla(`${c} usa archivos internos de otro prompt (${o})`); });
  (s.match(/`referencia\/[\w\-.]+\.md`/g) || []).map(x => x.replace(/`/g, '')).forEach(rel => { if (!fs.existsSync(path.join(raiz, p.carpeta, rel))) falla(`${c} cita ${rel} y no existe`); });
  (s.match(/for f in ([^;]+); do curl[^\n]*\$B\/(\S+)\/\$f\.json/g) || []).forEach(m => { const nombres = m.match(/for f in ([^;]+);/)[1].split(/\s+/), dir = m.match(/\$B\/(\S+)\/\$f/)[1]; nombres.forEach(nm => { if (!fs.existsSync(path.join(raiz, p.carpeta, dir, nm + '.json'))) falla(`${c} descarga ${dir}/${nm}.json y no existe en su carpeta`); }); });
  (s.match(/-o \S+ \$B\/\S+/g) || []).map(x => x.split('$B/')[1]).filter(x => !x.includes('<') && !x.includes('$') && !x.includes(';')).forEach(rel => { if (!fs.existsSync(path.join(raiz, p.carpeta, rel))) falla(`${c} descarga ${rel} y no existe en su carpeta`); });
  (s.match(/-o \S+ \$R\/\S+/g) || []).map(x => x.split('$R/')[1]).filter(x => !x.includes('<')).forEach(rel => { if (!fs.existsSync(path.join(raiz, rel))) falla(`${c} descarga ${rel} y no existe en el repositorio`); });
  if (!E.some(e => e.startsWith(c))) bien(`${c} aislado y con fronteras claras; todos los archivos que cita existen`);
}
// --- manifest.json de cada prompt: el mapa debe ser completo y verdadero ---
const todosLosArchivos = (dir) => fs.readdirSync(path.join(raiz, dir), { withFileTypes: true }).flatMap(n => {
  if (n.name === 'node_modules' || n.name === 'salida') return [];
  const r = dir + '/' + n.name; return n.isDirectory() ? todosLosArchivos(r) : [r];
});
for (const p of idx.prompts) {
  const c = `[${p.id}]`, mp = path.join(p.carpeta, 'manifest.json');
  if (!fs.existsSync(path.join(raiz, mp))) { falla(`${c} falta ${mp}`); continue; }
  let m; try { m = JSON.parse(leer(mp)); } catch (e) { falla(`${c} manifest.json inválido: ${e.message}`); continue; }
  if (m.id !== p.id) falla(`${c} el manifest dice id «${m.id}»`);
  const cubiertos = [...m.leer_primero, ...m.leer_segun_necesidad.map(x => x.archivo), ...m.ejecutar, ...m.pruebas_y_mantenimiento];
  [...m.leer_primero, ...m.ejecutar, ...m.pruebas_y_mantenimiento, ...m.leer_segun_necesidad.map(x => x.archivo).filter(f => !f.includes('<'))]
    .forEach(f => { if (!fs.existsSync(path.join(raiz, p.carpeta, f))) falla(`${c} el manifest lista ${f} y no existe`); });
  (m.carpetas_de_apoyo || []).forEach(d => { if (!fs.existsSync(path.join(raiz, p.carpeta, d))) falla(`${c} el manifest lista la carpeta ${d} y no existe`); });
  (m.datos_compartidos || []).forEach(f => { if (!fs.existsSync(path.join(raiz, f))) falla(`${c} dato compartido ${f} no existe`); });
  todosLosArchivos(p.carpeta).map(f => f.slice(p.carpeta.length + 1)).forEach(f => {
    const ok_ = cubiertos.includes(f) || (m.carpetas_de_apoyo || []).some(d => f.startsWith(d));
    if (!ok_) falla(`${c} el archivo ${f} no está en manifest.json (agrégalo para que Claude sepa que existe)`);
  });
  const sk = leer(p.skill);
  if (!/preparar-prompt\.sh/.test(sk)) falla(`${c} el SKILL.md no usa herramientas/preparar-prompt.sh`);
  const vSkill = (sk.match(/— v(\d+\.\d+)/) || [])[1];
  if (vSkill && !m.version.startsWith(vSkill)) falla(`${c} el SKILL.md dice v${vSkill} y el manifest v${m.version}`);
  if (!E.some(e => e.startsWith(c))) bien(`${c} manifest.json completo: todos los archivos están mapeados y existen`);
}
Object.keys(idx.recursos_compartidos || {}).forEach(r => { if (!fs.existsSync(path.join(raiz, r))) falla(`El recurso compartido ${r} no existe`); });
const md = leer('INDICE-PARA-CLAUDE.md');
(md.match(/\]\(([^)#]+)\)/g) || []).map(x => x.slice(2, -1)).filter(x => !/^https?:/.test(x)).forEach(rel => { if (!fs.existsSync(path.join(raiz, rel))) falla(`INDICE-PARA-CLAUDE.md enlaza a un archivo que no existe: ${rel}`); });
const rd = leer('README.md'); if (!rd.includes('INDICE-PARA-CLAUDE.md')) falla('El README no enlaza a INDICE-PARA-CLAUDE.md');
idx.prompts.forEach(p => { if (!rd.includes(p.skill_url_github)) falla(`El README no trae el enlace de [${p.id}]`); });
ok.forEach(m => console.log('✅ ' + m));
if (E.length) { console.error(E.map(e => '❌ ' + e).join('\n')); process.exit(1); }
console.log(`✅ Índice correcto: ${idx.prompts.length} prompts independientes.`);
