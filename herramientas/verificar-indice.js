#!/usr/bin/env node
/* verificar-indice.js — Comprueba que los prompts del repositorio estén aislados y bien enrutados.
 * Uso: node herramientas/verificar-indice.js           (desde cualquier carpeta)
 * Verifica: cada prompt del INDICE.json existe, tiene nombre y descripción con su límite («NO sirve para»), un bloque
 * «ALCANCE Y FRONTERAS», descarga archivos SOLO de su propia carpeta y las URL del índice apuntan a su SKILL.md. */
const fs = require('fs'), path = require('path');
const raiz = path.join(__dirname, '..'), E = [], ok = [];
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
  if (!E.some(e => e.startsWith(c))) bien(`${c} aislado y con fronteras claras`);
}
Object.keys(idx.recursos_compartidos || {}).forEach(r => { if (!fs.existsSync(path.join(raiz, r))) falla(`El recurso compartido ${r} no existe`); });
const md = leer('INDICE-PARA-CLAUDE.md');
(md.match(/\]\(([^)#]+)\)/g) || []).map(x => x.slice(2, -1)).filter(x => !/^https?:/.test(x)).forEach(rel => { if (!fs.existsSync(path.join(raiz, rel))) falla(`INDICE-PARA-CLAUDE.md enlaza a un archivo que no existe: ${rel}`); });
const rd = leer('README.md'); if (!rd.includes('INDICE-PARA-CLAUDE.md')) falla('El README no enlaza a INDICE-PARA-CLAUDE.md');
idx.prompts.forEach(p => { if (!rd.includes(p.skill_url_github)) falla(`El README no trae el enlace de [${p.id}]`); });
ok.forEach(m => console.log('✅ ' + m));
if (E.length) { console.error(E.map(e => '❌ ' + e).join('\n')); process.exit(1); }
console.log(`✅ Índice correcto: ${idx.prompts.length} prompts independientes.`);
